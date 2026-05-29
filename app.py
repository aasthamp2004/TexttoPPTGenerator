import base64
import copy
import ast
import json
import re
from uuid import uuid4
from openai import AzureOpenAI
import requests
import streamlit as st
from backend.config import (
    AZURE_KEY, AZURE_ENDPOINT, AZURE_API_VERSION, AZURE_DEPLOYMENT
)

_client = AzureOpenAI(
    api_key=AZURE_KEY,
    api_version=AZURE_API_VERSION,
    azure_endpoint=AZURE_ENDPOINT,
)

OUTLINE_URL = "http://127.0.0.1:8080/generate-outline"
BUILD_URL   = "http://127.0.0.1:8080/build-ppt"

LAYOUT_OPTIONS = [
    "title_cover","section_index","bullets","two_column","big_stat",
    "timeline","icon_grid","case_study","table","chart",
    "image_text_split","hybrid_insight",
]

st.set_page_config(page_title="AI PPT Generator", page_icon="💬", layout="wide")

# ─── CUSTOM CSS ───────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stChatMessage { border-radius: 12px; margin-bottom: 8px; }
    .streamlit-expanderHeader { font-weight: 600; }
    .download-section {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        border-radius: 16px;
        padding: 24px;
        text-align: center;
        margin: 24px 0;
    }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
#  STATE
# ═══════════════════════════════════════════════════════════════════════════

def init_state():
    defaults = {
        "messages": [{
            "role": "assistant",
            "content": (
                "👋 Tell me the presentation topic and I'll draft slide sections.\n\n"
                "After the outline is ready, you can:\n"
                "- **Edit slides directly** in the editor below\n"
                "- **Ask me in chat** to make changes, e.g. *'Edit slide 2, change title to Market Analysis'* "
                "or *'Add bullet to slide 3: Revenue grew 40% YoY'* or *'Remove slide 4'*"
            ),
        }],
        "outline_payload": None,
        "ppt_bytes":       None,
        "ppt_filename":    None,
        "topic":           "",
        "tone":            "Professional",
        "num_slides":      10,
        "logo_bytes":      None,
        "logo_name":       None,
        "content_image_bytes": None,
        "content_image_name": None,
        "token_usage":     {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0},
        "last_usage":      None,
        "session_id":      uuid4().hex,
        "pending_outline_context": None,
        "deck_history":    [],
        "current_deck_id": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def add_message(role, content):
    st.session_state.messages.append({"role": role, "content": content})


def add_deck_message(deck_id: str):
    if not deck_id:
        return
    st.session_state.messages.append({
        "role": "assistant",
        "content": "",
        "kind": "deck",
        "deck_id": deck_id,
    })

def usage_to_dict(usage):
    if not usage:
        return None
    return {
        "prompt_tokens": int(getattr(usage, "prompt_tokens", 0) or 0),
        "completion_tokens": int(getattr(usage, "completion_tokens", 0) or 0),
        "total_tokens": int(getattr(usage, "total_tokens", 0) or 0),
    }

def record_usage(kind: str, usage: dict):
    if not usage:
        return
    totals = st.session_state.token_usage
    for k in ("prompt_tokens", "completion_tokens", "total_tokens"):
        totals[k] += int(usage.get(k, 0) or 0)
    st.session_state.last_usage = {"kind": kind, **usage}


def save_logo_upload(files) -> str | None:
    if not files:
        return None
    file = files[0]
    name = getattr(file, "name", "logo.png")
    st.session_state.logo_bytes = file.getvalue()
    st.session_state.logo_name = name
    return name


def save_content_image_upload(files) -> str | None:
    if not files:
        return None
    file = files[0]
    name = getattr(file, "name", "image.png")
    st.session_state.content_image_bytes = file.getvalue()
    st.session_state.content_image_name = name
    return name


def should_treat_upload_as_logo(prompt: str, pending_outline_context=None) -> bool:
    if pending_outline_context and pending_outline_context.get("awaiting_logo"):
        return True
    text = (prompt or "").strip().lower()
    return any(phrase in text for phrase in (
        "logo", "brand logo", "company logo", "use as logo", "for branding", "branding"
    ))

def normalize_slide(slide: dict):
    slide = ensure_editor_id(slide or {})
    slide.setdefault("title", "New Slide")
    slide.setdefault("subtitle", "")
    slide.setdefault("layout", "bullets")
    slide.setdefault("icon", "▸")
    slide.setdefault("content", [])
    slide.setdefault("style", {})
    if not isinstance(slide.get("content"), list):
        slide["content"] = []
    return slide

def draft_slide_from_request(user_text: str, slides: list):
    titles = [str(s.get("title","")).strip() for s in (slides or []) if isinstance(s, dict)]
    prompt = f"""
Create a single new slide for a presentation.
User request: {user_text}
Existing slide titles: {titles}

Return ONLY valid JSON for a single slide object with fields:
title, subtitle, layout, icon, content (list), style (dict).
Choose the most appropriate layout.
"""
    resp = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.4,
    )
    raw = (resp.choices[0].message.content or "").strip()
    if not raw.startswith("{"):
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        raw = m.group() if m else raw
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("Invalid slide JSON")
    data.setdefault("title", "New Slide")
    data.setdefault("subtitle", "")
    data.setdefault("layout", "bullets")
    data.setdefault("icon", "▸")
    data.setdefault("content", ["Add your main point here."])
    data.setdefault("style", {})
    return data

def summarize_slides_for_llm(slides: list) -> list:
    summary = []
    for i, slide in enumerate(slides or [], start=1):
        if not isinstance(slide, dict):
            continue
        content = slide.get("content", []) or []
        if not isinstance(content, list):
            content = []
        summary.append({
            "slide_number": i,
            "title": str(slide.get("title", "")).strip(),
            "subtitle": str(slide.get("subtitle", "")).strip(),
            "layout": str(slide.get("layout", "bullets")).strip(),
            "content_preview": [str(x).strip() for x in content[:3] if str(x).strip()],
        })
    return summary

def needs_clarification_for_edit(user_text: str, slides: list):
    titles = [str(s.get("title","")).strip() for s in (slides or []) if isinstance(s, dict)]
    slide_summary = summarize_slides_for_llm(slides)
    recent_question = ""
    pending = st.session_state.get("pending_clarification") or {}
    if isinstance(pending, dict):
        recent_question = str(pending.get("question", "")).strip()
    prompt = f"""
You are checking if a user's edit request specifies a target slide clearly.
User request: {user_text}
Slide titles: {titles}
Current deck summary:
{json.dumps(slide_summary, indent=2)}
Recent clarification question to avoid repeating:
{recent_question or "None"}

Respond ONLY valid JSON:
{{ "needs_clarification": true/false, "question": "..." }}

Rules:
- Ask a clarification only when execution would otherwise risk editing the wrong slide or making up a factual/user-specific choice that cannot be inferred safely.
- Good reasons to clarify:
  1. Multiple slides match the request and no clear target is given.
  2. The request depends on a missing user-specific fact that should not be invented.
  3. The request conflicts with the current deck structure in a way that blocks execution.
- Do NOT ask for clarification for open-ended creative decisions the model can make reasonably on its own.
- Do NOT ask for clarification for things like:
  - how the slide should look visually
  - what detailed bullets/examples to add
  - how to expand or shorten content
  - choosing an appropriate layout, wording, or supporting detail
- If the user intent is clear enough, let the LLM decide and execute.
- The question must be specific and relevant to the ambiguity. Mention likely slide numbers/titles when helpful.
- Ask only one concise question.
- If the request is vague about both the target slide and the change, ask one compact question covering both.
- Do not repeat the same wording as the recent clarification question. Rephrase it naturally.
"""
    try:
        resp = _client.chat.completions.create(
            model=AZURE_DEPLOYMENT,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
        )
        raw = (resp.choices[0].message.content or "").strip()
        if not raw.startswith("{"):
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            raw = m.group() if m else raw
        data = json.loads(raw)
        if isinstance(data, dict) and data.get("needs_clarification"):
            return True, data.get("question") or "Which slide would you like to edit?"
    except Exception:
        pass
    cleaned = (user_text or "").strip().lower()
    slide_num = extract_slide_number(cleaned)
    generic_edit_patterns = [
        r"^\s*edit\s+slide\s+\d+\s*$",
        r"^\s*change\s+slide\s+\d+\s*$",
        r"^\s*update\s+slide\s+\d+\s*$",
        r"^\s*modify\s+slide\s+\d+\s*$",
        r"^\s*fix\s+slide\s+\d+\s*$",
        r"^\s*edit\s+\d+\s*$",
        r"^\s*change\s+\d+\s*$",
        r"^\s*update\s+\d+\s*$",
    ]
    if slide_num and any(re.fullmatch(pattern, cleaned) for pattern in generic_edit_patterns):
        title = ""
        if 1 <= slide_num <= len(slides):
            title = str((slides[slide_num - 1] or {}).get("title", "")).strip()
        if title:
            return True, f"What would you like me to change on slide {slide_num} ({title})?"
        return True, f"What would you like me to change on slide {slide_num}?"
    if cleaned in {"edit", "change", "update", "modify", "fix this", "change this", "edit this"}:
        if titles:
            examples = ", ".join(
                f"slide {i}: {title}" for i, title in enumerate(titles[:3], start=1) if title
            )
            if examples:
                return True, f"Which slide should I update, and what would you like changed? For example: {examples}."
        return True, "Which slide should I update, and what would you like changed?"
    return False, None

def render_clarification(question: str, intent_text: str = None):
    with st.chat_message("assistant"):
        st.markdown("**Clarification Needed**")
        st.markdown(question)
    add_message("assistant", f"Clarification Needed: {question}")
    if intent_text:
        st.session_state["pending_clarification"] = {
            "intent": intent_text,
            "question": question,
        }


def move_current_deck_to_chat_end():
    current_deck_id = st.session_state.get("current_deck_id")
    if not current_deck_id:
        return
    messages = st.session_state.get("messages", [])
    filtered = [
        msg for msg in messages
        if not (msg.get("kind") == "deck" and msg.get("deck_id") == current_deck_id)
    ]
    filtered.append({
        "role": "assistant",
        "content": "",
        "kind": "deck",
        "deck_id": current_deck_id,
    })
    st.session_state.messages = filtered

def slides_changed(old, new):
    if len(old) != len(new):
        return True
    for o, n in zip(old, new):
        if o.get("title") != n.get("title"):
            return True
        if o.get("content") != n.get("content"):
            return True
        if o.get("layout") != n.get("layout"):
            return True
        if o.get("stat") != n.get("stat"):
            return True
    return False


def reset_editor_widget_state():
    keys = [k for k in st.session_state.keys() if re.match(r"^[tslc]_[0-9a-f]{32}$", k)]
    for key in keys:
        st.session_state.pop(key, None)


def prime_editor_widget_state(slide: dict, force: bool = False):
    sid = slide.get("_editor_id")
    if not sid:
        return
    values = {
        f"t_{sid}": slide.get("title", ""),
        f"s_{sid}": slide.get("subtitle", ""),
        f"l_{sid}": slide.get("layout", "bullets"),
        f"c_{sid}": get_editor_content(slide),
    }
    for key, value in values.items():
        if force or key not in st.session_state:
            st.session_state[key] = value


def sync_all_editor_widgets(slides: list):
    reset_editor_widget_state()
    for slide in slides:
        if isinstance(slide, dict):
            prime_editor_widget_state(slide, force=True)


def archive_current_deck():
    outline = st.session_state.get("outline_payload")
    ppt_bytes = st.session_state.get("ppt_bytes")
    if not outline or not ppt_bytes:
        return
    deck_id = st.session_state.get("current_deck_id") or uuid4().hex
    archived = {
        "deck_id": deck_id,
        "topic": st.session_state.get("topic", ""),
        "tone": st.session_state.get("tone", "Professional"),
        "num_slides": st.session_state.get("num_slides", 10),
        "outline_payload": copy.deepcopy(outline),
        "ppt_bytes": ppt_bytes,
        "ppt_filename": st.session_state.get("ppt_filename", "presentation.pptx"),
    }
    history = st.session_state.get("deck_history", [])
    history = [deck for deck in history if deck.get("deck_id") != deck_id]
    history.append(archived)
    st.session_state.deck_history = history


def reset_current_deck(keep_messages: bool = True, preserve_logo: bool = False):
    if not keep_messages:
        st.session_state["messages"] = [{
            "role": "assistant",
            "content": (
                "👋 Tell me the presentation topic and I'll draft slide sections.\n\n"
                "After the outline is ready, you can:\n"
                "- **Edit slides directly** in the editor below\n"
                "- **Ask me in chat** to make changes, e.g. *'Edit slide 2, change title to Market Analysis'* "
                "or *'Add bullet to slide 3: Revenue grew 40% YoY'* or *'Remove slide 4'*"
            ),
        }]
    for k in ("outline_payload", "ppt_bytes", "ppt_filename", "topic", "logo_bytes", "logo_name", "content_image_bytes", "content_image_name", "pending_outline_context", "current_deck_id"):
        if preserve_logo and k in ("logo_bytes", "logo_name"):
            continue
        st.session_state[k] = None if k in ("outline_payload", "ppt_bytes", "ppt_filename", "logo_bytes", "logo_name", "content_image_bytes", "content_image_name", "current_deck_id") else ""
    st.session_state["tone"] = "Professional"
    st.session_state["num_slides"] = 10
    reset_editor_widget_state()


def is_new_presentation_request(prompt: str, has_outline: bool) -> bool:
    text = (prompt or "").strip().lower()
    if not text:
        return False
    # Basic smalltalk guard: avoid treating greetings/thanks as a deck request
    if not has_outline and (is_greeting(text) or is_smalltalk(text)):
        return False
    if not has_outline:
        return True
    if text.startswith(("new ppt", "new deck", "new presentation", "create a new", "generate a new", "start a new")):
        return True
    if text.startswith(("create ", "generate ", "build ", "make ")) and any(
        phrase in text for phrase in ("ppt", "deck", "presentation", "slides")
    ):
        return True
    # If a deck already exists and the user clearly asks for another presentation topic,
    # treat it as a new PPT even without the word "new".
    if has_outline and any(
        phrase in text
        for phrase in (
            "ppt on", "ppt about", "ppt for",
            "presentation on", "presentation about", "presentation for",
            "deck on", "deck about", "deck for",
            "slides on", "slides about", "slides for",
        )
    ):
        return True
    return False

def is_greeting(text: str) -> bool:
    cleaned = re.sub(r"[^\w\s]", "", (text or "").strip().lower())
    if not cleaned:
        return False
    greetings = {
        "hi", "hello", "hey", "hiya", "yo", "sup", "howdy",
        "good morning", "good afternoon", "good evening","hyy","heyy"
    }
    if cleaned in greetings:
        return True
    # Allow short greeting + name (e.g., "hi there", "hello team")
    tokens = cleaned.split()
    if len(tokens) <= 3 and tokens[0] in {"hi", "hello", "hey", "hiya", "yo", "howdy"}:
        return True
    return False

def is_smalltalk(text: str) -> bool:
    cleaned = re.sub(r"[^\w\s]", "", (text or "").strip().lower())
    if not cleaned:
        return False
    phrases = {
        "thanks", "thank you", "thx", "ty",
        "ok", "okay", "cool", "nice", "great", "awesome",
        "good", "fine", "sure", "yep", "yup", "yes", "no",
        "got it", "sounds good",
    }
    if cleaned in phrases:
        return True
    tokens = cleaned.split()
    if len(tokens) <= 3 and tokens[0] in {"thanks", "thank", "ok", "okay", "cool", "great"}:
        return True
    return False


def topic_from_prompt(prompt: str) -> str:
    text = (prompt or "").strip()
    patterns = [
        r"^(?:new|create|generate|build|make|start)\s+(?:a\s+)?(?:new\s+)?(?:ppt|deck|presentation|slides?)\s+(?:on|about|for)\s+",
        r"^(?:new|create|generate|build|make|start)\s+(?:a\s+)?(?:new\s+)?(?:ppt|deck|presentation|slides?)\s+",
    ]
    for pattern in patterns:
        cleaned = re.sub(pattern, "", text, flags=re.IGNORECASE).strip(" :.-")
        if cleaned and cleaned != text:
            return cleaned
    return text


def extract_num_slides(text: str):
    value = str(text or "").strip().lower()
    if not value:
        return None
    if re.fullmatch(r"\d{1,2}", value):
        try:
            return max(3, min(20, int(value)))
        except Exception:
            pass
    word_to_num = {
        "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
        "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
        "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16,
        "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20,
    }
    if value in word_to_num:
        return word_to_num[value]
    patterns = [
        r"\b(?:make|create|generate|build)\s+(?:me\s+|a\s+)?(\d+)\s+slides?\b",
        r"(\d+)\s*-\s*slide",
        r"(\d+)\s+slides?",
        r"slides?\s*[:=]?\s*(\d+)",
        r"deck of (\d+)",
        r"presentation of (\d+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, value)
        if match:
            try:
                count = int(match.group(1))
                return max(3, min(20, count))
            except Exception:
                pass
    return None


def should_skip_logo_step(text: str) -> bool:
    cleaned = str(text or "").strip().lower()
    if not cleaned:
        return False
    exact = {
        "skip", "no", "nope", "continue", "go ahead", "without logo",
        "no logo", "proceed", "dont use logo", "don't use logo",
        "no thanks", "not now", "without a logo", "no need",
    }
    if cleaned in exact:
        return True
    patterns = [
        r"\bno\b.*\blogo\b",
        r"\bwithout\b.*\blogo\b",
        r"\bdon'?t\b.*\buse\b.*\blogo\b",
        r"\bskip\b",
        r"\bcontinue\b",
        r"\bproceed\b",
        r"\bgo ahead\b",
        r"\bnot now\b",
    ]
    return any(re.search(pattern, cleaned) for pattern in patterns)

def apply_add_action(old_slides, slide, position):
    new_slide = ensure_editor_id(slide)
    new_slide["_editor_id"] = uuid4().hex
    slides = list(old_slides)
    if isinstance(position, int):
        idx = max(1, min(len(slides)+1, position))
        slides.insert(idx-1, new_slide)
        return slides
    pos = str(position or "").strip().lower()
    if pos in ("end","last","bottom"):
        slides.append(new_slide)
        return slides
    if pos in ("start","first","top"):
        slides.insert(0, new_slide)
        return slides
    slides.append(new_slide)
    return slides

def extract_slide_number(text: str):
    m = re.search(r"\bslide\s+(\d+)\b", (text or "").lower())
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None

def is_thank_you_request(text: str) -> bool:
    cleaned = (text or "").lower()
    return "thank you" in cleaned or re.search(r"\bthanks\b", cleaned)

def is_convert_request(text: str) -> bool:
    cleaned = (text or "").lower()
    return any(k in cleaned for k in ("convert", "change", "replace", "make", "turn"))

def make_thank_you_slide(slide: dict) -> dict:
    slide = ensure_editor_id(slide)
    slide["title"] = "Thank You"
    slide["subtitle"] = "Questions?"
    slide["layout"] = "title_cover"
    slide["content"] = []
    return slide

def refresh_section_index_slide(slides: list) -> list:
    normalized = [ensure_editor_id(s) for s in (slides or []) if isinstance(s, dict)]
    if len(normalized) < 2:
        return normalized
    index_slide = normalized[1]
    if str(index_slide.get("layout", "")).strip() != "section_index":
        return normalized
    sections = [
        str(s.get("title", "")).strip()
        for s in normalized[2:]
        if isinstance(s, dict) and str(s.get("title", "")).strip()
    ][:6]
    index_slide["sections"] = sections
    index_slide["content"] = sections
    normalized[1] = index_slide
    return normalized


def _image_caption_from_name(name: str) -> str:
    base = re.sub(r"\.[A-Za-z0-9]+$", "", str(name or "")).replace("_", " ").replace("-", " ").strip()
    return base[:60] or "Reference Image"


def attach_content_image_to_outline(outline: dict):
    if not isinstance(outline, dict) or not st.session_state.get("content_image_bytes"):
        return outline
    slides = [ensure_editor_id(s) for s in outline.get("slides", []) if isinstance(s, dict)]
    if not slides:
        return outline

    target_index = None
    for idx in range(2, len(slides)):
        if str(slides[idx].get("layout", "")).strip() == "image_text_split":
            target_index = idx
            break
    if target_index is None:
        for idx in range(2, len(slides)):
            layout = str(slides[idx].get("layout", "")).strip()
            if layout in ("bullets", "two_column", "big_stat", "case_study", "chart", "hybrid_insight", "icon_grid"):
                target_index = idx
                break
    if target_index is None:
        target_index = 2 if len(slides) > 2 else len(slides) - 1
    if target_index < 0:
        return outline

    slide = ensure_editor_id(slides[target_index])
    if str(slide.get("layout", "")).strip() != "image_text_split":
        for fld in ("steps","grid_items","left_points","right_points","table_columns","table_rows","chart_data","metrics"):
            slide.pop(fld, None)
        slide["layout"] = "image_text_split"
        content = slide.get("content", []) or []
        if not isinstance(content, list):
            content = []
        slide["content"] = content[:5] or [
            f"Visual reference related to {st.session_state.get('topic') or slide.get('title') or 'the topic'}.",
            "Use the image to support the main narrative on this slide.",
        ]
    slide["image_caption"] = _image_caption_from_name(st.session_state.get("content_image_name", ""))
    slide["image_side"] = "right"
    slide["use_uploaded_image"] = True
    slides[target_index] = slide
    outline["slides"] = refresh_section_index_slide(slides)
    return outline

# ═══════════════════════════════════════════════════════════════════════════
#  AUTO-REBUILD PPT
# ═══════════════════════════════════════════════════════════════════════════

def rebuild_ppt_from_outline():
    """Rebuild PPT from current outline_payload and store bytes in session."""
    outline = st.session_state.outline_payload
    if not outline or not outline.get("slides"):
        return None

    build_payload = sanitize_outline_for_build(copy.deepcopy(outline))
    files = None
    if st.session_state.get("logo_bytes") or st.session_state.get("content_image_bytes"):
        files = {}
        if st.session_state.get("logo_bytes"):
            files["logo"] = (
                st.session_state.get("logo_name", "logo.png"),
                st.session_state.logo_bytes,
                "image/png",
            )
        if st.session_state.get("content_image_bytes"):
            files["content_image"] = (
                st.session_state.get("content_image_name", "image.png"),
                st.session_state.content_image_bytes,
                "image/png",
            )
    try:
        resp = requests.post(
            BUILD_URL,
            data={
                "topic": st.session_state.topic or "Presentation",
                "tone": st.session_state.tone,
                "slides_json": json.dumps(build_payload),
            },
            files=files,
            timeout=(10, 180),
        )
        if resp.status_code == 200:
            result = resp.json()
            ppt_bytes = base64.b64decode(result["ppt_base64"])
            st.session_state.ppt_bytes = ppt_bytes
            slides = outline.get("slides", [])
            first_title = ""
            if isinstance(slides, list) and slides:
                first = slides[0] if isinstance(slides[0], dict) else {}
                first_title = str(first.get("title", "")).strip()
            file_topic = first_title or st.session_state.topic or "presentation"
            safe_name = re.sub(r"[^\w\-]+", "_", file_topic).strip("_") or "presentation"
            st.session_state.ppt_filename = f"{safe_name[:80]}.pptx"
            return ppt_bytes
        else:
            st.error(f"PPT rebuild failed: {resp.text}")
            return None
    except Exception as e:
        st.error(f"Rebuild error: {e}")
        return None


# ═══════════════════════════════════════════════════════════════════════════
#  SLIDE UTILS
# ═══════════════════════════════════════════════════════════════════════════

def ensure_editor_id(slide: dict) -> dict:
    slide = dict(slide or {})
    slide.setdefault("_editor_id",  uuid4().hex)
    slide.setdefault("title",       "Untitled Slide")
    slide.setdefault("subtitle",    "")
    slide.setdefault("layout",      "bullets")
    slide.setdefault("icon",        "▸")
    slide.setdefault("content",     [])
    slide.setdefault("style",       {})
    if not isinstance(slide["content"], list): slide["content"] = []
    slide["content"] = [str(x).strip() for x in slide["content"] if str(x).strip()]
    if not isinstance(slide["style"], dict):   slide["style"] = {}
    if slide.get("layout") == "icon_grid":
        slide["grid_items"] = normalize_icon_grid_items(slide, bullets_to_text(slide["content"]))
        if slide["grid_items"]:
            slide["content"] = [
                f"{item['title']}: {item['detail']}".rstrip(": ").strip()
                for item in slide["grid_items"]
            ]
    return slide

def bullets_to_text(items) -> str:
    if not isinstance(items, list): return ""
    return "\n".join(str(x).strip() for x in items if str(x).strip())

def text_to_bullets(value: str) -> list:
    bullets = []
    for line in str(value or "").splitlines():
        c = line.strip().lstrip("-").lstrip("•").strip()
        if c: bullets.append(c)
    return bullets


def seq_icon(idx: int) -> str:
    if 0 <= idx < 26:
        return chr(ord("A") + idx)
    return str(idx + 1)


def normalize_icon_grid_items(slide: dict, text: str = "") -> list:
    items = slide.get("grid_items", []) if isinstance(slide, dict) else []
    cleaned = []

    if isinstance(items, list):
        for i, item in enumerate(items):
            if not isinstance(item, dict):
                continue
            title = str(item.get("title", "")).strip()
            detail = str(item.get("detail", item.get("description", ""))).strip()
            if not title:
                continue
            icon = str(item.get("icon", "")).strip() or seq_icon(i)
            cleaned.append({"icon": icon, "title": title, "detail": detail})
    if cleaned:
        return cleaned[:4]

    source = str(text or "").strip()
    parsed_items = None
    if source:
        try:
            parsed = json.loads(source)
            if isinstance(parsed, dict) and isinstance(parsed.get("grid_items"), list):
                parsed_items = parsed.get("grid_items")
            elif isinstance(parsed, list):
                parsed_items = parsed
        except Exception:
            try:
                parsed = ast.literal_eval(source)
                if isinstance(parsed, dict) and isinstance(parsed.get("grid_items"), list):
                    parsed_items = parsed.get("grid_items")
                elif isinstance(parsed, list):
                    parsed_items = parsed
            except Exception:
                parsed_items = None

    if isinstance(parsed_items, list):
        for i, item in enumerate(parsed_items):
            if not isinstance(item, dict):
                continue
            title = str(item.get("title", "")).strip()
            detail = str(item.get("detail", item.get("description", ""))).strip()
            if not title:
                continue
            icon = str(item.get("icon", "")).strip() or seq_icon(i)
            cleaned.append({"icon": icon, "title": title, "detail": detail})
        if cleaned:
            return cleaned[:4]

    for i, line in enumerate(text_to_bullets(source)):
        if len(cleaned) >= 4:
            break
        parsed_line = None
        try:
            parsed_line = json.loads(line)
        except Exception:
            try:
                parsed_line = ast.literal_eval(line)
            except Exception:
                parsed_line = None
        if isinstance(parsed_line, dict):
            title = str(parsed_line.get("title", "")).strip()
            detail = str(parsed_line.get("detail", parsed_line.get("description", ""))).strip()
            icon = str(parsed_line.get("icon", "")).strip() or seq_icon(i)
            if title:
                cleaned.append({"icon": icon, "title": title, "detail": detail})
                continue
        if ":" in line:
            title, detail = [part.strip() for part in line.split(":", 1)]
        else:
            title, detail = line[:30].strip(), line.strip()
        if title:
            cleaned.append({"icon": seq_icon(i), "title": title, "detail": detail})
    return cleaned[:4]

def summarize_outline(slides: list) -> str:
    if not slides:
        return "I couldn't generate an outline yet. Try another prompt."
    lines = ["✅ I drafted these sections for the deck:\n"]
    for i, s in enumerate(slides, 1):
        lines.append(f"**{i}.** {s.get('title', f'Slide {i}')}")
    lines.append("\nEdit anything in the **Edit Sections** panel below, or ask me to make specific changes in chat.")
    return "\n".join(lines)

# ═══════════════════════════════════════════════════════════════════════════
#  LLM EDIT (FULL SLIDES RETURN)
# ═══════════════════════════════════════════════════════════════════════════

def interpret_edit_with_llm(user_text, slides):
    prompt = f"""
You are an AI presentation editor.

User request:
{user_text}

Current slides (FULL DECK):
{json.dumps(slides, indent=2)}

Your job:
- Understand the user's intent in natural language
- Modify ONLY the relevant slide(s)
- **Return the ENTIRE updated slides array (all slides, not just the changed one)**

DATA RULES (IMPORTANT):
- Each slide has fields like: title, subtitle, content, layout, stat
- If a slide uses "big_stat" layout:
    → "stat" field MUST be updated if user mentions a number or percentage
- If user asks for more/less content:
    → modify "content" array size
- NEVER ignore numerical changes (like %, numbers, metrics)
- Preserve all other data

IMPORTANT:
- Do NOT change unrelated slides
- Do NOT change structure randomly
- Apply changes exactly where needed
- If the user mentions a specific slide number or title, you MUST use action "edit" (never "add")
- Use action "add" ONLY if the user explicitly says add/insert/new slide
- If the user says "convert/change/replace/make slide X ..." treat it as an edit to that slide

Return ONLY valid JSON in ONE of these forms:

1) Clarification required:
{{
  "action": "clarify",
  "question": "Which slide would you like to edit? (e.g., slide 3 or the slide titled 'Market Analysis')"
}}

2) Add slide:
{{
  "action": "add",
  "position": "end" or 3,
  "slide": {{ ... full slide object with title, subtitle, layout, content, style ... }}
}}

3) Remove slide:
{{
  "action": "remove",
  "targets": [1, 4]
}}

4) Edit applied:
{{
  "action": "edit",
  "targets": [1, 3],
  "slides": [ ... FULL updated slides array ... ]
}}
"""
    response = _client.chat.completions.create(
        model=AZURE_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )

    usage = usage_to_dict(response.usage)

    raw = response.choices[0].message.content
    try:
        return json.loads(raw), usage
    except:
        import re
        match = re.search(r'\{.*\}', raw, re.DOTALL)
        if match:
            return json.loads(match.group()), usage
        raise ValueError("Invalid JSON from LLM")

# ═══════════════════════════════════════════════════════════════════════════
#  LAYOUT-AWARE EDITOR (unchanged from your original)
# ═══════════════════════════════════════════════════════════════════════════

def get_editor_content(slide: dict) -> str:
    layout = slide.get("layout","bullets")
    if layout == "two_column":
        lp = slide.get("left_points",[]) or []
        rp = slide.get("right_points",[]) or []
        combined = list(lp) + list(rp)
        if combined: return bullets_to_text(combined)
    elif layout == "timeline":
        steps = slide.get("steps",[]) or []
        if steps:
            return "\n".join(
                f"{s.get('label',''): <20}{s.get('detail','')}"
                for s in steps if isinstance(s, dict)
            )
    elif layout == "icon_grid":
        items = slide.get("grid_items",[]) or []
        if items:
            return "\n".join(
                f"{g.get('title','')}: {g.get('detail','')}"
                for g in items if isinstance(g, dict)
            )
    elif layout == "table":
        cols = slide.get("table_columns",[]) or []
        rows = slide.get("table_rows",[]) or []
        if cols:
            lines = [" | ".join(str(c) for c in cols)]
            for row in rows:
                lines.append(" | ".join(str(v) for v in row))
            return "\n".join(lines)
    elif layout == "chart":
        data = slide.get("chart_data",[]) or []
        if data:
            return "\n".join(
                f"{d.get('label','')} : {d.get('value','')}"
                for d in data if isinstance(d, dict)
            )
    elif layout == "case_study":
        lines = []
        if slide.get("company"):    lines.append(f"Company: {slide['company']}")
        if slide.get("result"):     lines.append(f"Result: {slide['result']}")
        for m in (slide.get("metrics") or []):
            if isinstance(m,dict): lines.append(f"  {m.get('label','')}: {m.get('value','')}")
        lines.extend(slide.get("content",[]) or [])
        return "\n".join(lines)
    elif layout in ("big_stat","hybrid_insight"):
        lines = []
        if slide.get("stat"):       lines.append(f"STAT: {slide['stat']}")
        if slide.get("stat_label"): lines.append(f"LABEL: {slide['stat_label']}")
        lines.extend(slide.get("content",[]) or [])
        return "\n".join(lines)
    elif layout == "section_index":
        secs = slide.get("sections",[]) or slide.get("content",[]) or []
        return bullets_to_text(secs)
    return bullets_to_text(slide.get("content",[]))

def save_editor_content(slide: dict, text: str) -> dict:
    layout = slide.get("layout","bullets")
    lines  = text_to_bullets(text)
    if layout == "two_column":
        mid = max(1, len(lines)//2)
        slide["left_points"]  = lines[:mid]
        slide["right_points"] = lines[mid:]
        slide["content"]      = lines
    elif layout == "timeline":
        steps = []
        for line in lines:
            if ":" in line:
                parts = line.split(":",1)
                steps.append({"label":parts[0].strip(),"detail":parts[1].strip()})
            elif len(line) > 25:
                steps.append({"label":line[:20].strip(),"detail":line[20:].strip()})
            else:
                steps.append({"label":line,"detail":""})
        slide["steps"]   = steps
        slide["content"] = lines
    elif layout == "icon_grid":
        slide["grid_items"] = normalize_icon_grid_items(slide, text)
        slide["content"] = [
            f"{item['title']}: {item['detail']}".rstrip(": ").strip()
            for item in slide["grid_items"]
        ]
    elif layout == "table":
        if lines:
            slide["table_columns"] = [c.strip() for c in lines[0].split("|") if c.strip()]
            slide["table_rows"]    = [
                [c.strip() for c in row.split("|")] for row in lines[1:]
            ]
        slide["content"] = lines
    elif layout == "chart":
        data = []
        for line in lines:
            if ":" in line:
                p = line.split(":",1)
                try: data.append({"label":p[0].strip(),"value":int(float(p[1].strip()))})
                except: data.append({"label":p[0].strip(),"value":50})
            else:
                data.append({"label":line,"value":50})
        slide["chart_data"] = data[:5]
        slide["content"]    = lines
    elif layout == "case_study":
        content_lines = []
        for line in lines:
            ll = line.lower()
            if ll.startswith("company:"):
                slide["company"] = line.split(":",1)[1].strip()
            elif ll.startswith("result:"):
                slide["result"] = line.split(":",1)[1].strip()
            else:
                content_lines.append(line)
        slide["content"] = content_lines
    elif layout in ("big_stat","hybrid_insight"):
        content_lines = []
        for line in lines:
            ll = line.lower()
            if ll.startswith("stat:"):
                slide["stat"] = line.split(":",1)[1].strip()
            elif ll.startswith("label:"):
                slide["stat_label"] = line.split(":",1)[1].strip()
            else:
                content_lines.append(line)
        slide["content"] = content_lines
    elif layout == "section_index":
        slide["sections"] = lines
        slide["content"]  = lines
    else:
        slide["content"] = lines
    return slide


def get_archived_deck(deck_id: str):
    for deck in st.session_state.get("deck_history", []):
        if deck.get("deck_id") == deck_id:
            return deck
    return None


def render_archived_deck(deck: dict, index_hint: int = 1):
    if not isinstance(deck, dict):
        return
    deck_outline = deck.get("outline_payload") or {}
    deck_slides = deck_outline.get("slides", []) if isinstance(deck_outline, dict) else []
    deck_label = deck.get("ppt_filename") or f"deck_{index_hint}.pptx"
    st.divider()
    st.markdown(f"### 📚 Previous Deck {index_hint}: {deck.get('topic') or deck_label}")
    st.caption(f"{len(deck_slides)} slides • {deck.get('tone', 'Professional')}")
    for idx, slide in enumerate(deck_slides, start=1):
        if not isinstance(slide, dict):
            continue
        title = str(slide.get("title", f"Slide {idx}")).strip() or f"Slide {idx}"
        layout = str(slide.get("layout", "bullets")).strip() or "bullets"
        with st.expander(f"**Slide {idx}** — {title}  `{layout}`", expanded=False):
            subtitle = str(slide.get("subtitle", "")).strip()
            if subtitle:
                st.caption(subtitle)
            content_text = get_editor_content(ensure_editor_id(slide))
            st.text_area(
                "Content / talking points",
                value=content_text,
                height=160,
                disabled=True,
                key=f"archived_{deck.get('deck_id','deck')}_{idx}",
            )
    if deck.get("ppt_bytes"):
        st.download_button(
            label=f"⬇️ Download {deck_label}",
            data=deck["ppt_bytes"],
            file_name=deck_label,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            key=f"download_archived_{deck.get('deck_id','deck')}",
        )


def render_current_deck():
    outline_payload = st.session_state.get("outline_payload")
    if not outline_payload:
        return

    st.divider()
    st.markdown("## 📝 Edit Sections")
    st.caption(
        "Edit headings, content, and layouts directly here — or ask me in chat above. "
        "Changes are saved automatically and the PPT is rebuilt instantly."
    )

    col_add, col_regen = st.columns([1, 1])
    with col_add:
        if st.button("➕ Add Section", use_container_width=True):
            n = len(outline_payload.get("slides",[])) + 1
            outline_payload["slides"].append(ensure_editor_id({
                "title": f"New Section {n}", "subtitle": "", "layout": "bullets",
                "content": ["Add your main point here.", "Supporting detail.", "Key takeaway."],
                "style": {},
            }))
            rebuild_ppt_from_outline()
            st.rerun()
    with col_regen:
        if st.button("🔄 Regenerate Outline", use_container_width=True):
            if not st.session_state.topic.strip():
                st.warning("Enter a presentation request in chat first.")
            else:
                with st.spinner("Refreshing outline..."):
                    try:
                        refreshed = request_outline(st.session_state.topic, st.session_state.num_slides, st.session_state.tone)
                        st.session_state.outline_payload = refreshed
                        sync_all_editor_widgets(refreshed.get("slides", []))
                        rebuild_ppt_from_outline()
                        reply = summarize_outline(refreshed.get("slides", []))
                        add_message("assistant", reply)
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Outline refresh failed: {exc}")

    slides = outline_payload.get("slides", [])
    hints = {
        "two_column":       "One bullet per line. First half → Left column, second half → Right column.",
        "timeline":         "Format: `Phase Label: Detail description` (one per line)",
        "icon_grid":        "Format: `Title: Detail description` (4 lines for 4 grid items)",
        "table":            "First line = column names separated by |. Then data rows with |.",
        "chart":            "Format: `Label : value` (numeric value 1-100, one per line)",
        "case_study":       "Start lines with `Company:` and `Result:` then add bullet points.",
        "big_stat":         "Start with `STAT: value` and `LABEL: description` then add bullet points.",
        "hybrid_insight":   "Start with `STAT: value` and `LABEL: description` then add bullet points.",
        "section_index":    "One section title per line.",
    }

    for index, slide in enumerate(slides):
        slide = ensure_editor_id(slide)
        slides[index] = slide
        original_slide = copy.deepcopy(slide)
        sid = slide["_editor_id"]
        layout = slide.get("layout", "bullets")
        hint = hints.get(layout, "One bullet point per line. Use **text** for bold.")
        prime_editor_widget_state(slide)

        with st.expander(
            f"**Slide {index+1}** — {slide.get('title','Untitled')}  `{layout}`",
            expanded=index < 2
        ):
            c1, c2 = st.columns([3, 1])
            with c1:
                title = st.text_input("Heading", key=f"t_{sid}")
            with c2:
                current_layout = st.session_state.get(f"l_{sid}", layout)
                if current_layout not in LAYOUT_OPTIONS:
                    current_layout = layout if layout in LAYOUT_OPTIONS else "bullets"
                new_layout = st.selectbox(
                    "Layout",
                    LAYOUT_OPTIONS,
                    index=LAYOUT_OPTIONS.index(current_layout),
                    key=f"l_{sid}",
                )

            subtitle = st.text_input("Subtitle (optional)", key=f"s_{sid}")

            if new_layout != layout:
                for fld in ("steps","grid_items","left_points","right_points",
                            "table_columns","table_rows","chart_data","metrics"):
                    slide.pop(fld, None)
                slide["layout"] = new_layout
                layout = new_layout

            bullet_text = st.text_area(
                "Content / talking points",
                height=180,
                help=hint,
                key=f"c_{sid}",
            )
            st.caption(f"💡 **Tip:** {hint}")

            slide["title"]    = title.strip() or f"Slide {index+1}"
            slide["subtitle"] = subtitle.strip()
            slide["layout"]   = layout
            slide = save_editor_content(slide, bullet_text)
            slides[index] = slide
            if slide != original_slide:
                rebuild_ppt_from_outline()

            uc, dc, rc = st.columns(3)
            if uc.button("⬆ Move Up", key=f"u_{sid}", disabled=index==0, use_container_width=True):
                slides[index-1], slides[index] = slides[index], slides[index-1]
                rebuild_ppt_from_outline()
                st.rerun()
            if dc.button("⬇ Move Down", key=f"d_{sid}", disabled=index==len(slides)-1, use_container_width=True):
                slides[index+1], slides[index] = slides[index], slides[index+1]
                rebuild_ppt_from_outline()
                st.rerun()
            if rc.button("🗑 Remove", key=f"r_{sid}", use_container_width=True):
                slides.pop(index)
                rebuild_ppt_from_outline()
                st.rerun()

    if st.session_state.get("ppt_bytes"):
        st.divider()
        st.markdown("### ✅ Presentation Ready")
        st.markdown(
            f"Your **{len(outline_payload.get('slides',[])) if outline_payload else 0} slide** presentation has been generated. "
            "Click below to download."
        )
        st.download_button(
            label="⬇️ Download PPT",
            data=st.session_state.ppt_bytes,
            file_name=st.session_state.get("ppt_filename", "presentation.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            type="primary",
        )

# ═══════════════════════════════════════════════════════════════════════════
#  API CALLS
# ═══════════════════════════════════════════════════════════════════════════

def request_outline(topic: str, num_slides: int, tone: str) -> dict:
    resp = requests.post(
        OUTLINE_URL,
        data={"topic": topic, "num_slides": num_slides, "tone": tone},
        timeout=(10, 180),
    )
    if resp.status_code != 200: raise RuntimeError(resp.text)
    payload = resp.json()
    usage = payload.get("usage")
    if not usage:
        st.warning("⚠️ No token usage received from backend")
    if isinstance(usage, dict):
        record_usage("outline", usage)
    payload["slides"] = [
        ensure_editor_id(s) for s in payload.get("slides", []) if isinstance(s, dict)
    ]
    return payload

def sanitize_outline_for_build(payload: dict) -> dict:
    safe = {
        "design_system": payload.get("design_system", {}) if isinstance(payload, dict) else {},
        "slides": [],
    }
    for slide in (payload.get("slides",[]) if isinstance(payload,dict) else []):
        if not isinstance(slide, dict): continue
        layout = str(slide.get("layout","bullets")).strip() or "bullets"
        s = {
            "title":    str(slide.get("title","")).strip(),
            "subtitle": str(slide.get("subtitle","")).strip(),
            "layout":   layout,
            "icon":     str(slide.get("icon","▸")).strip() or "▸",
            "content":  text_to_bullets(bullets_to_text(slide.get("content",[]))),
            "style":    slide.get("style",{}) if isinstance(slide.get("style"),dict) else {},
        }
        if layout == "two_column":
            s["left_title"]   = slide.get("left_title","Left")
            s["right_title"]  = slide.get("right_title","Right")
            lp = slide.get("left_points",[])  or []
            rp = slide.get("right_points",[]) or []
            if not lp and not rp:
                mid = max(1, len(s["content"])//2)
                lp, rp = s["content"][:mid], s["content"][mid:]
            s["left_points"]  = lp
            s["right_points"] = rp
        elif layout == "big_stat":
            s["stat"]       = slide.get("stat","—")
            s["stat_label"] = slide.get("stat_label","")
            s["stat_source"]= slide.get("stat_source","")
        elif layout == "timeline":
            s["steps"] = slide.get("steps",[]) or []
        elif layout == "icon_grid":
            s["grid_items"] = normalize_icon_grid_items(slide, bullets_to_text(slide.get("content", [])))
            s["content"] = [
                f"{item['title']}: {item['detail']}".rstrip(": ").strip()
                for item in s["grid_items"]
            ]
        elif layout == "case_study":
            s["company"] = slide.get("company","")
            s["result"]  = slide.get("result","")
            s["metrics"] = slide.get("metrics",[]) or []
        elif layout == "table":
            s["table_columns"] = slide.get("table_columns",[]) or []
            s["table_rows"]    = slide.get("table_rows",[])    or []
        elif layout == "chart":
            s["chart_title"]  = slide.get("chart_title","")
            s["chart_data"]   = slide.get("chart_data",[])  or []
            s["chart_source"] = slide.get("chart_source","")
        elif layout == "image_text_split":
            s["image_caption"] = slide.get("image_caption","")
            s["image_side"]    = slide.get("image_side","right")
            s["use_uploaded_image"] = bool(slide.get("use_uploaded_image"))
        elif layout == "hybrid_insight":
            s["stat"]       = slide.get("stat","—")
            s["stat_label"] = slide.get("stat_label","")
            s["chart_data"] = slide.get("chart_data",[]) or []
        elif layout == "section_index":
            s["sections"] = slide.get("sections", s["content"])
        safe["slides"].append(s)
    return safe

# ═══════════════════════════════════════════════════════════════════════════
#  MAIN UI
# ═══════════════════════════════════════════════════════════════════════════

init_state()

st.title("💬 AI PPT Chat Builder")
st.caption("Describe your deck, edit it in chat or in the panel below, then download.")

# ── Sidebar ────────────────────────────────────────────────────────────────
st.sidebar.title("⚙️ Deck Settings")
tone          = st.sidebar.selectbox("Presentation Style", ["Professional","Creative","Educational"], index=["Professional","Creative","Educational"].index(st.session_state.tone))
if st.sidebar.button("🔄 Reset Conversation", use_container_width=True):
    for k in ("messages","outline_payload","ppt_bytes","ppt_filename","topic","token_usage","last_usage","session_id","pending_outline_context","deck_history","current_deck_id","content_image_bytes","content_image_name"):
        st.session_state.pop(k, None)
    init_state()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("Token Usage")
last_usage = st.session_state.get("last_usage")
if last_usage:
    st.sidebar.write(f"Last request: {last_usage.get('kind','')}")
    st.sidebar.write(f"Input: {last_usage.get('prompt_tokens',0)}")
    st.sidebar.write(f"Output: {last_usage.get('completion_tokens',0)}")
    st.sidebar.write(f"Total: {last_usage.get('total_tokens',0)}")
else:
    st.sidebar.write("No usage yet.")

totals = st.session_state.get("token_usage", {})
st.sidebar.write("Session total")
st.sidebar.write(f"Input: {totals.get('prompt_tokens',0)}")
st.sidebar.write(f"Output: {totals.get('completion_tokens',0)}")
st.sidebar.write(f"Total: {totals.get('total_tokens',0)}")

st.sidebar.markdown("---")
st.sidebar.subheader("Session")
st.sidebar.code(st.session_state.get("session_id",""))
msgs = st.session_state.get("messages", [])
user_msgs = sum(1 for m in msgs if m.get("role") == "user")
assistant_msgs = sum(1 for m in msgs if m.get("role") == "assistant")
st.sidebar.write(f"User turns: {user_msgs}")
st.sidebar.write(f"Assistant turns: {assistant_msgs}")
st.sidebar.write(f"Total turns: {len(msgs)}")

# ── Chat history (always render BEFORE input) ──────────────────────────────
archived_index = 0
for msg in st.session_state.messages:
    if msg.get("kind") == "deck":
        deck_id = msg.get("deck_id")
        if deck_id and deck_id == st.session_state.get("current_deck_id"):
            render_current_deck()
        else:
            archived_index += 1
            render_archived_deck(get_archived_deck(deck_id), archived_index)
        continue
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ── Chat input ────────────────────────────────────────────────────────────
chat_value = st.chat_input(
    "Describe a topic, ask for edits, or upload a logo here",
    accept_file=True,
    file_type=["png", "jpg", "jpeg"],
)

if chat_value:
    uploaded_files = getattr(chat_value, "files", []) if hasattr(chat_value, "files") else []
    prompt = chat_value.text if hasattr(chat_value, "text") else chat_value
    prompt = (prompt or "").strip()
    pending_outline_context = st.session_state.get("pending_outline_context")
    uploaded_logo_name = None
    uploaded_content_image_name = None
    if uploaded_files:
        if should_treat_upload_as_logo(prompt, pending_outline_context):
            uploaded_logo_name = save_logo_upload(uploaded_files)
        else:
            uploaded_content_image_name = save_content_image_upload(uploaded_files)

    if (uploaded_logo_name or uploaded_content_image_name) and not prompt:
        reply = None
        if pending_outline_context and pending_outline_context.get("awaiting_logo") and uploaded_logo_name:
            with st.chat_message("user"):
                st.markdown(f"Uploaded logo: `{uploaded_logo_name}`")
            add_message("user", f"Uploaded logo: {uploaded_logo_name}")
            prompt = "use uploaded logo"
        elif pending_outline_context and pending_outline_context.get("awaiting_logo") and uploaded_content_image_name:
            with st.chat_message("user"):
                st.markdown(f"Uploaded image: `{uploaded_content_image_name}`")
            add_message("user", f"Uploaded image: {uploaded_content_image_name}")
            prompt = "use uploaded image"
        elif st.session_state.get("outline_payload"):
            if uploaded_content_image_name:
                st.session_state.outline_payload = attach_content_image_to_outline(st.session_state.outline_payload)
            rebuild_ppt_from_outline()
            if uploaded_logo_name:
                reply = f"Logo uploaded: `{uploaded_logo_name}`. I rebuilt the current deck with it."
            else:
                reply = f"Image uploaded: `{uploaded_content_image_name}`. I placed it into a relevant slide and rebuilt the current deck."
        else:
            if uploaded_logo_name:
                reply = f"Logo uploaded: `{uploaded_logo_name}`. I’ll use it for the deck. If you want, tell me the PPT topic or say `use this logo`."
            else:
                reply = f"Image uploaded: `{uploaded_content_image_name}`. I’ll place it at a relevant spot in the PPT when I generate the deck."
        if reply:
            with st.chat_message("assistant"):
                st.markdown(reply)
            add_message("assistant", reply)
        if not prompt:
            st.stop()
    if (uploaded_logo_name or uploaded_content_image_name) and prompt:
        with st.chat_message("user"):
            if uploaded_logo_name:
                st.markdown(f"{prompt}\n\nUploaded logo: `{uploaded_logo_name}`")
                add_message("user", f"{prompt}\n\nUploaded logo: {uploaded_logo_name}")
            else:
                st.markdown(f"{prompt}\n\nUploaded image: `{uploaded_content_image_name}`")
                add_message("user", f"{prompt}\n\nUploaded image: {uploaded_content_image_name}")
    else:
        if not prompt:
            st.stop()
        with st.chat_message("user"):
            st.markdown(prompt)
        add_message("user", prompt)

    # If we previously asked a clarification, merge the answer with the original intent.
    pending = st.session_state.get("pending_clarification")
    if pending and pending.get("intent"):
        prompt = f"{pending['intent']}\nClarification answer: {prompt}"
        st.session_state.pop("pending_clarification", None)

    outline_payload = st.session_state.outline_payload
    slides = outline_payload.get("slides", []) if outline_payload else []
    pending_outline_context = st.session_state.get("pending_outline_context")

    # Friendly greeting flow (avoid auto-generating a deck on "hi")
    if not pending_outline_context and is_greeting(prompt):
        if outline_payload:
            reply = (
                "Hi there! 👋 Want to tweak a slide or start a new deck? "
                "You can say something like: \"Edit slide 2, change title to Market Analysis\"."
            )
        else:
            reply = (
                "Hi there! 👋 Tell me the presentation topic and I’ll draft the slide sections. If you want branding, you can upload a logo right here in the chat box."
            )
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.stop()

    if not pending_outline_context and is_smalltalk(prompt):
        if outline_payload:
            reply = (
                "Got it. Want to tweak a slide or start a new deck? "
                "You can say something like: \"Edit slide 2, change title to Market Analysis\"."
            )
        else:
            reply = (
                "Whenever you're ready, tell me the presentation topic and I’ll draft the slide sections. If you want branding, upload a logo here in chat."
            )
        with st.chat_message("assistant"):
            st.markdown(reply)
        add_message("assistant", reply)
        st.stop()

    # CASE 1: FIRST TIME OR EXPLICIT NEW DECK → GENERATE OUTLINE
    if pending_outline_context or is_new_presentation_request(prompt, bool(outline_payload)):
        with st.spinner("✍️ Drafting slide sections..."):
            try:
                current_turn_has_logo = bool(uploaded_logo_name)
                current_turn_has_content_image = bool(uploaded_content_image_name)
                if pending_outline_context:
                    topic_prompt = pending_outline_context.get("topic", "")
                    requested_slides = pending_outline_context.get("num_slides") or extract_num_slides(prompt) or extract_num_slides(topic_prompt) or 10
                    awaiting_logo = bool(pending_outline_context.get("awaiting_logo"))
                else:
                    topic_prompt = topic_from_prompt(prompt)
                    requested_slides = extract_num_slides(prompt) or extract_num_slides(topic_prompt)
                    awaiting_logo = False
                    if requested_slides is None:
                        st.session_state.pending_outline_context = {"topic": topic_prompt}
                        reply = "How many slides would you like? If you don’t have a preference, I’ll make 10."
                        with st.chat_message("assistant"):
                            st.markdown(reply)
                        add_message("assistant", reply)
                        st.stop()
                if not awaiting_logo:
                    st.session_state.pending_outline_context = {
                        "topic": topic_prompt,
                        "num_slides": requested_slides,
                        "awaiting_logo": True,
                    }
                    reply = "Would you like to upload a logo for branding? Upload it here."
                    with st.chat_message("assistant"):
                        st.markdown(reply)
                    add_message("assistant", reply)
                    st.stop()
                if awaiting_logo and not current_turn_has_logo and not should_skip_logo_step(prompt):
                    reply = "You can upload the logo here, or reply `skip` and I’ll continue without one."
                    with st.chat_message("assistant"):
                        st.markdown(reply)
                    add_message("assistant", reply)
                    st.stop()
                new_logo_bytes = st.session_state.get("logo_bytes") if current_turn_has_logo else None
                new_logo_name = st.session_state.get("logo_name") if current_turn_has_logo else None
                preserve_existing_content_image = not outline_payload
                new_content_image_bytes = st.session_state.get("content_image_bytes") if (current_turn_has_content_image or preserve_existing_content_image) else None
                new_content_image_name = st.session_state.get("content_image_name") if (current_turn_has_content_image or preserve_existing_content_image) else None
                if outline_payload:
                    archive_current_deck()
                    reset_current_deck(keep_messages=True)
                if new_logo_bytes:
                    st.session_state.logo_bytes = new_logo_bytes
                    st.session_state.logo_name = new_logo_name
                if new_content_image_bytes:
                    st.session_state.content_image_bytes = new_content_image_bytes
                    st.session_state.content_image_name = new_content_image_name
                payload = request_outline(topic_prompt, requested_slides, tone)
                payload = attach_content_image_to_outline(payload)
                st.session_state.outline_payload = payload
                st.session_state.topic = topic_prompt
                st.session_state.num_slides = requested_slides
                st.session_state.tone = tone
                st.session_state.current_deck_id = uuid4().hex
                st.session_state.pending_outline_context = None
                sync_all_editor_widgets(payload.get("slides", []))
                # Auto-build PPT after outline
                rebuild_ppt_from_outline()
                reply = summarize_outline(payload.get("slides", []))
                if not st.session_state.get("logo_bytes"):
                    reply += "\n\nIf you want branding on the deck, upload a logo in the chat box and I’ll use it."
                with st.chat_message("assistant"):
                    st.markdown(reply)
                add_message("assistant", reply)
                add_deck_message(st.session_state.current_deck_id)
                st.rerun()
            except Exception as e:
                st.error(f"Outline generation failed: {e}")

    # CASE 2: EDIT EXISTING SLIDES (LLM)
    else:
        # Deterministic handling for "convert slide X into thank you slide"
        target_num = extract_slide_number(prompt)
        if outline_payload and target_num and is_thank_you_request(prompt) and is_convert_request(prompt):
            if 1 <= target_num <= len(slides):
                slides[target_num - 1] = make_thank_you_slide(slides[target_num - 1])
                st.session_state.outline_payload["slides"] = slides
                sync_all_editor_widgets(slides)
                rebuild_ppt_from_outline()
                move_current_deck_to_chat_end()
                reply = "✅ Slide updated to a Thank You slide and PPT rebuilt. You can now download."
                with st.chat_message("assistant"):
                    st.markdown(reply)
                add_message("assistant", reply)
                st.rerun()
        try:
            needs_clarify, question = needs_clarification_for_edit(prompt, slides)
            if needs_clarify:
                render_clarification(question or "Which slide would you like to edit?", intent_text=prompt)
                st.stop()
            result, usage = interpret_edit_with_llm(prompt, slides)
            record_usage("edit", usage)
            action = (result or {}).get("action", "")
            if action == "clarify":
                render_clarification(result.get("question") or "Which slide would you like to edit?", intent_text=prompt)
                st.stop()
            if action == "add":
                slide = result.get("slide")
                if not isinstance(slide, dict):
                    slide = draft_slide_from_request(prompt, slides)
                if not slide.get("title") or not slide.get("content"):
                    slide = draft_slide_from_request(prompt, slides)
                slide = normalize_slide(slide)
                position = result.get("position", "end")
                updated_slides = refresh_section_index_slide(apply_add_action(slides, slide, position))
                st.session_state.outline_payload["slides"] = updated_slides
                sync_all_editor_widgets(updated_slides)
                rebuild_ppt_from_outline()
                move_current_deck_to_chat_end()
                reply = "✅ Slide added and PPT rebuilt."
                with st.chat_message("assistant"):
                    st.markdown(reply)
                add_message("assistant", reply)
                st.rerun()
            if action == "remove":
                targets = result.get("targets", [])
                if isinstance(targets, list):
                    resolved = []
                    for t in targets:
                        if isinstance(t, int) and 1 <= t <= len(slides):
                            resolved.append(t)
                    updated_slides = refresh_section_index_slide(
                        [s for i, s in enumerate(slides, start=1) if i not in resolved]
                    )
                    if updated_slides:
                        st.session_state.outline_payload["slides"] = updated_slides
                        sync_all_editor_widgets(updated_slides)
                        rebuild_ppt_from_outline()
                        move_current_deck_to_chat_end()
                        reply = "✅ Slide(s) removed and PPT rebuilt."
                        with st.chat_message("assistant"):
                            st.markdown(reply)
                        add_message("assistant", reply)
                        st.rerun()
                render_clarification("Which slide(s) should I remove?")
                st.stop()
            if action == "edit":
                updated_slides = refresh_section_index_slide(
                    [ensure_editor_id(s) for s in result.get("slides", [])]
                )
                st.session_state.outline_payload["slides"] = updated_slides
                sync_all_editor_widgets(updated_slides)
                # Auto-rebuild PPT
                rebuild_ppt_from_outline()
                move_current_deck_to_chat_end()
                reply = "✅ Slides updated and PPT rebuilt. You can now download."
                with st.chat_message("assistant"):
                    st.markdown(reply)
                add_message("assistant", reply)
                st.rerun()
        except Exception as e:
            st.error(f"LLM edit failed: {e}")