import os
from dotenv import load_dotenv
from openai import OpenAI

# Load environment variables
load_dotenv()

# =========================
# 🔹 Azure OpenAI (LLM)
# =========================
# AZURE_KEY = os.getenv("AZURE_OPENAI_LLM_KEY")
# AZURE_ENDPOINT = os.getenv("AZURE_LLM_ENDPOINT")
# AZURE_API_VERSION = os.getenv("AZURE_LLM_API_VERSION")
# AZURE_DEPLOYMENT = os.getenv("AZURE_LLM_DEPLOYMENT_41_MINI")

api_key=os.getenv("OPENAI_API_KEY"),  # your OpenRouter key
base_url="https://openrouter.ai/api/v1",
model="openai/gpt-3.5-turbo",  # free model
temperature=0.7

# =========================
# 🔹 Hugging Face (optional - for image)
# =========================
# HF_API_KEY = os.getenv("HF_API_KEY")

# =========================
# 🔹 Debug check (optional)
# =========================
if not AZURE_KEY:
    print("⚠️ Warning: AZURE_OPENAI_LLM_KEY not set")

if not AZURE_ENDPOINT:
    print("⚠️ Warning: AZURE_LLM_ENDPOINT not set")