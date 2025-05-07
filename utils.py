# utils.py
import streamlit as st
import os

# --- Constants ---
# Use gpt-4o-mini as gpt-4.1-nano is not a standard public model name.
# This can be changed to other models like "gpt-4-turbo" if preferred.
OPENAI_MODEL_NAME = "gpt-4o-mini"
MAX_TOKENS_WEBSITE_SCRAPE_ASSIST = 4000 # Max tokens for LLM to process for website data extraction
MAX_CONTENT_TOKENS = 2000 # Max tokens for content generation calls, adjust as needed

# --- Functions ---
def load_openai_api_key():
    """Loads the OpenAI API key from Streamlit secrets."""
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
        if not api_key or not api_key.startswith("sk-"):
            st.error("OpenAI API key is not set correctly in secrets.toml. It should start with 'sk-'.")
            return None
        return api_key
    except KeyError:
        st.error("OpenAI API key not found in st.secrets. Make sure OPENAI_API_KEY is set in .streamlit/secrets.toml")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred while loading the API key: {e}")
        return None

def get_model_name():
    """Returns the configured OpenAI model name."""
    return OPENAI_MODEL_NAME

def get_max_scrape_tokens():
    """Returns max tokens for website scraping LLM call."""
    return MAX_TOKENS_WEBSITE_SCRAPE_ASSIST

def get_max_content_tokens():
    """Returns max tokens for content generation LLM calls."""
    return MAX_CONTENT_TOKENS