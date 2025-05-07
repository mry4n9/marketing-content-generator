# scraper.py
import requests
from bs4 import BeautifulSoup
import json
from openai import OpenAI
from utils import get_model_name, get_max_scrape_tokens

def get_website_text_content(url):
    """Fetches and extracts visible text content from a URL."""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status() # Raise an exception for HTTP errors
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Remove script and style elements
        for script_or_style in soup(["script", "style"]):
            script_or_style.decompose()
            
        # Get text
        text = soup.get_text(separator=' ', strip=True)
        
        # Limit text length to avoid excessive token usage for LLM processing
        # A more sophisticated chunking/summarization might be needed for very large pages
        max_len = get_max_scrape_tokens() * 3 # Approx 3 chars per token
        return text[:max_len]
    except requests.exceptions.RequestException as e:
        print(f"Error fetching URL {url}: {e}")
        return None
    except Exception as e:
        print(f"Error parsing content from {url}: {e}")
        return None

def extract_structured_data_from_text(text_content, api_key):
    """Uses OpenAI to extract structured company information from text."""
    if not text_content:
        return None

    client = OpenAI(api_key=api_key)
    model_name = get_model_name()

    prompt = f"""
    Analyze the following website text content and extract the specified company information.
    If a piece of information is not explicitly found or cannot be reasonably inferred, use "Not found" or an empty list/string as appropriate for the field.
    Return the information as a JSON object with the following keys:
    - "company_name": The official name of the company.
    - "tagline": The company's main tagline or slogan.
    - "mission_statement": The company's mission statement, if available.
    - "industry": The primary industry the company operates in.
    - "products_services": A list of key products, services, or offerings.
    - "usps_value_proposition": Unique Selling Propositions or the core value proposition.
    - "target_audience": A description of the typical target audience.
    - "tone_of_voice": The perceived tone of voice (e.g., formal, friendly, technical, inspirational).
    - "ctas": A list of main Call-to-Actions found (e.g., "Request a Demo", "Shop Now").

    Website Text:
    "{text_content[:get_max_scrape_tokens()*2]}" 

    JSON Output:
    """ # Truncate again to be safe with prompt length

    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": "You are an expert in extracting structured information from website content. Output ONLY the JSON object."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.2 # Lower temperature for more factual extraction
        )
        extracted_json_str = response.choices[0].message.content
        extracted_data = json.loads(extracted_json_str)
        
        # Ensure all keys are present, defaulting to "Not found" or empty list
        keys_to_ensure = ["company_name", "tagline", "mission_statement", "industry", 
                          "products_services", "usps_value_proposition", "target_audience", 
                          "tone_of_voice", "ctas"]
        
        for key in keys_to_ensure:
            if key not in extracted_data:
                if key in ["products_services", "ctas"]:
                    extracted_data[key] = []
                else:
                    extracted_data[key] = "Not found"
            elif extracted_data[key] is None: # Handle nulls from LLM
                 if key in ["products_services", "ctas"]:
                    extracted_data[key] = []
                 else:
                    extracted_data[key] = "Not found"


        return extracted_data
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON from OpenAI response: {e}")
        print(f"Problematic JSON string: {extracted_json_str}")
        return None
    except Exception as e:
        print(f"Error calling OpenAI for structured data extraction: {e}")
        return None


def scrape_website_data(url, api_key):
    """
    Scrapes website for company info.
    First, gets all text. Then, uses LLM to extract structured info.
    """
    print(f"Scraping website: {url}")
    website_text = get_website_text_content(url)
    if not website_text:
        print("Failed to retrieve website content.")
        return None
    
    print("Extracting structured data using LLM...")
    structured_data = extract_structured_data_from_text(website_text, api_key)
    
    if structured_data:
        print("Successfully extracted structured data.")
    else:
        print("Failed to extract structured data using LLM.")
        # Fallback: try to get at least the title as company name
        try:
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            title_tag = soup.find('title')
            company_name_from_title = title_tag.string.strip() if title_tag else "Unknown Company"
            return {
                "company_name": company_name_from_title,
                "tagline": "Not found", "mission_statement": "Not found", "industry": "Not found",
                "products_services": [], "usps_value_proposition": "Not found",
                "target_audience": "Not found", "tone_of_voice": "Not found", "ctas": []
            }
        except Exception:
             return {
                "company_name": "Unknown Company",
                "tagline": "Not found", "mission_statement": "Not found", "industry": "Not found",
                "products_services": [], "usps_value_proposition": "Not found",
                "target_audience": "Not found", "tone_of_voice": "Not found", "ctas": []
            }

    return structured_data