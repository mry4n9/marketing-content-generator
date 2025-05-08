# app.py
import streamlit as st
import re # For sanitizing company name for filename
import time # For potential delays if rate limits are hit often
from utils import load_openai_api_key
from scraper import scrape_website_data
from doc_parser import extract_text_from_uploaded_files
from openai_handler import (
    generate_email_content,
    generate_linkedin_facebook_content,
    generate_google_search_ads,
    generate_google_display_ads,
    generate_reasoning_text
)
from excel_generator import create_excel_workbook

# --- Page Configuration ---
st.set_page_config(page_title="Marketing Content Generator", layout="wide", initial_sidebar_state="expanded")

# --- Load API Key ---
OPENAI_API_KEY = load_openai_api_key()

# --- Helper function for URL formatting --- Point 4
def format_url(url_input):
    """Prepends https:// if scheme is missing."""
    if url_input and not (url_input.startswith("http://") or url_input.startswith("https://")):
        return "https://" + url_input
    return url_input

# --- Main App UI ---
st.title("🚀 Marketing Content Generator")
st.markdown("""
This tool helps generate tailored marketing content for clients. 
Upload client materials, specify objectives, and let AI do the heavy lifting!
""")

if not OPENAI_API_KEY:
    st.error("OpenAI API Key is not configured. Please add it to `.streamlit/secrets.toml` and restart the app.")
    st.stop()

# --- Input Fields ---
st.sidebar.header("Client & Campaign Setup")

# Use different variable names for raw input vs formatted URL
client_website_url_raw = st.sidebar.text_input("Client's Website URL", placeholder="www.example.com")
additional_materials = st.sidebar.file_uploader(
    "Upload Additional Materials (PDF, PPTX)", 
    type=["pdf", "pptx"], 
    accept_multiple_files=True
)

lead_objective_options = ["Demo Booking", "Sales Meeting"]
lead_objective_type = st.sidebar.selectbox("Lead Objective Type", lead_objective_options)
lead_objective_url_raw = st.sidebar.text_input("URL for Selected Lead Objective", placeholder="www.example.com/book-demo")

downloadable_asset_url_raw = st.sidebar.text_input(
    "URL of a Key Downloadable Asset (Optional, e.g., whitepaper)", 
    placeholder="www.example.com/assets/whitepaper.pdf"
)

num_content_pieces = st.sidebar.slider("Number of Content Pieces per Objective (for Email/Social)", 10, 20, 10)

# --- Generate Button ---
if st.sidebar.button("✨ Generate Content", type="primary", use_container_width=True):
    # Point 4: Format URLs before validation and use
    client_website_url = format_url(client_website_url_raw)
    lead_objective_url = format_url(lead_objective_url_raw)
    downloadable_asset_url = format_url(downloadable_asset_url_raw)
    
    # --- Input Validation ---
    valid_inputs = True
    if not client_website_url_raw: # Check raw input for presence
        st.error("Please provide the client's website URL.")
        valid_inputs = False
    # After formatting, check if it's a valid http/https URL (basic check)
    elif client_website_url and not (client_website_url.startswith("http://") or client_website_url.startswith("https://")):
        st.error("Please enter a valid website URL (e.g., https://www.example.com or www.example.com).")
        valid_inputs = False
        
    if not lead_objective_url_raw: # Check raw input for presence
        st.error("Please provide the URL for the selected lead objective.")
        valid_inputs = False
    elif lead_objective_url and not (lead_objective_url.startswith("http://") or lead_objective_url.startswith("https://")):
        st.error("Please enter a valid URL for the lead objective (e.g., https://www.example.com/demo or www.example.com/demo).")
        valid_inputs = False

    # Only validate downloadable_asset_url if raw input was provided
    if downloadable_asset_url_raw and downloadable_asset_url and not (downloadable_asset_url.startswith("http://") or downloadable_asset_url.startswith("https://")):
        st.error("Please enter a valid URL for the downloadable asset or leave it blank.")
        valid_inputs = False
        
    if valid_inputs:
        with st.spinner("Hold tight! Generating amazing content... This might take a few minutes... ⏳"):
            try:
                # 1. Scrape Website
                st.subheader("Step 1: Scraping Website Data...")
                scraped_data = scrape_website_data(client_website_url, OPENAI_API_KEY)
                if not scraped_data:
                    st.error("Failed to scrape website data. Please check the URL and try again.")
                    st.stop() # Use st.stop() to halt execution cleanly on critical failure
                st.success(f"Successfully scraped data for: {scraped_data.get('company_name', 'Unknown Company')}")
                with st.expander("View Scraped Data"):
                    st.json(scraped_data)

                # 2. Parse Uploaded Documents
                st.subheader("Step 2: Processing Uploaded Documents...")
                additional_docs_text = ""
                if additional_materials:
                    additional_docs_text = extract_text_from_uploaded_files(additional_materials)
                    st.success(f"Successfully processed {len(additional_materials)} uploaded document(s).")
                    with st.expander("View Extracted Text from Documents (First 500 chars)"):
                        st.text(additional_docs_text[:500] + "..." if additional_docs_text else "No text extracted.")
                else:
                    st.info("No additional documents uploaded.")

                # 3. Generate Content (Store all in a dictionary)
                all_generated_content = {}
                
                # Emails
                st.subheader(f"Step 3.1: Generating {num_content_pieces} Email Versions...")
                email_content = generate_email_content(
                    OPENAI_API_KEY, scraped_data, additional_docs_text, 
                    lead_objective_type, lead_objective_url, downloadable_asset_url, 
                    num_content_pieces
                )
                all_generated_content["email"] = email_content
                st.success(f"Generated {len(email_content)} email versions.")

                # LinkedIn Ads
                st.subheader(f"Step 3.2: Generating LinkedIn Ad Versions...")
                # time.sleep(3) # Optional: Add delay if hitting rate limits frequently
                linkedin_content = generate_linkedin_facebook_content(
                    OPENAI_API_KEY, "LinkedIn", scraped_data, additional_docs_text,
                    lead_objective_type, lead_objective_url, downloadable_asset_url,
                    num_content_pieces 
                )
                all_generated_content["linkedin"] = linkedin_content
                st.success(f"Generated {len(linkedin_content)} LinkedIn ad versions.")

                # Facebook Ads
                st.subheader(f"Step 3.3: Generating Facebook Ad Versions...")
                # time.sleep(3) # Optional: Add delay
                facebook_content = generate_linkedin_facebook_content(
                    OPENAI_API_KEY, "Facebook", scraped_data, additional_docs_text,
                    lead_objective_type, lead_objective_url, downloadable_asset_url,
                    num_content_pieces
                )
                all_generated_content["facebook"] = facebook_content
                st.success(f"Generated {len(facebook_content)} Facebook ad versions.")

                # Google Search Ads
                st.subheader(f"Step 3.4: Generating Google Search Ad Copy...")
                # time.sleep(2) # Optional: Add delay
                google_search_content = generate_google_search_ads(
                    OPENAI_API_KEY, scraped_data, additional_docs_text,
                    lead_objective_type, lead_objective_url, downloadable_asset_url
                )
                all_generated_content["google_search"] = google_search_content
                st.success("Generated Google Search ad copy.")
                
                # Google Display Ads
                st.subheader(f"Step 3.5: Generating Google Display Ad Copy...")
                # time.sleep(2) # Optional: Add delay
                google_display_content = generate_google_display_ads(
                    OPENAI_API_KEY, scraped_data, additional_docs_text,
                    lead_objective_type, lead_objective_url, downloadable_asset_url
                )
                all_generated_content["google_display"] = google_display_content
                st.success("Generated Google Display ad copy.")

                # Reasoning Text
                st.subheader(f"Step 3.6: Generating Reasoning Text...")
                # time.sleep(1) # Optional: Add delay
                reasoning = generate_reasoning_text(
                    OPENAI_API_KEY, scraped_data, additional_docs_text,
                    lead_objective_type, lead_objective_url, downloadable_asset_url
                )
                all_generated_content["reasoning_text"] = reasoning
                # Point 5: Reasoning error (Rate Limit) - Display warning in UI
                if "Error code: 429" in reasoning and "rate_limit_exceeded" in reasoning:
                    st.warning(
                        "⚠️ Reasoning generation hit an API rate limit. "
                        "The detailed error message has been included in the 'Reasoning' sheet of the Excel file. "
                        "To resolve this, you may need to check your OpenAI account's rate limits, add a payment method, or wait before trying again. "
                        "Other content has been generated successfully."
                    )
                elif "Error generating reasoning text" in reasoning: # Catch other reasoning errors
                     st.warning(f"⚠️ Could not fully generate reasoning text. The error has been included in the Excel: {reasoning[:100]}...")
                else:
                    st.success("Generated reasoning text.")


                # 4. Create Excel File
                st.subheader("Step 4: Compiling Excel Report...")
                company_name_raw_for_file = scraped_data.get("company_name", "client") # Use a distinct var name
                company_name_for_file = re.sub(r'[^\w\s-]', '', company_name_raw_for_file).strip().replace(' ', '_')
                if not company_name_for_file: company_name_for_file = "client_content" # Fallback
                
                excel_file_name = f"{company_name_for_file}_lead_content.xlsx"
                
                excel_bytes = create_excel_workbook(all_generated_content, scraped_data, company_name_for_file)
                st.success("Excel report compiled successfully!")

                # 5. Enable Download
                st.subheader("Step 5: Download Your Report")
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_bytes,
                    file_name=excel_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.balloons()

            except Exception as e:
                st.error(f"An unexpected error occurred during content generation: {e}")
                import traceback
                st.error(f"Traceback: {traceback.format_exc()}")
    # This else corresponds to 'if valid_inputs:'
    # else:
    #     st.warning("Please correct the input errors above before generating content.") # This message is implicitly handled by individual error messages now.

else:
    st.info("Fill in the details in the sidebar and click 'Generate Content' to start.")

st.sidebar.markdown("---")
st.sidebar.markdown("Developed for internal branding & marketing consultancy.")