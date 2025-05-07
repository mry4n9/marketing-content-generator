# openai_handler.py
import json
from openai import OpenAI
from utils import get_model_name, get_max_content_tokens

def _call_openai_api(api_key, system_prompt, user_prompt, expecting_json=True):
    """Helper function to call OpenAI API."""
    client = OpenAI(api_key=api_key)
    model_name = get_model_name()
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    
    try:
        completion_args = {
            "model": model_name,
            "messages": messages,
            "temperature": 0.7, # Creative tasks can have higher temperature
            "max_tokens": get_max_content_tokens() 
        }
        if expecting_json:
            completion_args["response_format"] = {"type": "json_object"}

        response = client.chat.completions.create(**completion_args)
        content = response.choices[0].message.content

        if expecting_json:
            try:
                return json.loads(content)
            except json.JSONDecodeError as e:
                print(f"JSON Decode Error: {e}\nRaw content: {content}")
                # Fallback: try to extract JSON from a potentially messy string
                try:
                    # Find the first '{' and last '}'
                    start_index = content.find('{')
                    end_index = content.rfind('}')
                    if start_index != -1 and end_index != -1 and end_index > start_index:
                        json_str_candidate = content[start_index : end_index+1]
                        return json.loads(json_str_candidate)
                    else: # Try to find JSON array
                        start_index = content.find('[')
                        end_index = content.rfind(']')
                        if start_index != -1 and end_index != -1 and end_index > start_index:
                            json_str_candidate = content[start_index : end_index+1]
                            return json.loads(json_str_candidate)
                except json.JSONDecodeError:
                    print("Fallback JSON extraction also failed.")
                    raise # Re-raise original error if fallback fails
                raise # Re-raise original error if initial parsing fails
        return content # For non-JSON responses (like reasoning text)
    except Exception as e:
        print(f"Error calling OpenAI API: {e}")
        raise # Re-raise to be handled by caller

def _build_base_context_prompt(scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url):
    context = f"""
    ### Client Information:
    - Company Name: {scraped_data.get('company_name', 'N/A')}
    - Tagline: {scraped_data.get('tagline', 'N/A')}
    - Mission Statement: {scraped_data.get('mission_statement', 'N/A')}
    - Industry: {scraped_data.get('industry', 'N/A')}
    - Products/Services: {', '.join(scraped_data.get('products_services', [])) if scraped_data.get('products_services') else 'N/A'}
    - USPs/Value Proposition: {scraped_data.get('usps_value_proposition', 'N/A')}
    - Target Audience: {scraped_data.get('target_audience', 'N/A')}
    - Tone of Voice: {scraped_data.get('tone_of_voice', 'N/A')}
    - Existing CTAs: {', '.join(scraped_data.get('ctas', [])) if scraped_data.get('ctas') else 'N/A'}

    ### Additional Information from Uploaded Documents:
    {additional_docs_text if additional_docs_text else "No additional documents provided."}

    ### Marketing Campaign Goals:
    - Primary Lead Objective: {lead_objective_type}
    - URL for Primary Objective: {lead_objective_url}
    """
    if downloadable_asset_url:
        context += f"- URL for Downloadable Asset Promotion: {downloadable_asset_url}\n"
    return context

def generate_email_content(api_key, scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url, num_emails):
    base_context = _build_base_context_prompt(scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url)
    system_prompt = "You are a creative marketing copywriter specializing in email campaigns. Generate content as a JSON list of objects."
    user_prompt = f"""
    {base_context}

    ### Task:
    Generate {num_emails} unique email versions for a marketing campaign.
    The primary objective for these emails is: "{lead_objective_type}".
    The main CTA should direct to: {lead_objective_url}.

    ### Email Structure (for each email):
    - "Objective": "{lead_objective_type}" (This should be the value of the primary lead objective)
    - "Headline": A captivating headline for the email content itself (not the subject line).
    - "SubjectLine": An engaging email subject line (max 70 characters).
    - "Body": Email body text. If generating multiple emails ({num_emails} > 1), ensure weekly progression or evolving messaging across the emails. Embed the CTA naturally within the body. Max 300 words.
    - "CTA": A brief description of the call to action text used in the body (e.g., "Book a Demo", "Learn More Here"). This CTA should align with the objective URL: {lead_objective_url}.

    ### Output Format:
    Return a JSON list where each element is an object representing one email, with keys "Objective", "Headline", "SubjectLine", "Body", "CTA".
    Example for one email:
    {{
        "Objective": "{lead_objective_type}",
        "Headline": "Unlock New Growth Opportunities",
        "SubjectLine": "Your Path to Success Starts Here!",
        "Body": "Discover how [Company Name] can help you achieve... Our solutions offer [USP]. Click here to [CTA description linked to {lead_objective_url}].",
        "CTA": "Book Your Free Consultation"
    }}
    
    Generate exactly {num_emails} such email objects in the list.
    """
    try:
        response_data = _call_openai_api(api_key, system_prompt, user_prompt, expecting_json=True)
        if isinstance(response_data, list) and all(isinstance(item, dict) for item in response_data):
            # Add "Version #"
            for i, item in enumerate(response_data):
                item["Version #"] = i + 1
            return response_data
        else: # LLM might return a dict with a key like "emails"
            if isinstance(response_data, dict):
                for key in response_data: # try to find the list
                    if isinstance(response_data[key], list):
                        for i, item in enumerate(response_data[key]):
                            item["Version #"] = i + 1
                        return response_data[key]
            print(f"Unexpected JSON structure for emails: {response_data}")
            return [] # Fallback
    except Exception as e:
        print(f"Error generating email content: {e}")
        return [] # Return empty list on error

def generate_linkedin_facebook_content(api_key, platform, scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url, num_pieces_per_objective):
    base_context = _build_base_context_prompt(scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url)
    objectives = ["Brand Awareness", "Demand Gen", "Demand Capture"]
    all_ads = []
    
    # Determine destination URL logic
    # If downloadable_asset_url is provided, it can be used for some CTAs.
    # lead_objective_url is for Demo Booking/Sales Meeting.
    # The LLM should pick the most relevant one based on the ad's specific objective.

    system_prompt = f"You are a creative marketing copywriter specializing in {platform} ads. Generate content as a JSON list of objects."

    for i, ad_objective in enumerate(objectives):
        user_prompt = f"""
        {base_context}

        ### Task:
        Generate {num_pieces_per_objective} unique {platform} ad versions.
        The specific objective for this batch of ads is: "{ad_objective}".
        
        Available destination URLs:
        1. Primary Objective URL ({lead_objective_type}): {lead_objective_url}
        {f"2. Downloadable Asset URL: {downloadable_asset_url}" if downloadable_asset_url else ""}

        Choose the most appropriate destination URL based on the ad's message and objective ({ad_objective}).

        ### {platform} Ad Structure (for each ad):
        - "AdName": A descriptive ad name (up to 255 characters). Example: "{scraped_data.get('company_name', 'Client')} - {ad_objective} Ad - V{{version_num}}"
        - "Objective": "{ad_objective}"
        """
        if platform == "LinkedIn":
            user_prompt += """
        - "IntroductoryText": Hook in the first 150 chars, max 500 chars. Embed relevant emojis.
        - "ImageCopy": Short text for the accompanying image/visual (max 30 words).
        - "Headline": Ad headline (up to 70 characters).
        - "Destination": The chosen URL (from available URLs) for this ad.
        - "CTAButton": Suggested CTA button text. Options: [Learn More, Download, Register, Request Demo, Sign Up, Get Offer]. Choose one.
            """
        elif platform == "Facebook":
            user_prompt += """
        - "PrimaryText": Hook in the first 125 chars, max 500 chars. Embed relevant emojis.
        - "ImageCopy": Short text for the accompanying image/visual (max 30 words).
        - "Headline": Ad headline (up to 27 characters).
        - "LinkDescription": Short description appearing below headline (up to 27 characters).
        - "Destination": The chosen URL (from available URLs) for this ad.
        - "CTAButton": Suggested CTA button text. Options: [Learn More, Download, Book Now, Sign Up, Get Offer, Shop Now]. Choose one.
            """
        user_prompt += f"""
        ### Output Format:
        Return a JSON list where each element is an object representing one ad.
        Generate exactly {num_pieces_per_objective} such ad objects in the list for the "{ad_objective}" objective.
        """
        try:
            print(f"Generating {platform} ads for objective: {ad_objective}")
            response_data = _call_openai_api(api_key, system_prompt, user_prompt, expecting_json=True)
            
            current_ads = []
            if isinstance(response_data, list) and all(isinstance(item, dict) for item in response_data):
                current_ads = response_data
            elif isinstance(response_data, dict): # LLM might return a dict with a key like "ads"
                 for key in response_data:
                    if isinstance(response_data[key], list):
                        current_ads = response_data[key]
                        break
            
            if not current_ads:
                print(f"No ads generated or unexpected format for {platform} - {ad_objective}")

            # Add Version # and ensure Objective is set
            for k, ad_item in enumerate(current_ads):
                ad_item["Version #"] = (i * num_pieces_per_objective) + k + 1
                ad_item["Objective"] = ad_objective # Ensure objective is correctly set
            all_ads.extend(current_ads)

        except Exception as e:
            print(f"Error generating {platform} content for objective {ad_objective}: {e}")
            # Add placeholder if generation fails for this objective to maintain structure
            for k in range(num_pieces_per_objective):
                 placeholder_ad = {
                    "Version #": (i * num_pieces_per_objective) + k + 1,
                    "AdName": f"Error generating ad - {ad_objective} - V{k+1}",
                    "Objective": ad_objective,
                 }
                 if platform == "LinkedIn":
                    placeholder_ad.update({"IntroductoryText": "Error", "ImageCopy": "Error", "Headline": "Error", "Destination": "Error", "CTAButton": "Error"})
                 elif platform == "Facebook":
                    placeholder_ad.update({"PrimaryText": "Error", "ImageCopy": "Error", "Headline": "Error", "LinkDescription": "Error", "Destination": "Error", "CTAButton": "Error"})
                 all_ads.append(placeholder_ad)
                 
    return all_ads


def generate_google_search_ads(api_key, scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url):
    base_context = _build_base_context_prompt(scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url)
    system_prompt = "You are an expert Google Search Ads copywriter. Generate content as a JSON object."
    user_prompt = f"""
    {base_context}

    ### Task:
    Generate copy for Google Search Ads (Responsive Search Ad format).
    The ads should drive traffic towards "{lead_objective_type}" at {lead_objective_url}, or promote the downloadable asset if relevant ({downloadable_asset_url}).

    ### Ad Components:
    - Headlines: Create 15 unique headlines. Each headline must be a maximum of 30 characters.
    - Descriptions: Create 4 unique descriptions. Each description must be a maximum of 90 characters.

    ### Output Format:
    Return a JSON object with two keys: "headlines" (a list of 15 strings) and "descriptions" (a list of 4 strings).
    Example:
    {{
        "headlines": ["Headline 1 (<=30)", "Headline 2 (<=30)", ... (15 total)],
        "descriptions": ["Description 1 (<=90)", "Description 2 (<=90)", ... (4 total)]
    }}
    Ensure all character limits are strictly followed.
    """
    try:
        response_data = _call_openai_api(api_key, system_prompt, user_prompt, expecting_json=True)
        if isinstance(response_data, dict) and "headlines" in response_data and "descriptions" in response_data:
            # Validate counts and lengths (optional, but good practice)
            response_data["headlines"] = [h[:30] for h in response_data.get("headlines", [])][:15]
            response_data["descriptions"] = [d[:90] for d in response_data.get("descriptions", [])][:4]
            # Fill if not enough generated
            while len(response_data["headlines"]) < 15: response_data["headlines"].append("Generated Headline Placeholder")
            while len(response_data["descriptions"]) < 4: response_data["descriptions"].append("Generated Description Placeholder")
            return response_data
        else:
            print(f"Unexpected JSON structure for Google Search ads: {response_data}")
            return {"headlines": ["Error"]*15, "descriptions": ["Error"]*4} # Fallback
    except Exception as e:
        print(f"Error generating Google Search ad content: {e}")
        return {"headlines": ["Error generating headline"]*15, "descriptions": ["Error generating description"]*4}

def generate_google_display_ads(api_key, scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url):
    base_context = _build_base_context_prompt(scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url)
    system_prompt = "You are an expert Google Display Ads copywriter. Generate content as a JSON object."
    user_prompt = f"""
    {base_context}

    ### Task:
    Generate copy for Google Display Ads (Responsive Display Ad format).
    The ads should drive traffic towards "{lead_objective_type}" at {lead_objective_url}, or promote the downloadable asset if relevant ({downloadable_asset_url}).

    ### Ad Components:
    - Headlines: Create 5 unique short headlines. Each headline must be a maximum of 30 characters.
    - Descriptions: Create 5 unique descriptions. Each description must be a maximum of 90 characters.
    (Note: Google Display also uses Long Headlines (90 chars) and Business Name (25 chars), but we'll focus on these for now as per spec)

    ### Output Format:
    Return a JSON object with two keys: "headlines" (a list of 5 strings) and "descriptions" (a list of 5 strings).
    Example:
    {{
        "headlines": ["Short Headline 1 (<=30)", ... (5 total)],
        "descriptions": ["Description 1 (<=90)", ... (5 total)]
    }}
    Ensure all character limits are strictly followed.
    """
    try:
        response_data = _call_openai_api(api_key, system_prompt, user_prompt, expecting_json=True)
        if isinstance(response_data, dict) and "headlines" in response_data and "descriptions" in response_data:
            response_data["headlines"] = [h[:30] for h in response_data.get("headlines", [])][:5]
            response_data["descriptions"] = [d[:90] for d in response_data.get("descriptions", [])][:5]
            while len(response_data["headlines"]) < 5: response_data["headlines"].append("Generated Headline Placeholder")
            while len(response_data["descriptions"]) < 5: response_data["descriptions"].append("Generated Description Placeholder")
            return response_data
        else:
            print(f"Unexpected JSON structure for Google Display ads: {response_data}")
            return {"headlines": ["Error"]*5, "descriptions": ["Error"]*5} # Fallback
    except Exception as e:
        print(f"Error generating Google Display ad content: {e}")
        return {"headlines": ["Error generating headline"]*5, "descriptions": ["Error generating description"]*5}

def generate_reasoning_text(api_key, scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url):
    base_context = _build_base_context_prompt(scraped_data, additional_docs_text, lead_objective_type, lead_objective_url, downloadable_asset_url)
    system_prompt = "You are a marketing strategy analyst. Provide a concise reasoning statement."
    user_prompt = f"""
    {base_context}

    ### Task:
    You have just assisted in generating various marketing content pieces (emails, LinkedIn ads, Facebook ads, Google Search ads, Google Display ads) based on the information above.
    Please provide a brief reasoning statement (2-3 paragraphs, approx 150-250 words) explaining:
    1. How the scraped company information (e.g., USPs, target audience, tone of voice) was leveraged to tailor the generated content.
    2. How the user's marketing objectives (lead objective, URLs) influenced the messaging and calls to action in the content.
    3. Any general strategies or considerations applied during content generation for this specific client.

    Keep the tone professional and insightful, suitable for an internal consultancy tool. Do not output JSON, just the text.
    """
    try:
        reasoning = _call_openai_api(api_key, system_prompt, user_prompt, expecting_json=False)
        return reasoning if reasoning else "Error generating reasoning text."
    except Exception as e:
        print(f"Error generating reasoning text: {e}")
        return f"Error generating reasoning text: {e}"