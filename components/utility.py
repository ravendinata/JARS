import google.generativeai as genai
from termcolor import colored

import config

def validate_genai_api_key():
    """
    Validates the GenAI API key.
    
    Returns:
        bool: Whether the API key is valid.
    """
    try:
        genai.configure(api_key = config.get_config("genai_api_key"))
        model = genai.GenerativeModel('gemini-pro')
        model.generate_content("This is a test.")
        return True
    except Exception as e:
        print(colored(f"(!) Error: {e}", "red"))
        return False