import pythoncom
from win32com.client import Dispatch, GetActiveObject
from win32com.client.dynamic import CDispatch

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

def check_word_status():
    """
    Check MS Word installation and running status.
    Returns tuple of (is_running: bool, office_version: str | None, error: str | None)
    """
    try:
        # Initialize COM in thread
        pythoncom.CoInitialize()
        
        # Check for running instances first using GetActiveObject
        word_is_open = False
        try:
            word = GetActiveObject("Word.Application")
            word_is_open = True
            print("  An open MS Word instance is detected. Will not close it.")
        except pythoncom.com_error:
            pass

        # Only create new instance if needed
        if not word_is_open:
            word = Dispatch("Word.Application")
            print("  No open MS Word instance detected. Will create new instance to check version.")

        if isinstance(word, CDispatch):
            office_version = word.Version
            if not word_is_open:
                word.Quit()
            print(f"  Microsoft Office {office_version} detected.")
            return word_is_open, office_version, None
        else:
            return word_is_open, None, "Failed to get Word.Application dispatch"
                
    except pythoncom.com_error as e:
        return word_is_open, None, f"COM Error: {str(e)}"
    except Exception as e:
        return word_is_open, None, str(e)
    finally:
        pythoncom.CoUninitialize()