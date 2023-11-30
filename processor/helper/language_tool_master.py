import language_tool_python as ltp
import os

_tool = None

def check_java():
    import shutil
    if shutil.which("java"):
        return True
    else:
        print("It seems like you are trying to a feature from Language Tools. Languages Tools need Java but Java is not installed. Please install Java to use this feature.")
        return False    

def get_tool():
    """Returns a LanguageTool object."""
    print("[  ] Looking for LanguageTool instance…")

    if not check_java:
        return
    
    global _tool
    if _tool is None:
        print("[  ] LanguageTool instance not found. Creating new instance…")
        print(f"[  ] Looking for Language Tool server files in {os.environ.get('LTP_PATH')}…")
        
        if os.environ.get('LTP_PATH') is None:
            print("[  ] Language Tool server files not found. Downloading server files…")
        else:
            print("[OK] Server files downloaded.")

        _tool = ltp.LanguageTool('en-GB', config = {"cacheSize": 1000, "pipelineCaching": True, "maxCheckThreads": 20})
    
    print("[OK] LanguageTool instance found.")
    return _tool

def close_tool():
    """Closes the LanguageTool object."""
    print("[  ] Closing LanguageTool instance…")
    global _tool
    if _tool is not None:
        _tool.close()
        _tool = None
    
    print("[OK] LanguageTool instance closed.")