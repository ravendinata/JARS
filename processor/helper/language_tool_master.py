import language_tool_python as ltp

_tool = None

def get_tool():
    """Returns a LanguageTool object."""
    print("[  ] Looking for LanguageTool instance…")
    
    global _tool
    if _tool is None:
        print("[  ] LanguageTool instance not found. Creating new instance…")
        _tool = ltp.LanguageTool('en-UK', config = {"cacheSize": 1000, "pipelineCaching": True, "maxCheckThreads": 20})
    
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