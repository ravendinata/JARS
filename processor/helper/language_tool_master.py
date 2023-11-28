import language_tool_python as ltp

_tool = None

def get_tool():
    """Returns a LanguageTool object."""
    global _tool
    if _tool is None:
        _tool = ltp.LanguageTool('en-UK', config = {"cacheSize": 1000, "pipelineCaching": True, "maxCheckThreads": 20})
    return _tool

def close_tool():
    """Closes the LanguageTool object."""
    global _tool
    if _tool is not None:
        _tool.close()
        _tool = None