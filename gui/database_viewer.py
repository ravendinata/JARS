import pandas as pd
import webview

@staticmethod
def TableViewer(data, window_title = "JARS Database Table Viewer", report_title = "Unnamed Report"):
    df = pd.DataFrame(data)
    df.fillna("-", inplace = True)
    data_html = df.to_html()

    with open("gui/database_viewer.html", "r") as f:
        html = f.read()
        html = html.replace("<!-- REPORT NAME -->", report_title)
        html = html.replace("<!-- DATA -->", data_html)
        html = html.replace('class="dataframe"', 'id="datatable" class="table table-striped"')

    webview.create_window(title = window_title, html = html, text_select = True, confirm_close = True, maximized = True)
    webview.start()