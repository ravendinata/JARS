import flask
import pandas as pd
import webview

@staticmethod
def TableViewer(data, window_title = "JARS Database Table Viewer", report_title = "Unnamed Report"):
    """
    Opens a new database viwer window for viewing data in a table form.
    
    Args:
        data (list): The data to display in the table.
        window_title (str): The title of the window.
        report_title (str): The title of the report.
    """
    df = pd.DataFrame(data)
    df.fillna("-", inplace = True)
    data_html = df.to_html()

    with open("gui/templates/database_viewer.html", "r") as f:
        html = f.read()
        html = html.replace("<!-- REPORT NAME -->", report_title)
        html = html.replace("<!-- DATA -->", data_html)
        html = html.replace('class="dataframe"', 'id="datatable" class="table table-striped"')

    app = flask.Flask(__name__, template_folder = "templates/", static_folder = "static/")

    @app.route("/")
    def index():
        return html

    webview.create_window(title = window_title, url = app, text_select = True, confirm_close = True, maximized = True)
    webview.start()