# from datetime import datetime

# import processor.grader_report as gr
# import processor.semester_report as sr

# # source_6 = gr.GraderReport("Z:\Work\JAC\Project JARS\Playground GR - 6 Items.xlsm")
# # source_7 = gr.GraderReport("Z:\Work\JAC\Project JARS\Playground GR - 7 Items.xlsm")
# source_9 = gr.GraderReport("Z:\Work\JAC\Project JARS\Playground GR - 9 Items.xlsm")
# # source_10 = gr.GraderReport("Z:\Work\JAC\Project JARS\Playground GR - 10 Items.xlsm")

# # sr.Generator("Z:\Work\JAC\Project JARS\PG Test\Six", source_6).generate_all(autocorrect = False, force = True)
# # sr.Generator("Z:\Work\JAC\Project JARS\PG Test\Seven", source_7).generate_all(autocorrect = True, force = True)
# sr.Generator("Z:\Work\JAC\Project JARS\PG Test\\Nine", source_9, date = datetime(2001, 11, 23)).generate_all(autocorrect = False, force = True)
# # sr.Generator("Z:\Work\JAC\Project JARS\PG Test\Ten", source_10).generate_all(autocorrect = False, force = True)

# # from docx2pdf import convert
# # convert("Z:\Work\JAC\Project JARS\PG Test\\Nine")

# import processor.integrity as integrity
# integrous = integrity.verify_pdf("Z:\Work\JAC\Project JARS\PG Test\Six Int\Hara Yune.pdf", verbose = True)
# integrous = integrity.verify_pdf("Z:\Work\JAC\Project JARS\PG Test\Six Int\Raven Limadinata.pdf")
# print(integrous)

import customtkinter as ctk
import tkinter as tk
import moodle.database as db
import pandas as pd
import webview

moodle = db.Database()
data = moodle.query_from_file(r"resources\moodle_custom_queries\On-demand Parameter-based\Students whose attendance is below n%.sql", 
                              Session_Start_Date = "2023-12-01", Minimum_Percent = 100)
df = pd.DataFrame(data)
df.fillna("-", inplace = True)
data_html = df.to_html(index = False)

with open("gui/templates/database_viewer.html", "r") as f:
    html = f.read()
    html = html.replace("<!-- REPORT NAME -->", "Config Changes Log")
    html = html.replace("<!-- DATA -->", data_html)
    html = html.replace('class="dataframe"', 'id="datatable" class="table table-striped"')

# import webbrowser
# import tempfile

# with tempfile.NamedTemporaryFile("w", delete = False, suffix = ".html") as f:
#     f.write(html)
#     webbrowser.open(f.name)
    
import flask

app = flask.Flask(__name__, template_folder = "gui/templates/", static_folder = "gui/static/")

@app.route("/")
def index():
    return html

webview.create_window("JARS Report Processor", url = app, text_select = True, confirm_close = True, maximized = True)
webview.start(debug = True)