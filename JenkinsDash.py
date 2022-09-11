import dash
from dash import dash_table
from dash import html
#import dash_html_components as html
from dash import dcc
import pandas as pd
import plotly.express as px
import requests as req
import json as js
import xlrd
import re
import datetime
from xlwt import Workbook
import xlsxwriter
import getpass


user = getpass.getuser()
loc = "/Users/"+user+"/jenkinjobs/jobsdata.xlsx"
filePath = "/Users/"+user+"/jenkinjobs/jobsreport.xlsx"
# To open Workbook
wb1 = xlsxwriter.Workbook(filePath)
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet1 = wb.sheet_by_index(1)
username = sheet1.cell_value(1, 0)
password = sheet1.cell_value(1, 1)
e = 0
joblink = 0
for j in range(sheet.ncols):
    t = sheet.cell_value(0, j)
    m1 = re.search('http://(.+?).<servername>', t)
    if m1:
        found1 = m1.group(1)
        #print found1
    sheet1 = wb1.add_worksheet(found1)
    sheet1.write(0, 0, 'Name')
    sheet1.write(0, 1, 'Status')
    sheet1.write(0, 2, 'Passed')
    sheet1.write(0, 3, 'Last Build Date')
    sheet1.write(0, 4, 'Build Stability')
    sheet1.write(0, 5, 'Job Link')
    for i in range(sheet.nrows):
        #print(sheet.cell_value(i, j))
        if i != 0 and sheet.cell_value(i, j):
            s = sheet.cell_value(0, j) + '/' +sheet.cell_value(i, j) + '/api/json?pretty=true'
            #print s
            r = req.get(s, auth=(username, password))
            #print(r.text)
            json_object = js.loads(r.text)
            for k in json_object["healthReport"]:
                if "Build stability" in k["description"]:
                    # print(k["description"])
                    x = k["description"]
                    pattern = 'Build stability: '
                    p = re.split(pattern, x)[1]

            url = json_object["builds"][0]['url'] + 'api/json?pretty=true'
            joblink = json_object["builds"][0]['url']
            #print("URL-----> " +url)
            m = re.search('/job/(.+?)/', url)
            if m:
                found = m.group(1)
                #print found
            r1 = req.get(url, auth=(username, password))
            json_object1 = js.loads(r1.text)
            url2 = json_object1["result"]
            if url2 is None:
                url2 = "INPROGRESS"
            for cs in json_object1["actions"]:
                k = cs.get("_class", '')
                if k == 'hudson.maven.reporters.SurefireAggregatedReport':
                    a = float(cs.get("failCount"))
                    b = float(cs.get("skipCount"))
                    c = float(cs.get("totalCount"))
                    #print a
                    #print b
                    #print c
                    d = c - (a + b)
                    e = int(float(d / c) * 100)
                    #print d
                    #print e
            #print url2
            url3 = json_object1["timestamp"]
            date_crop = url3 / 1000;
            date_time = datetime.datetime.fromtimestamp(date_crop)
            datetime_str = date_time.strftime("%d-%m-%Y")
            # print datetime_str
            sheet1.write(i, 0, found)
            sheet1.write(i, 1, url2)
            sheet1.write(i, 2, e)
            sheet1.write(i, 3, datetime_str)
            sheet1.write(i, 4, p)
            sheet1.write(i, 5, joblink)
wb1.close()


app = dash.Dash(__name__,title='Jenkins Job Analytics')
all_sheets_df = pd.read_excel(filePath, sheet_name=None)
#print(all_sheets_df)
print("Sheets in the jobdata.xlsx.")
print("Please wait .... Grabbing the job's details from jenkins....")
print(all_sheets_df.keys())
all_sheets = all_sheets_df.keys()
sheet_list = []
all_dfs = []

for sheet in all_sheets_df.keys():
    df1 = pd.read_excel(filePath, sheet)
    #joblink = "http://" + sheet + ".<server>.org:8080/job/" + df1['Name']
    #df1['Job Link'] = joblink
    # format dataframe column of urls so that it displays as hyperlink
    def display_links(df1):
        links = df1['Job Link'].to_list()
        rows = []
        for x in links:
            link = '[Job Link](' + str(x) + ')'
            rows.append(link)
        return rows

    df1['Job Link'] = display_links(df1)
    dftable = [
        dash_table.DataTable(
            columns=[{"name": i, "id": i,'presentation':'markdown'} for i in df1.columns],
            data=df1.to_dict('records'),
            page_action='none',
            filter_action="native",
            style_table={'overflowX': 'auto','overflowY': 'auto','width': '80%','margin-left': 'auto','margin-right': 'auto','height': '400px'},
            style_header={ 'font_family': 'verdana','backgroundColor': '#1e4569','font_size': '16px', 'fontWeight': 'bold', 'color': 'white', 'height': '50px','textAlign': 'center'},
            style_cell={'whiteSpace': 'normal','height': '5px','textAlign': 'left'},
            # style_cell_conditional=[
            #     {'if': {'column_id': 'Status'},
            #      'width': '100px'},
            #     {'if': {'column_id': 'Status'},
            #      'width': '100px'},
            #     {'if': {'column_id': 'Passed'},
            #      'width': '80px'},
            #     {'if': {'column_id': 'Job Link'},
            #      'width': '80px'},
            #     {'if': {'column_id': 'Last Build Date'},
            #      'width': '80px'},
            # ],

            style_data_conditional=[
                {   'color': 'black', },
                {
                    'if': {
                        'filter_query': '{Status} = "Success" ||  {Status} = "SUCCESS"',
                        'column_id': 'Status'
                    },
                    'color': 'green',
                    'fontWeight': 'bold'
                },
                {
                    'if': {
                        'filter_query': '{Status} = "Failure" ||  {Status} = "FAILURE"',
                        'column_id': 'Status'
                    },
                    'color': 'red',
                    'fontWeight': 'bold'
                },
                {
                    'if': {
                        'filter_query': '{Status} = "Unstable" ||  {Status} = "UNSTABLE"',
                        'column_id': 'Status'
                    },
                    'color': 'blue',
                    'fontWeight': 'bold'
                },
                {
                    'if': {
                        'filter_query': '{Status} = "INPROGRESS" ||  {Status} = "Inprogress"',
                        'column_id': 'Status'
                    },
                    'color': 'grey',
                    'fontWeight': 'bold'
                },
                {
                    'if': {
                        'filter_query': '{{{}}} is blank'.format('Status'),
                        'column_id': 'Status'
                    },
                    'color': 'grey',
                    'fontWeight': 'bold'
                },
            ]
        )
    ]

  #  sheet_list.append(dcc.Tab(dftable,label=sheet,id=sheet,value=sheet,selected_className='custom-tab--selected'))
    sheet_list.append(dcc.Tab(dftable, label=sheet, id=sheet, value=sheet,
                              style={'fontWeight': 'bold',
                              },
                              selected_style = {
                                  'borderTop': '5px solid #e36209',
                                  'fontWeight': 'bold',
                              },))

    for tab_name, df in all_sheets_df.items():
        df['sheet_name'] = tab_name
        all_dfs.append(df)
        final_df = pd.concat(all_dfs, ignore_index=True)
        fig = px.pie(final_df.to_dict('records'), names="Status", hole=.5, title='<b>Overall Health of Jobs</b>', color='Status',
                 color_discrete_map={"ABORTED": "grey",
                                     "SUCCESS": "green",
                                     "FAILURE": "red",
                                     "UNSTABLE": "blue",
                                     "INPROGRESS": "grey",
                                     })
    pieChart = dcc.Graph(id='pie', figure=fig)

app.layout = html.Div([
    html.Div(
        html.H1(
            children="Jenkins Job's Summary from various servers",
            style={
                'textAlign': 'center',
                'color': 'white',
                'padding-top':'20px'
            }
        ),
        style = {'backgroundColor': 'rgb(127, 150, 200)','height': '90px'},
    ),
    html.Div(children=html.Div([
        html.H2(
            children ="Tabular representation of the jenkins jobs from various servers. Click on tabs to view job's status",
            style={
                'textAlign': 'center',
            }),
    ])),
    dcc.Tabs(sheet_list,
             id="tabs-with-classes",
             value='tab-1',
             colors={
                 "border": "white",
                 "primary": "grey",
                 "background": "silver"
             },

    ),
    html.Div(id="tab-content", className="p-4"),
    html.Div(pieChart,style={"height": "50%", "width": "95%"}),

    html.Div(
        html.H4(children="Developed by: Manikanta B & Shwetha BM",
                style={'textAlign': 'bottom-left',
                       }),
        ),

])

if __name__ == "__main__":
    app.run_server(debug=True)
