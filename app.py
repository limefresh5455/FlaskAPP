from flask import Flask, render_template,request,jsonify
import pandas as pd
# import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
# from matplotlib import patches 
# from matplotlib.patches import Wedge
# from io import BytesIO
# import base64
# import matplotlib.pyplot as plt
# from matplotlib.gridspec import GridSpec
# import plotly.graph_objs as go
# import plotly.io as pio
# from flask_cors import CORS

file_path = 'https://bpxai-my.sharepoint.com/personal/manas_shalgar_bpx_ai/_layouts/15/download.aspx?share=EW7IZR3VH_ZDp_rhgbv9GlQBowPIosXVpfUCXsAUTYNJ8Q'

app = Flask(__name__, static_folder='static')

def main_all_data():
    #file_path = 'Mockup_Dashboard_cleandatav2.xlsx'

# Read the data from each sheet
    df_sheet_name4 = pd.read_excel(file_path, sheet_name='Total Risk')
    df_sheet_name = pd.read_excel(file_path, sheet_name='Financials Data')
    df_sheet_name3 = pd.read_excel(file_path, sheet_name='PM Defined Status')
    df_sheet_name2 = pd.read_excel(file_path, sheet_name='Project Status')
    df_sheet_name5 = pd.read_excel(file_path, sheet_name='Issues')
    df_sheet_name6 = pd.read_excel(file_path, sheet_name='Project Data')

    project_dict = {}
# Process data from 'Total Risk' sheet
    for index, row in df_sheet_name4.iterrows():
        project_id = row['Project ID']
        project_dict[project_id] = {
            'Pace': row['Pace'],
            'Execution': row['Execution'],
            'Resources': row['Resources']
        }
    # Process data from 'Financials Data' sheet
    for index, row in df_sheet_name.iterrows():
        project_id = row['Project ID']
        if project_id not in project_dict:
            project_dict[project_id] = {}
        project_dict[project_id].update({
            'LT_Budget': row['LT_Budget'],
            'LT_Budget_Cashout': row['LT_Budget - Cashout'],
            'LT_Budget_Accrual': row['Cashout - Accrual'],
            'DateDiff': row['Accrual #'],
            'milestonecompletion_per':row['% Completed'],
            'actualprogess':row['Total Completed'],
            'Totalmilestone':row['Total Possible'],
            'Q1':row['Q1'],'Q2':(row['Q2']-row['Q1']),
            'Q3':(row['Q3']-row['Q2']),'Q4':(row['Q4']-row['Q3'])
        })

# Process data from 'PM Defined Status' sheet
    for index, row in df_sheet_name3.iterrows():
        project_id = row['Project ID']
        if project_id not in project_dict:
            project_dict[project_id] = {}
        project_dict[project_id].update({
            'Overall': row['OVERALL'],
            'Scope': row['Scope'],
            'Schedule': row['Schedule'],
            'Budget': row['Budget']
        })
    #print(f"After PM Defined Status: {project_id} -> {project_dict[project_id]}")

# Process final status data
    for index, row in df_sheet_name2.iterrows():
        project_id = row['Project ID']
        if project_id not in project_dict:
            project_dict[project_id] = {}
        project_dict[project_id].update({
        'Status': row['Project Status']
    })
    
#Process Issues Data
    for index, row in df_sheet_name5.iterrows():
        project_id=row['Project ID']
        if project_id not in project_dict:
            project_dict[project_id] = {}
        project_dict[project_id].update({
            'IssuesCount': row['Count'],
            'Issues': row['Lookup']
        })

#Project Data
    for index, row in df_sheet_name6.iterrows():
        project_id=row['Project ID']
        if project_id not in project_dict:
            project_dict[project_id] = {}
        project_dict[project_id].update({
            'startdate': row['Start Date'],
            'enddate': row['Due Date'],
            'type': row['Type']
        })
    
    return project_dict

def indexpage_top():
	df_budget_sheet = pd.read_excel(file_path, sheet_name='Gauges')
	budget_total_value = round(df_budget_sheet['Total Budget'].sum(), 2)
	budget_current_value = round(df_budget_sheet['Current Spend'].sum(), 2)
	status_current = round(df_budget_sheet['Current Status'].sum(),2)
	status_total = 100
	return budget_total_value,budget_current_value,status_current
####################################################################

def financedata_indexpage():
    df_sheet_name = pd.read_excel(file_path, sheet_name='Financials Data')
    financedata={}
    for index, row in df_sheet_name.iterrows():
        project_id=row['Project ID']
        if project_id not in financedata:
            financedata[project_id] = {}
        financedata[project_id].update({
            'projectid':row['Project ID'],
            'milestone':round((row['% Completed']*100),2),
            'recentdate':row['MostRecentDate'],

            })
    return financedata

################################################
#@app.route('/projectdata')
def indexpage_projectdata():
    df_sheet_name = pd.read_excel(file_path, sheet_name='Project Data')
    projectdata={}
    for index, row in df_sheet_name.iterrows():
        project_id=row['Project ID']
        if project_id not in projectdata:
            projectdata[project_id] = {}

        projectdata[project_id].update({
            'projectid':row['Project ID'],
            'status':row['Status'],
            'startdate':row['Start Date'],
            'duedate':row['Due Date'],
            'type':row['Type']
            })
    return projectdata

##############################################################################

@app.route('/funnel')
def funnel():
    df_sheet_name = pd.read_excel(file_path, sheet_name='Funnel')
    funnelchart={}
    for index, row in df_sheet_name.iterrows():
        project_id=row['Project Status']
        if project_id not in funnelchart:
            funnelchart[project_id] = {}
        funnelchart[project_id].update({
                'projectstatus':row['Project Status'],
                'conversion':row['Conversions']
            })
    return jsonify(funnelchart)

#############################################
def data_1():
    df_sheet_name2 = pd.read_excel(file_path, sheet_name='Financials Data')
    project_dict1={}
    for index, row in df_sheet_name2.iterrows():
        project_id=row['Project ID']
        if project_id not in project_dict1:
            project_dict1[project_id] = {}
        project_dict1[project_id].update({
            'ltbudget': row['LT_Budget'],
            'accrual': row['Accrual Unit'] * row['Accrual #'],
            'cashout': row['Cash Out']
        })
    return project_dict1

@app.route('/top5')
def get_data1():
    project_dict1 = data_1()
    sorted_accruals = sorted(project_dict1.items(), key=lambda x: x[1]['accrual'])
    top_5_accruals = sorted_accruals[-5:]
    top_5_output = [(proj_id,details['ltbudget'] ,details['accrual'], details['cashout']) for proj_id, details in top_5_accruals]
    return jsonify(top_5_output)

@app.route('/bottom5')
def get_data2():

    project_dict1 = data_1()
    sorted_accruals = sorted(project_dict1.items(), key=lambda x: x[1]['accrual'])
    bottom_5_accruals = sorted_accruals[:5]
    bottom_5_output = [(proj_id, details['ltbudget'] , details['accrual'], details['cashout']) for proj_id, details in bottom_5_accruals]

    return jsonify(bottom_5_output)
##########################################################
@app.route('/')
def index():
	#data=main_all_data()
    totalbudge,currentspend,currentstatus = indexpage_top()
    data1={'a':totalbudge,'b':currentspend,'c':currentstatus}
    projectdata=indexpage_projectdata()
    financedate = financedata_indexpage()
   
    return render_template('index.html',data1=data1,data2=projectdata,data3=financedate)

if __name__ == '__main__':

    app.run(debug=True)