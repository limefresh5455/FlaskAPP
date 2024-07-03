from flask import Flask, render_template,request,jsonify
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
from matplotlib import patches
from io import BytesIO
import base64
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import plotly.graph_objs as go
import plotly.io as pio

from flask_cors import CORS



plt.switch_backend('Agg')
app = Flask(__name__, static_folder='static')
CORS(app)

import random
import string

def generate_random_string(length):
    letters_and_digits = string.ascii_letters + string.digits
    return ''.join(random.choices(letters_and_digits, k=length))

def projectType():
    df_sheet_name_projectType = pd.read_excel('Mockup_Dashboard_2024-05-30.xlsx', sheet_name='Home page metrics')
    barchartType = df_sheet_name_projectType['Bar chart project type']
    ProjectTypeCount = df_sheet_name_projectType['Count']

# Create the donut plot
    fig, ax = plt.subplots()
    size = 0.3

# Pie chart, where the slices will be ordered and plotted counter-clockwise:
    labels = barchartType
    sizes = ProjectTypeCount
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728']  # Define your own colors if necessary

# Outer ring
    wedges, texts, autotexts = ax.pie(sizes, radius=1, labels=labels, colors=colors, autopct='%1.1f%%',
                                 wedgeprops=dict(width=size, edgecolor='w'))

# Add data count labels
    for i, a in enumerate(autotexts):
        a.set_text(f'{sizes[i]}')

# Inner ring
    ax.pie([sum(sizes)], radius=1-size, labels=[sum(sizes)], labeldistance=0.35,
       wedgeprops=dict(width=size, edgecolor='w'))

# Draw center circle for a donut chart
    centre_circle = plt.Circle((0,0),0.70,fc='white')
    fig.gca().add_artist(centre_circle)

# Equal aspect ratio ensures that pie is drawn as a circle.
    ax.axis('equal')  

    plt.title('Project Type', pad=20)
    plt.savefig("static/output.jpg")
    #plt.show()
##########################################################

def funnelChart():
    df_sheet_name_funnel = pd.read_excel("Mockup_Dashboard_2024-05-30.xlsx", sheet_name='Funnel')
    funnel_data = df_sheet_name_funnel['Conversions']
    project_status_label = df_sheet_name_funnel['Project Status']

    labels = project_status_label
    sizes = funnel_data  # Widths of the bars
    colors = ['#004c6d', '#00a2ce', '#40a828', '#f37a20']  # Colors for the bars

    fig, ax = plt.subplots()

# Creating the bars
    for i in range(len(sizes)):
        ax.barh(y=i, width=sizes[i], color=colors[i], edgecolor='black', height=0.8)
    
        ax.text(sizes[i] / 2, i, sizes[i], ha='center', va='center', color='black', fontsize=12)

# Setting the y-axis labels
    ax.set_yticks(range(len(sizes)))
    ax.set_yticklabels(labels)

# Invert y-axis to have the funnel shape
    ax.invert_yaxis()

# Removing spines and ticks for better aesthetics
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.xaxis.set_ticks_position('none')
    ax.yaxis.set_ticks_position('none')

# Adding a title
    plt.title('Funnel Chart')
    plt.savefig("static/funnel.jpg")
# Show the plot
    #plt.show()

##################################################################
def financial_activity():
    df_sheet_name5 = pd.read_excel('Mockup_Dashboard_2024-05-30.xlsx', sheet_name='Accruals')
    acc_projetid = df_sheet_name5['Projects']
    acc_cashout = round(df_sheet_name5['C/LT']*100,2)
    acc_accrual = round(df_sheet_name5['AtD / LT']*100,2)
    data = sorted(zip(acc_projetid, acc_accrual, acc_cashout))

# Unpack sorted data
    projects, accrual, cashout = zip(*data)

# Plotting
    fig, ax = plt.subplots(figsize=(12, 8))

# Plot each bar as 100% (i.e., full length)

    for i, (proj, acc, cash) in enumerate(data):
        ax.barh(i, 100, color='lightgray')
        ax.barh(i, cash, color='pink')
        ax.plot([acc, acc], [i - 0.4, i + 0.4], color='black', linewidth=3)
        ax.text(acc + 1, i, f'{acc:.2f}%', va='center', ha='left', fontsize=10, color='black')
    #ax.text(cash + 1, i, f'{cash:.2f}%', va='center', ha='left', fontsize=10, color='blue')


# Adding labels and title
    ax.set_yticks(np.arange(len(data)))
    ax.set_yticklabels(projects)
    ax.set_xlabel('Percentage')
    ax.set_title('Top 5 Accrual-Cashout Deltas')
    ax.invert_yaxis()

# Displaying the chart
    plt.savefig("static/cashout_vs_accrual.jpg")
    #plt.show()

#####################################################
def piechart_calc_num(num):
  if int(num) == 1:
    return [0.25,0.75]
  elif int(num) == 2:
    return [0.5,0.5]
  elif int(num) == 3:
    return [0.75,0.25]
  elif int(num) == 4:
    return [1,0]
  else:
    return [1,0]

def pichart_create(data):
    import matplotlib.gridspec as gridspec
    fig = plt.figure(figsize=(10, 5))
    gs = gridspec.GridSpec(len(data), 4, wspace=0.4, hspace=0.4)

    row = 0
    titlepiechart = ['Q1','Q2','Q3','Q4']
    for project, ratios in data.items():
        for col in range(4):
            ax = fig.add_subplot(gs[row, col])
            data_r = ratios[col]
            colors = ['black', 'white']
            ax.pie(data_r, colors=colors, startangle=90, counterclock=False, wedgeprops={'edgecolor': 'black'})
            ax.set_title(f'{titlepiechart[row]}', fontsize=7, fontweight='bold')
        row += 1

# Save the figure
    random_string = generate_random_string(10)
    plt.savefig('static/'+random_string+'.png')
    return random_string

#####################################################
def budget_cashout_accrual_chart(projectid,totalbd,accrualbd,cashoutbd,lastdate):
    fig, ax = plt.subplots(figsize=(12, 6))

# Plotting the left-hand side bar chart for last invoice day
    #y_pos = np.arange(len([lastdate]))
    ltdate = -abs(lastdate)
    ax.barh(1, ltdate, color='orange', alpha=0.6, label='Last Invoice Day')
    ax.text(ltdate + 1, 0, str(lastdate), color='black', ha='right')

# Plotting the right-hand side bar chart for cashout and accruals
    #for i in range(len(projectid)):
    ax.barh(0, totalbd, color='grey', alpha=0.1, left=10)  # Shifted to the right by 150
    ax.barh(0, cashoutbd, color='black', alpha=0.7, left=10, label='Cashout' )
    ax.barh(0, accrualbd, left=cashoutbd + 10, color='lightblue', alpha=0.6, label='Accruals' )
    ax.plot([accrualbd+cashoutbd, accrualbd+cashoutbd], [0 - 0.4, 0 + 0.4], color='black', linewidth=3)

# Adding project labels and cashout values
    # for i in range(len(projectid)):
    ax.text(0, 0, projectid, color='blue', va='center', ha='right')
    ax.text(totalbd + 151, 0, f"${accrualbd:,}", color='black', va='center')

# Customizing the plot
    ax.axis('off')
    # ax.set_yticks(1)
    # ax.set_yticklabels([])
    ax.set_xlabel('Days / Percentage')
    ax.set_title('Budget-Accrual-Cashout-LastInvoiceDate')

    ax.legend(loc='lower right')

    plt.tight_layout()
    random_string = generate_random_string(10)
    plt.savefig('static/'+random_string+'.png')
    return random_string
    # plt.show()

####################################################
def get_project_data(project_id):
    filtered_unique = df_unique[df_unique['Project ID'] == project_id]
    if filtered_unique.empty:
        return f"No data found for Project ID {project_id}"
    
    proj_id = filtered_unique.iloc[0]['Project ID']
    proj_type = filtered_unique.iloc[0]['Project Type']
    total_budget = filtered_unique.iloc[0]['Budget']
    use_budget = result_budget[result_budget['Project ID'] == proj_id]['Used Budget'].values[0]
    total_datediff = result_schedule[result_schedule['Project ID'] == proj_id]['DateDiff'].values[0]
    avg_budget = round((use_budget / total_budget), 2)
    
    # Placeholder for overall_bin_dict data (replace with actual data)
    overall_bin_dict = {
        "Project 01": ["Bin1", "Bin2", "Bin3", "Bin4", "Bin5", "Bin6"],
        "Project 02": ["BinA", "BinB", "BinC", "BinD", "BinE", "BinF"],
        # Add other projects as needed
    }
    
    z = overall_bin_dict.get(proj_id, ["N/A"] * 6)
    combined_data = [avg_budget, total_datediff] + z

    return combined_data


###################################
@app.route('/')
def index():
    projectType()
    funnelChart()
    file_path="Mockup_Dashboard_2024-05-30.xlsx"
    df_sheet2= pd.read_excel(file_path, sheet_name='Pivot Tables')
    completed = df_sheet2.iloc[1, 1]
    delayed = df_sheet2.iloc[2, 1]
    ongoing = df_sheet2.iloc[3, 1]

    feedbacka = df_sheet2.iloc[2, 18]
    feedbackb = df_sheet2.iloc[3, 18]
    feedbackc = round(df_sheet2.iloc[4, 18],3)


    df_sheet= pd.read_excel("Mockup_Dashboard_2024-05-30.xlsx", sheet_name='Data')
    accrual_amount = df_sheet['Accrual Amount'].sum()
    cashout = round(df_sheet['Used Budget'].sum(),2) # usebudget
    totalbudget = df_sheet['Budget'].sum()

    hoursused = df_sheet['Hours Used'].sum()
    estimatedHours = df_sheet['Hours Estimated (Project Total)'].sum()
    
    return render_template('index.html',totalbudget=totalbudget,cashout=cashout,accrual=accrual_amount,completed=completed,delayed=delayed,ongoing=ongoing,hoursused=hoursused,estimatedHours=estimatedHours,a=feedbacka,b=feedbackb,c=feedbackc)

@app.route('/financial')
def financial():
    financial_activity()
    return render_template('finance.html')

@app.route('/project')
def project_actibity():
    df_sheet= pd.read_excel("Mockup_Dashboard_2024-05-30.xlsx", sheet_name='Data')

    projectid=df_sheet['Project ID'].unique()
    unique_list = []
    for x in projectid:
        if x not in unique_list:
            unique_list.append(x)
    print(unique_list)
    return render_template('project.html',ids=unique_list)

@app.route('/process')
def process():
    mdata = request.args.get('data')

    # Load the Excel file
    file_path = "Mockup_Dashboard_2024-05-30.xlsx"

    # Load the 'Project Status' sheet and filter the data
    df_sheet_name2 = pd.read_excel(file_path, sheet_name='Project Status')
    range1 = df_sheet_name2.iloc[12:34, 0:2]
    range1.columns = ['Project ID', 'Status']
    filtered_data = range1[range1['Project ID'] == mdata]

    # Load the 'ProjectData Test' sheet and filter the data
    df_sheet_name3 = pd.read_excel(file_path, sheet_name='ProjectData Test')
    filtered_data2 = df_sheet_name3[df_sheet_name3['Projects'] == mdata]

    if filtered_data.empty or filtered_data2.empty:
        return jsonify({'result': 'No data found for the given Project ID'})

    projects = filtered_data2['Projects'].values[0]
    overall = filtered_data2['OVERALL'].values[0]
    bin_data = filtered_data2['Bin'].values[0]
    total = filtered_data2['Total Possible'].values[0]
    q1 = piechart_calc_num(pd.to_numeric(filtered_data2['Q1'], errors='coerce').fillna(total).values[0])
    q2 = piechart_calc_num(pd.to_numeric(filtered_data2['Q2'], errors='coerce').fillna(total).values[0])
    q3 = piechart_calc_num(pd.to_numeric(filtered_data2['Q3'], errors='coerce').fillna(total).values[0])
    q4 = piechart_calc_num(pd.to_numeric(filtered_data2['Q4'], errors='coerce').fillna(total).values[0])
    status = filtered_data['Status'].values[0]
    data={projects:[q1,q2,q3,q4]}
    piechartname = pichart_create(data)
    # Concatenate the results into a single string
    # Load the Excel sheet
###############################################
    df_sheet_name = pd.read_excel(file_path, sheet_name='Data')

# Remove duplicates based on 'Project ID' and 'Project Type'
    df_unique = df_sheet_name.drop_duplicates(subset=['Project ID', 'Project Type'])

# Calculate the sum of 'Accrual amount' for each 'Project ID'
    accrual_amount = df_sheet_name.groupby('Project ID')['Accrual Amount'].sum().reset_index()
    accrual_amount.columns = ['Project ID', 'Accrual Amount']

# Calculate the sum of 'Used Budget' for each 'Project ID'
    result_budget = df_sheet_name.groupby('Project ID')['Used Budget'].sum().reset_index()
    result_budget.columns = ['Project ID', 'Used Budget']

# Calculate the sum of 'DateDiff' for each 'Project ID'
    result_schedule = df_sheet_name.groupby('Project ID')['DateDiff'].sum().reset_index()
    result_schedule.columns = ['Project ID', 'DateDiff']
    project_id = mdata
    filtered_unique = df_unique[df_unique['Project ID'] == project_id]
    if filtered_unique.empty:
        return f"No data found for Project ID {project_id}"
    
    proj_id = filtered_unique.iloc[0]['Project ID']
    #proj_type = filtered_unique.iloc[0]['Project Type']
    total_budget = filtered_unique.iloc[0]['Budget']
    accrual_amount_budget = accrual_amount[accrual_amount['Project ID'] == proj_id]['Accrual Amount'].values[0]
    use_budget = result_budget[result_budget['Project ID'] == proj_id]['Used Budget'].values[0]
    total_datediff = result_schedule[result_schedule['Project ID'] == proj_id]['DateDiff'].values[0]
    avg_budget = round((use_budget / total_budget), 2)
    #print(avg_budget)
    casg_acc_barchart = budget_cashout_accrual_chart(proj_id,total_budget,accrual_amount_budget,use_budget,total_datediff)
################################################
    
    res = {'Projectid':proj_id,
    'status':status,
    'img':piechartname+".png",
    'img2':casg_acc_barchart+".png",
    'overall':int(overall),
    'scope':int(bin_data),
    'schedule':int(total_datediff),
    'avg_budget':str(avg_budget),
    'accrual_amount':int(accrual_amount_budget)
    }
    # print(res)
    # res=f"{res}"
    return jsonify({'result': res})
##################################################
@app.route('/testdata', methods=['GET'])
def data():
    # Example JSON data
    # data = {
    #     'name': 'John Doe',
    #     'age': 30,
    #     'city': 'New York'
    # }
    df_sheet_name5 = pd.read_excel('Mockup_Dashboard_2024-05-30.xlsx', sheet_name='Accruals')
    acc_projetid = df_sheet_name5['Projects']
    acc_cashout = round(df_sheet_name5['C/LT']*100,2)
    acc_accrual = round(df_sheet_name5['AtD / LT']*100,2)
    totlbudget = df_sheet_name5['LT_cost']
    data = sorted(zip(acc_projetid, acc_accrual, acc_cashout,totlbudget))

# Unpack sorted data
    projects, accrual, cashout,total_budget = zip(*data)

    res = {"a": projects,"b":accrual,"c":cashout,"d":total_budget}
    return jsonify(res)

@app.route('/test')
def testdata():
    return render_template('portfolio.html')

data = {
    "message": "Hello from Flask!"
}

@app.route('/api/data', methods=['GET'])
def get_data():
    return jsonify(data)

@app.route('/api/data', methods=['POST'])
def post_data():
    new_data = request.get_json()
    return jsonify(new_data), 201

if __name__ == '__main__':

    app.run(debug=True)
