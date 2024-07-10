from flask import Flask, render_template,request,jsonify
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
from matplotlib import patches 
from matplotlib.patches import Wedge
from io import BytesIO
import base64
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import plotly.graph_objs as go
import plotly.io as pio
from flask_cors import CORS
from datetime import date


plt.switch_backend('Agg')
app = Flask(__name__, static_folder='static')
CORS(app)

import random
import string

file = 'Mockup_Dashboard_cleandatav2.xlsx'

def generate_random_string(length):
    letters_and_digits = string.ascii_letters + string.digits
    return ''.join(random.choices(letters_and_digits, k=length))

def projectType():
    df_sheet_name_projectType = pd.read_excel(file, sheet_name='Project Types')
    barchartType = df_sheet_name_projectType['Bar chart type']
    ProjectTypeCount = df_sheet_name_projectType['Count']

# Create the donut plot
    fig, ax = plt.subplots()
    size = 0.3
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

    # Outer ring
    wedges, texts, autotexts = ax.pie(ProjectTypeCount, radius=1, labels=barchartType, colors=colors, autopct='',
                                      wedgeprops=dict(width=size, edgecolor='w'))

    # Add data count labels inside the slices
    for i, wedge in enumerate(wedges):
        ang = (wedge.theta2 - wedge.theta1)/2. + wedge.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        ax.text(x * 0.85, y * 0.85, f'{ProjectTypeCount[i]}', ha='center', va='center',
                color='white' if colors[i] in ['#1f77b4', '#d62728', '#9467bd'] else 'black',
                weight='bold')

    # Inner ring
    total = sum(ProjectTypeCount)
    ax.pie([total], radius=1-size, labels=[total], labeldistance=-0,
           wedgeprops=dict(width=size, edgecolor='w'), colors=['white'], textprops=dict(color='black', weight='bold'))

    # Draw center circle for a donut chart
    centre_circle = plt.Circle((0,0),0.30,fc='white')
    fig.gca().add_artist(centre_circle)

    # Equal aspect ratio ensures that pie is drawn as a circle.
    ax.axis('equal')  

    plt.title('Project Type', pad=20)
    plt.savefig("static/output.jpg")
    #plt.show()
##########################################################

import pandas as pd
import matplotlib.pyplot as plt

def funnelChart():
    df_sheet_name_funnel = pd.read_excel("Mockup_Dashboard_cleandatav2.xlsx", sheet_name='Funnel')
    funnel_data = df_sheet_name_funnel['Conversions']
    project_status_label = df_sheet_name_funnel['Project Status']

    labels = project_status_label
    sizes = funnel_data  # Widths of the bars
    colors = ['#004c6d', '#00a2ce', '#40a828', '#f37a20']  # Colors for the bars

    fig, ax = plt.subplots()

    # Creating the bars centered to look like a funnel
    for i in range(len(sizes)):
        left = (max(sizes) - sizes[i]) / 2
        ax.barh(y=i, width=sizes[i], left=left, color=colors[i], edgecolor='black', height=0.8)
        ax.text(max(sizes) / 2, i, sizes[i], ha='center', va='center', color='black', fontsize=12)

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
    ax.xaxis.set_visible(False)

    # Adding a title
    plt.title('Funnel Chart')
    plt.savefig("static/funnel.jpg")
    # plt.show()

##################################################################
def financial_activity():
    df_sheet_name5 = pd.read_excel(file, sheet_name='Financials Data')
    acc_projectid = df_sheet_name5['Project ID']
    lt_budget = df_sheet_name5['LT_Budget']
    acc_cashout = df_sheet_name5['Cash Out']
    acc_accrual = df_sheet_name5['Accrual to Date']
    
    # Combine data into a list of tuples and filter out cases where accrual to date is 0
    data = list(zip(acc_projectid, lt_budget, acc_accrual, acc_cashout))
    data = [entry for entry in data if entry[2] > 0]
    data = sorted(data, key=lambda x: abs(x[2] - x[3]), reverse=True)
    
    # Select the top 5 highest deltas
    top_5_highest = data[:5]
    
    # Select the top 5 lowest deltas
    top_5_lowest = sorted(data, key=lambda x: abs(x[2] - x[3]))[:5]
    
    # Unpack all data for plotting all projects
    acc_projectid, lt_budget_all, acc_accrual, acc_cashout = zip(*data)
    
    # Determine the maximum budget for setting the bar lengths
    max_budget = max(lt_budget_all)
    
    # Plotting for top 5 highest deltas
    fig, ax = plt.subplots(figsize=(12, 8))
    for i, (proj, budget, acc, cash) in enumerate(top_5_highest):
        ax.barh(i, max_budget, color='lightgray')
        ax.barh(i, cash, color='pink')
        ax.plot([acc, acc], [i - 0.4, i + 0.4], color='black', linewidth=3)
        ax.text(cash + 3, i, f'${cash:,.2f}', va='center', ha='left', fontsize=10, color='black')
    ax.set_yticks(np.arange(len(top_5_highest)))
    ax.set_yticklabels([proj for proj, _, _, _ in top_5_highest])
    ax.set_xlabel('Amount ($)')
    ax.set_title('Top 5 Highest Cashout-Accrual Deltas')
    ax.invert_yaxis()
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.savefig("static/cashout_vs_accrual_top5_highest.jpg")
    
    # Plotting for top 5 lowest deltas
    fig, ax = plt.subplots(figsize=(12, 8))
    for i, (proj, budget, acc, cash) in enumerate(top_5_lowest):
        ax.barh(i, max_budget, color='lightgray')
        ax.barh(i, cash, color='pink')
        ax.plot([acc, acc], [i - 0.4, i + 0.4], color='black', linewidth=3)
        ax.text(cash + 3, i, f'${cash:,.2f}', va='center', ha='left', fontsize=10, color='black')
    ax.set_yticks(np.arange(len(top_5_lowest)))
    ax.set_yticklabels([proj for proj, _, _, _ in top_5_lowest])
    ax.set_xlabel('Amount ($)')
    ax.set_title('Top 5 Lowest Cashout-Accrual Deltas')
    ax.invert_yaxis()
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.savefig("static/cashout_vs_accrual_top5_lowest.jpg")
    
    # Plotting for all projects sorted by deltas
    fig, ax = plt.subplots(figsize=(12, 16))
    for i, (proj, budget, acc, cash) in enumerate(data):
        ax.barh(i, max_budget, color='lightgray')
        ax.barh(i, cash, color='pink')
        ax.plot([acc, acc], [i - 0.4, i + 0.4], color='black', linewidth=3)
        ax.text(cash + 3, i, f'${cash:,.2f}', va='center', ha='left', fontsize=10, color='black')
    ax.set_yticks(np.arange(len(data)))
    ax.set_yticklabels([proj for proj, _, _, _ in data])
    ax.set_xlabel('Amount ($)')
    ax.set_title('All Projects Cashout-Accrual Deltas')
    ax.invert_yaxis()
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.savefig("static/cashout_vs_accrual_all.jpg")
    # plt.show()

#####################################################
def progress_chart():
    # Load data
    df_sheet_name_project = pd.read_excel(file, sheet_name='Project Data')
    df_sheet_name_milestones = pd.read_excel(file, sheet_name='Financials Data')

    project_id = df_sheet_name_project['Project ID']
    start_date = pd.to_datetime(df_sheet_name_project['Start Date'])
    due_date = pd.to_datetime(df_sheet_name_project['Due Date'])
    today = pd.Timestamp(date.today())
    num_milestones_completed = df_sheet_name_milestones['% Completed']
    max_milestones = 1  # 100%

    # Calculate progress
    active_duration = (today - start_date).apply(lambda x: x.days)
    total_duration = (due_date - start_date).apply(lambda x:x.days)
    progress = active_duration / total_duration

    # Combine data into a list of tuples and filter out cases where progress is 0
    data = list(zip(project_id, progress, num_milestones_completed))
    data = [entry for entry in data if entry[1] > 0]
    data = sorted(data, key=lambda x: abs(x[1] - x[2]), reverse=True)

    # Select the top 5 highest deltas
    top_5_highest = data[:5]

    # Select the top 5 lowest deltas
    top_5_lowest = sorted(data, key=lambda x: abs(x[1] - x[2]))[:5]

    # Unpack all data for plotting all projects
    project_id, progress_all, num_milestones_completed = zip(*data)

    # Determine the maximum progress for setting the bar lengths
    max_progress = max(progress_all)

    # Plotting for top 5 highest deltas
    fig, ax = plt.subplots(figsize=(12, 8))
    for i, (proj, prog, milestones) in enumerate(top_5_highest):
        ax.barh(i, max_progress, color='lightgray')
        ax.barh(i, milestones, color='green')
        ax.plot([prog, prog], [i - 0.4, i + 0.4], color='black', linewidth=3)
        ax.text(milestones + 0.03, i, f'{milestones * 100:.2f}%', va='center', ha='left', fontsize=10, color='black')
    ax.set_yticks(np.arange(len(top_5_highest)))
    ax.set_yticklabels([proj for proj, _, _ in top_5_highest])
    ax.set_xlabel('Progress (%)')
    ax.set_title('Top 5 Highest Progress-Milestones Deltas')
    ax.invert_yaxis()
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.savefig("static/progress_vs_milestones_top5_highest.jpg")

    # Plotting for top 5 lowest deltas
    fig, ax = plt.subplots(figsize=(12, 8))
    for i, (proj, prog, milestones) in enumerate(top_5_lowest):
        ax.barh(i, max_progress, color='lightgray')
        ax.barh(i, milestones, color='green')
        ax.plot([prog, prog], [i - 0.4, i + 0.4], color='black', linewidth=3)
        ax.text(milestones + 0.03, i, f'{milestones * 100:.2f}%', va='center', ha='left', fontsize=10, color='black')
    ax.set_yticks(np.arange(len(top_5_lowest)))
    ax.set_yticklabels([proj for proj, _, _ in top_5_lowest])
    ax.set_xlabel('Progress (%)')
    ax.set_title('Top 5 Lowest Progress-Milestones Deltas')
    ax.invert_yaxis()
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.savefig("static/progress_vs_milestones_top5_lowest.jpg")

    # Plotting for all projects sorted by deltas
    fig, ax = plt.subplots(figsize=(12, 16))
    for i, (proj, prog, milestones) in enumerate(data):
        ax.barh(i, max_progress, color='lightgray')
        ax.barh(i, milestones, color='green')
        ax.plot([prog, prog], [i - 0.4, i + 0.4], color='black', linewidth=3)
        ax.text(milestones + 0.03, i, f'{milestones * 100:.2f}%', va='center', ha='left', fontsize=10, color='black')
    ax.set_yticks(np.arange(len(data)))
    ax.set_yticklabels([proj for proj, _, _ in data])
    ax.set_xlabel('Progress (%)')
    ax.set_title('All Projects Progress-Milestones Deltas')
    ax.invert_yaxis()
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.savefig("static/progress_vs_milestones_all.jpg")
    # plt.show()
        

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
def gauge_chart():
    # Data
    df_budget_sheet = pd.read_excel(file, sheet_name='Gauges')
    budget_total_value = round(df_budget_sheet['Total Budget'].sum(), 2)
    budget_current_value = round(df_budget_sheet['Current Spend'].sum(), 2)
    status_current = round(df_budget_sheet['Current Status'].sum() / df_budget_sheet['Total Status'].sum() * 100, 2)
    status_total = 100

    def gauge(current_value, total_value, word, colors=['#e15759', '#fecd57', '#55a868']):
        if total_value != 0:
            arrow = current_value / total_value
        else:
            arrow = 0.5
        
        # Draw the sectors
        fig, ax = plt.subplots()
        ang_range = np.linspace(0, np.pi, len(colors) + 1)  # Angles for sectors
        radius = 2
        for i, (angle, color) in enumerate(zip(ang_range[:-1], colors)):
            ax.add_patch(Wedge((0., 0.), radius, np.degrees(angle), np.degrees(ang_range[i+1]), facecolor=color, edgecolor='white'))

        # Draw the arrow
        arrow_angle = np.degrees((1 - arrow) * np.pi)  # Adjust arrow to point correctly
        dx = np.cos(arrow_angle * np.pi / 180) * radius
        dy = np.sin(arrow_angle * np.pi / 180) * radius
        ax.arrow(0, 0, dx, dy, width=0.1, head_width=0.3, head_length=0.3, fc='k', ec='k')

        # Annotate the min, max, and current value
        ax.text(-radius, -radius * 0.2, '$0', ha='center', va='center', fontsize=12)
        if "budget" in word.lower():
            ax.text(radius, -radius * 0.2, f'${total_value:,}', ha='center', va='center', fontsize=12)
        else:
            ax.text(radius, -radius * 0.2, f'{total_value:,}%', ha='center', va='center', fontsize=12)
        ax.text(np.cos((1 - arrow) * np.pi) * (radius + 0.4), np.sin((1 - arrow) * np.pi) * (radius + 0.4), f'${current_value:,}', ha='center', va='center', fontsize=12)

        ax.xaxis.set_visible(False)
        ax.yaxis.set_visible(False)
        plt.axis('off')
        #plt.savefig('static/guageCHART.png')
        plt.savefig('static/' + word + '.png')
        # plt.show()

    gauge(budget_current_value, budget_total_value, 'Budget Spent')
    gauge(status_current, status_total, 'PM Overall Status')



####################################################

## IS THIS CODE STILL BEING USED?

# def get_project_data(project_id):
#     filtered_unique = df_unique[df_unique['Project ID'] == project_id]
#     if filtered_unique.empty:
#         return f"No data found for Project ID {project_id}"
    
#     proj_id = filtered_unique.iloc[0]['Project ID']
#     proj_type = filtered_unique.iloc[0]['Project Type']
#     total_budget = filtered_unique.iloc[0]['Budget']
#     use_budget = result_budget[result_budget['Project ID'] == proj_id]['Used Budget'].values[0]
#     total_datediff = result_schedule[result_schedule['Project ID'] == proj_id]['DateDiff'].values[0]
#     avg_budget = round((use_budget / total_budget), 2)
    
#     # Placeholder for overall_bin_dict data (replace with actual data)
#     overall_bin_dict = {
#         "Project 01": ["Bin1", "Bin2", "Bin3", "Bin4", "Bin5", "Bin6"],
#         "Project 02": ["BinA", "BinB", "BinC", "BinD", "BinE", "BinF"],
#         # Add other projects as needed
#     }
    
#     z = overall_bin_dict.get(proj_id, ["N/A"] * 6)
#     combined_data = [avg_budget, total_datediff] + z

#     return combined_data


###################################
@app.route('/')
def index():
    projectType()
    funnelChart()
    financial_activity()
    gauge_chart()

    file_path="Mockup_Dashboard_cleandatav2.xlsx"
    df_sheet2= pd.read_excel(file_path, sheet_name='Project Data')
    completed = round(df_sheet2.iloc[1, 14],0)
    to_do = round(df_sheet2.iloc[2,14],0)
    
    df_feedback=pd.read_excel(file_path, sheet_name='Feedback')
    feedbacka = round(df_feedback.iloc[1, 10],2)
    feedbackb = round(df_feedback.iloc[2, 10],2)
    feedbackc = round(df_feedback.iloc[3, 10],2)


    df_sheet= pd.read_excel("Mockup_Dashboard_cleandatav2.xlsx", sheet_name='Financials Data')
    accrual_amount = round(df_sheet['Accrual to Date'].sum(),2)
    cashout = round(df_sheet['Cash Out'].sum(),2) # usebudget
    totalbudget = round(df_sheet['LT_Budget'].sum(), 2)
    
    df_hours_sheet= pd.read_excel("Mockup_Dashboard_cleandatav2.xlsx", sheet_name='Project Data')
    hoursused = df_hours_sheet['Hours Used'].sum()
    estimatedHours = df_hours_sheet['Hours Estimated'].sum()
    
    return render_template('index.html',totalbudget=totalbudget,cashout=cashout,accrual=accrual_amount,completed=completed,to_do=to_do,hoursused=hoursused,estimatedHours=estimatedHours,a=feedbacka,b=feedbackb,c=feedbackc)

@app.route('/financial')
def financial():
    financial_activity()
    file_path="Mockup_Dashboard_cleandatav2.xlsx"
    df_sheet2= pd.read_excel(file_path, sheet_name='Project Data')
    completed = round(df_sheet2.iloc[1, 14],0)
    to_do = round(df_sheet2.iloc[2,14],0)


    df_sheet= pd.read_excel("Mockup_Dashboard_cleandatav2.xlsx", sheet_name='Financials Data')
    accrual_amount = round(df_sheet['Accrual to Date'].sum(),2)
    cashout = round(df_sheet['Cash Out'].sum(),2) # usebudget
    totalbudget = round(df_sheet['LT_Budget'].sum(), 2)

    return render_template('finance.html', totalbudget=totalbudget, cashout=cashout, accrual=accrual_amount, completed=completed, to_do=to_do)

@app.route('/project')
def project_activity():
    df_sheet= pd.read_excel("Mockup_Dashboard_cleandatav2.xlsx", sheet_name='Project Data')

    projectid=df_sheet['Project ID'].unique()
    unique_list = []
    for x in projectid:
        if x not in unique_list:
            unique_list.append(x)
    print(unique_list)
    return render_template('project.html',ids=unique_list)

#priject activity page
@app.route('/process')
def process():
    mdata = request.args.get('data')

    project_dict = main_all_data()
    milestonedata = milstone_gauntchart()
    if mdata in project_dict:
        value = project_dict[mdata]
        value2 = milestonedata[mdata]
    else:
        print("Key not found")

    return jsonify({'result': value,'result2':value2})
##################################################
def determine_final_status(status_list):
    if 'Done' in status_list:
        return 'Done'
    elif 'Started' in status_list:
        return 'Started'
    else:
        return 'Not Started'

def milstone_gauntchart():
    file_path = file
    # Read the data from the 'Financials Data' sheet
    df_sheet_name = pd.read_excel(file_path, sheet_name='Financials Data')   
    # Initialize an empty dictionary to store milestones
    milestones = {}   
    # Iterate over each row in the dataframe
    for index, row in df_sheet_name.iterrows():
        project_id = row['Project ID']       
        # Ensure the project_id exists in the dictionary
        if project_id not in milestones:
            milestones[project_id] = {}        
        # Update the dictionary with milestone data, excluding NaN and NaT values
        milestones[project_id].update({
            'm1': row['Milestone1'] if pd.notna(row['Milestone1']) else None,
            'm2': row['Milestone2'] if pd.notna(row['Milestone2']) else None,
            'm3': row['Milestone3'] if pd.notna(row['Milestone3']) else None,
            'm4': row['Milestone4'] if pd.notna(row['Milestone4']) else None,
            'm5': row['Milestone5'] if pd.notna(row['Milestone5']) else None,
            'm6': row['Milestone6'] if pd.notna(row['Milestone6']) else None,
            'm7': row['Milestone7'] if pd.notna(row['Milestone7']) else None,
            'm8': row['Milestone8'] if pd.notna(row['Milestone8']) else None,
            'm9': row['Milestone9'] if pd.notna(row['Milestone9']) else None,
            'm10': row['Milestone10'] if pd.notna(row['Milestone10']) else None
        })        
        # Remove keys with None values
        milestones[project_id] = {k: v for k, v in milestones[project_id].items() if v is not None}
    
    return milestones

################################
def main_all_data():
    file_path = file

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
###############################
@app.route('/portfolio')
def portfolio():
    project_dict = main_all_data()
    def is_nan(value):
        return value != value
# Remove entries with nan keys
    cleaned_data = {k: v for k, v in project_dict.items() if not (is_nan(k) or any(is_nan(vv) for vv in v.values()))}

    #print(project_dict)
    return render_template('portfolio.html',data=cleaned_data)


#####################################################
@app.route('/testdata', methods=['GET'])
def data():
    # Example JSON data
    # data = {
    #     'name': 'John Doe',
    #     'age': 30,
    #     'city': 'New York'
    # }
    df_sheet_name5 = pd.read_excel(file, sheet_name='Financials Data')
    acc_projetid = df_sheet_name5['Project ID']
    acc_cashout = round(df_sheet_name5['Cash Out'],2)
    acc_accrual = round(df_sheet_name5['Accrual to Date'],2)
    totlbudget = df_sheet_name5['LT_Budget']
    data = sorted(zip(acc_projetid, acc_accrual, acc_cashout,totlbudget))

# Unpack sorted data
    projects, accrual, cashout,total_budget = zip(*data)

    res = {"a": projects,"b":accrual,"c":cashout,"d":total_budget}
    return jsonify(res)



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
