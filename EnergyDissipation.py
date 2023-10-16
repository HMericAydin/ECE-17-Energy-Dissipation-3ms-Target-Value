#!/usr/bin/env python
# coding: utf-8

# In[4]:


from tkinter.filedialog import askopenfilename
import matplotlib.pyplot as plt
from matplotlib.ticker import (AutoMinorLocator, MultipleLocator)
import pandas as pd
import numpy as np
import os
from datetime import *
from matplotlib.backends.backend_pdf import PdfPages

## Browse Your File As .xlsx Format ##

filename = askopenfilename()

###
    # The example file has a sample number at the end of the description name...
    # ...In order to implement only sample "T" code to the graph and year...
    # ...That code is used
###
title_name = os.path.basename(filename).split('/')[-1] # Name split
title_name_w_out_dir = title_name.split('.')[-2] # Name split
T_code = title_name_w_out_dir.split(' ')[-1] # T-Code
Curr_date = str(date.today())
year = Curr_date.split('-')[0] # To implement only the year to the graph

df = pd.read_excel(filename, sheet_name='Data1').to_numpy() ## Type = numpy.ndarray ##

## Input Values
velocity = 24.1 #km/h

###
    # The values selected according to demands in the excel file. You can change the desired...
    # ...input columns by checking the example file
###

Time_col = df[1:(len(df)),0] #Time Values Column
AccAve_CFC_col = df[1:(len(df)),4] #Average Acc Column
Acc1_CFC_col = df[1:(len(df)),5] #The First Acc Column
Acc2_CFC_col = df[1:(len(df)),6] #The Second Acc Column

################################################################################################################################

### Algorithm
# ThirtyLoopsEvaluation Function: The actual aim is identifying whether the test data...
# ... has a significant value during 3ms.
def ThirtyLoopsEvaluation(Loops,time):
    if(len(Time_col)-(time+30)>=31):
        for i in range(time,time+30):
            cc = Loops[time]
            dd = Loops[i+1]
            if(i != time+29):
                if(cc<=dd):
                    return False
                    break
                else:
                    continue
            else:
                if(cc<=dd):
                    return False
                    break
                else:
                    return True
    else:
        return False

#AverageHighestValue Function: Determination the minimum acceleration value during 3ms for each gap.    
def AverageHighestValue(CFC):
    Max = float(0)
    max_i = 0
    for i in range(0,len(Time_col)-1):
        Temp_max = float(0)
        if(ThirtyLoopsEvaluation(CFC,i)):
            Temp_max = CFC[i]
            if(Max>=Temp_max):
                Max = Temp_max
                max_i = i
    return Max, max_i+3

# Determination of the first hitting point
def FirstInitialPoint(CFC):
    Min = float(0)
    min_i = 0
    for i in range(0,len(Time_col)-1):
        if (CFC[i]<=-7):
            Min = CFC[i]
            min_i = i
            break
    return Min, min_i+3

# Reducing the X-Limit range of the graph
def ReduceArray(CFC):
    reduced_Time_col = []
    for i in range(0,len(Time_col)):
        reduced_Time_col.append(Time_col[i]-Time_col[FirstInitialPoint(CFC)[1]]+0.002)
    return reduced_Time_col

################################################################################################################################
# Evaluation of Desired Values For Average Acceleration Values
def EnergyDissipationAverage(Values):
    # General Values #
    fig, ax = plt.subplots()
    ms_Reduce_Time_col = ReduceArray(AccAve_CFC_col)
    
    for j in range(0,len(ms_Reduce_Time_col)):
        ms_Reduce_Time_col[j]=ms_Reduce_Time_col[j]*1000
        
    Max_Time_Start = AverageHighestValue(Values)[1]
    Max_Val = AverageHighestValue(Values)[0]
    
    #########################  3ms - Start/End Lines  ##############################
    plt.axvline(x = ms_Reduce_Time_col[Max_Time_Start-2], color = 'b', label = "The 3ms Interval", linestyle='--', linewidth=0.5)
    plt.axvline(x = ms_Reduce_Time_col[Max_Time_Start+28], color = 'b', linestyle='--', linewidth=0.5)
    ################################################################################
    
    ############################ Graphs and Titles #################################
    plt.plot(ms_Reduce_Time_col,Values, color = 'b', label = "Acc Average", linewidth=0.5)
    plt.xlabel("Time - ms")
    plt.ylabel("g - m/s$^2$")
    plt.title("Acc. Average - "+"Sample Code: ("+T_code+") -"+" Year: ("+year+") ")
    plt.rc('legend', fontsize=5)
    plt.legend()
    ################################################################################
    
    ######################### Dashed Lines and Grid Gaps ###########################
    ax.xaxis.set_major_locator(MultipleLocator(10))
    ax.yaxis.set_major_locator(MultipleLocator(10))
    
    ax.xaxis.set_minor_locator(AutoMinorLocator(10)) # Change minor ticks to show every 5. (10/2 = 5)
    ax.yaxis.set_minor_locator(AutoMinorLocator(2))
    
    ax.grid(which='major', color='#CCCCCC', linestyle='--')
    ax.grid(which='minor', color='#CCCCCC', linestyle=':')
    ################################################################################
    
    ##################### Limitation and Text Modification #########################
    plt.xlim([-5,95]) # X-Limit
    plt.ylim([-100, 40]) # Y-Limit

    plt.figtext(0.15, -0.035, "Accelerations - Average:    Max:"+str("{:.1f}".format(max(Values)))+"g      Min:"+str("{:.1f}".format(min(Values)))+"g"+"      v="+str(velocity)+" km/h"+ "      m=6.8 kg", ha="left", fontsize=8, color= 'b')
    plt.figtext(0.15, -0.075, "3ms-Values - Average:      "+str("{:.1f}".format(Max_Val))+"g       "+str("{:.1f}".format(ms_Reduce_Time_col[Max_Time_Start-3]))+"ms - "+str("{:.1f}".format(ms_Reduce_Time_col[Max_Time_Start+27]))+"ms", ha="left", fontsize=8, color= 'b')
    ################################################################################
    
    ############################ Save The Data As .svg Format ######################
    image_format = 'svg' # e.g .png, .svg, etc.
    image_name = title_name_w_out_dir+'_Acc_Average.svg'

    fig.savefig(image_name, format=image_format, dpi=1200)
    ################################################################################
    
    return plt.show()

# Evaluation of Desired Values For Acceleration 1 & Acceleration 2 Values
def EnergyDissipationAcc1_Acc2(Acc1,Acc2):
    # General Values #
    fig, ax = plt.subplots()
    ms_Reduce_Time_col = ReduceArray(AccAve_CFC_col)
    
    for j in range(0,len(ms_Reduce_Time_col)):
        ms_Reduce_Time_col[j]=ms_Reduce_Time_col[j]*1000
        
    Max_Time_Start1 = AverageHighestValue(Acc1)[1]
    Max_Time_Start2 = AverageHighestValue(Acc2)[1]
    Max_Val1 = AverageHighestValue(Acc1)[0]
    Max_Val2 = AverageHighestValue(Acc2)[0]
    
    
    #########################  3ms - Start/End Lines  ##############################
    plt.axvline(x = ms_Reduce_Time_col[Max_Time_Start1-2], color = 'g', label = "The Acc1 3ms Interval", linestyle='--', linewidth=0.8)
    plt.axvline(x = ms_Reduce_Time_col[Max_Time_Start1+28], color = 'g', linestyle='--', linewidth=0.5)
    plt.axvline(x = ms_Reduce_Time_col[Max_Time_Start2-2], color = 'r', label = "The Acc2 3ms Interval", linestyle='--', linewidth=0.8)
    plt.axvline(x = ms_Reduce_Time_col[Max_Time_Start2+28], color = 'r', linestyle='--', linewidth=0.5)
    ################################################################################
    
    ############################ Graphs and Titles #################################
    plt.plot(ms_Reduce_Time_col,Acc1, color = 'g', label = "Acc1 Values", linewidth=0.5)
    plt.plot(ms_Reduce_Time_col,Acc2, color = 'r', label = "Acc2 Values", linewidth=0.5)
    plt.xlabel("Time - ms")
    plt.ylabel("g - m/s$^2$")
    plt.title("Acc1/Acc2 - "+"Sample Code: ("+T_code+") -"+" Year: ("+year+") ")
    plt.rc('legend', fontsize=5)
    plt.legend()
    ################################################################################
    
    ######################### Dashed Lines and Grid Gaps ###########################
    ax.xaxis.set_major_locator(MultipleLocator(10))
    ax.yaxis.set_major_locator(MultipleLocator(10))
    
    ax.xaxis.set_minor_locator(AutoMinorLocator(10)) # Change minor ticks to show every 5. (10/2 = 5)
    ax.yaxis.set_minor_locator(AutoMinorLocator(2))
    
    ax.grid(which='major', color='#CCCCCC', linestyle='--')
    ax.grid(which='minor', color='#CCCCCC', linestyle=':')
    ################################################################################
    
    ##################### Limitation and Text Modification #########################
    plt.xlim([-5, 95]) # X-Limit
    plt.ylim([-100, 40]) # Y-Limit
    
    ## Acc 1 Info
    plt.figtext(0.15, -0.035, "Accelerations - Acc1:    Max:"+str("{:.1f}".format(max(Acc1)))+"g      Min:"+str("{:.1f}".format(min(Acc1)))+"g", ha="left", fontsize=8, color= 'g')
    plt.figtext(0.15, -0.075, "3ms-Values - Acc1:      "+str("{:.1f}".format(Max_Val1))+"g       "+str("{:.1f}".format(ms_Reduce_Time_col[Max_Time_Start1-3]))+"ms - "+str("{:.1f}".format(ms_Reduce_Time_col[Max_Time_Start1+27]))+"ms", ha="left", fontsize=8, color= 'g')
    
    ## Acc 2 Info
    plt.figtext(0.15, -0.14, "Accelerations - Acc2:    Max:"+str("{:.1f}".format(max(Acc2)))+"g      Min:"+str("{:.1f}".format(min(Acc2)))+"g", ha="left", fontsize=8, color= 'r')
    plt.figtext(0.15, -0.18, "3ms-Values - Acc2:      "+str("{:.1f}".format(Max_Val2))+"g       "+str("{:.1f}".format(ms_Reduce_Time_col[Max_Time_Start2-3]))+"ms - "+str("{:.1f}".format(ms_Reduce_Time_col[Max_Time_Start2+27]))+"ms", ha="left", fontsize=8, color= 'r')
    ################################################################################
    
    ############################ Save The Data As .svg Format ######################
    image_format = 'svg' # e.g .png, .svg, etc.
    image_name1 = title_name_w_out_dir+'_Acc1_Acc2.svg'

    fig.savefig(image_name1, format=image_format, dpi=1200)
    ################################################################################
    
    return plt.show()

###
    #Additional Notes:
    # 1) Run the code
    # 2) Choose the Excel file. (The file name should be ended with SPACE + T-Code of the Sample).
    # 3) Go "Home Page" tab in Chrome. Select your graphs with identified name.
        # 3.1) Before download the file change the viewBox value in the 4th line in the page...
        #... to see the graph description in the below. For Average value increase the viewBox value 40...
        #... for the Acc1/Acc2 increase the viewBox value 80 which represent the Height value in the viewBox.
        # 3.2) File-->Save, File-->Download
###

EnergyDissipationAverage(AccAve_CFC_col)
EnergyDissipationAcc1_Acc2(Acc1_CFC_col,Acc2_CFC_col)

