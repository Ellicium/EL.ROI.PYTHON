import datetime
import json
import logging,pycountry
import math
import os
import re
from time import time

import pandas as pd
from dotenv import load_dotenv
from fastapi.logger import logger
from ..config.logger_config import get_logger
from openpyxl import load_workbook
import numpy as np
import math

# # logger = get_logger()
load_dotenv()




# ----------------------------------------/


def modify_excel_fields(excel_file):
    try:
        filepath=excel_file
        # filepath=r"C:\Users\Krushna_Kadam\Downloads\RPA-ROI-Calculator v1.0.xlsx"
        # Load the Excel data into a pandas DataFrame
        roi = pd.read_excel(filepath, sheet_name="ROI", header=[2,3 ])
        roi.columns = pd.MultiIndex.from_tuples([(col[0].strip(), col[1].strip()) for col in roi.columns])

        columns_to_change = ['License Cost (Per Day)','Development & Support']  # Columns to rename using sub-column names

        # Create new column names
        new_columns = [
            col[1] if col[0] in columns_to_change else col[0] 
            for col in roi.columns
        ]

        # Assign new column names
        roi.columns = new_columns

        roi.columns = roi.columns.str.strip()


        input_columns=['#Priority','UseCase','Already in Production','AverageTransactions(Annual)','OCR/NON-OCR','Frequency','Complexity','Manual MHT / AHT in Min /Transaction']

        for i in roi.columns:
            if i not in input_columns:
                roi[i]=None

        #read assumptions sheet
        assumptions = pd.read_excel(filepath, sheet_name="Assumptions")
        assumptions =assumptions.set_index(assumptions.columns[0]).T.reset_index(drop=True)
        assumptions.columns = assumptions.columns.str.strip()

        def calculate_BOT_HT_AHT(Complexity_val):
            percentage=0
            # print(Complexity_val)
            if Complexity_val.lower().strip()=='medium':
                percentage=assumptions['BOT Accuracy - Medium'][0]
            elif Complexity_val.lower().strip()=='complex':
                percentage=assumptions['BOT Accuracy - Complex'][0]
            else:
                percentage=assumptions['BOT Accuracy - Simple'][0]
            print('percentage',percentage)
            return percentage
            
        def check_prod(val):
            if str(val).lower()=='y':
                return 0
            return 1
        
        def check_ocr(val):
            if str(val).lower().strip()=='ocr':
                return 1
            return 0

        roi=roi[:-1]
        roi['Manual FTE Required/Day']=roi['AverageTransactions(Annual)']*roi['Manual MHT / AHT in Min /Transaction']/assumptions['Mins Per Hour'][0]/assumptions['Working Days in a Year'][0]/assumptions['Working Hours per Day'][0]
        roi['FTE Cost (Per Day)']=roi['Manual FTE Required/Day']*(assumptions['Monthly FTE Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Seat Cost (Per Day)']=roi['Manual FTE Required/Day']*(assumptions['Monthly Seat Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Employee Cost (Others - Paid by )* Per Day']=roi['Manual FTE Required/Day']*(assumptions['Monthly FTEOther Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Other OPEX Cost*Per Day']=roi['Manual FTE Required/Day']*(assumptions['Monthly Other OPEX'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Total Cost (Per Day)']=roi['Manual FTE Required/Day']*(((assumptions['Monthly FTE Cost'][0]+assumptions['Monthly Seat Cost'][0]+assumptions['Monthly FTEOther Cost'][0]+assumptions['Monthly Other OPEX'][0]))/assumptions['Monthly Paid Days for FTE'][0])
        roi['calculate_percentage']=roi['Complexity'].apply(calculate_BOT_HT_AHT)
        roi['FTE Saved/Day']=roi['Manual FTE Required/Day']*roi['calculate_percentage']

        roi['BOT MHT / AHT (Min)']=(roi['Manual MHT / AHT in Min /Transaction']-(roi['Manual MHT / AHT in Min /Transaction']*roi['calculate_percentage']))+(roi['Manual MHT / AHT in Min /Transaction']-(roi['Manual MHT / AHT in Min /Transaction']*roi['calculate_percentage']))*.02



        roi['BOT FTE Required / Day']=roi['Manual FTE Required/Day']-roi['FTE Saved/Day']
        roi['FTE Cost After BOT Implementation(Per Day)']=roi['BOT FTE Required / Day']*(assumptions['Monthly FTE Cost'][0]+assumptions['Monthly Seat Cost'][0]+assumptions['Monthly FTEOther Cost'][0]+assumptions['Monthly Other OPEX'][0])/assumptions['BOT Number of Days Per Month'][0]
        roi['Infra Cost (Per Day)']=(((assumptions['Monthly VM Cost'][0]*assumptions["Number of VM's"][0])/assumptions['Number of Processes'][0])/assumptions['BOT Number of Days Per Month'][0])
        roi['Runner']=(((roi['AverageTransactions(Annual)']/assumptions['BOT Days in a Year'][0])*roi['BOT MHT / AHT (Min)'])/assumptions['Mins Per Hour'][0])*(assumptions['Monthly Runner License Cost'][0]/(assumptions['BOT Number of Days Per Month'][0]*assumptions['BOT Number of Hours in a Day'][0]))

        roi['Creater']=(assumptions['Monthly Creater License Cost'][0]/assumptions['BOT Number of Days Per Month'][0])/assumptions['Creater License Average Out Time'][0]
        roi['OCR']=(assumptions['Document Automation Cost'][0]/assumptions['Number Of Documents per OCR License'][0])*(roi['AverageTransactions(Annual)']/assumptions['BOT Days in a Year'][0])

        roi['checkProduction']=roi['Already in Production'].apply(check_prod)

        roi['check_ocr']=roi['OCR/NON-OCR'].apply(check_ocr)
        roi['OCR']=roi['OCR']*roi['check_ocr']
        roi['Support']=(math.ceil(len(roi['UseCase'])/8)*100000/len(roi['UseCase']))/assumptions['BOT Number of Days Per Month'][0]
        
        # roi['OCR']=roi['OCR']*roi['checkProduction']


        row_cols_to_sum=['FTE Cost After BOT Implementation(Per Day)',
            'Infra Cost (Per Day)', 'Runner', 'OCR', 'Support']

        roi['Total Cost (Excluding Development & Creater Cost)'] = roi[row_cols_to_sum].sum(axis=1)
        roi['ROI in Rupees (Per Day)']=roi['Total Cost (Per Day)']-roi['Total Cost (Excluding Development & Creater Cost)']
        roi['ROI In Percent']=(roi['Total Cost (Per Day)']-roi['Total Cost (Excluding Development & Creater Cost)'])/roi['Total Cost (Per Day)']
        
        # (math.ceil(len(roi['UseCase'])/8*100000)/assumptions['Number of Processes'][0])/assumptions['BOT Number of Days Per Month'][0]

        roi.drop(columns=['calculate_percentage','checkProduction','check_ocr'], inplace=True)

        # Load the original Excel file

        print('load_workbook')
        wb = load_workbook(filepath)
        print('load_workbook donme')
        ws = wb["ROI"]

        # Load the transformed data
        updated_df = roi
        max_row=0

        updated_df = updated_df.replace({np.nan: None})


        for row_idx, row in enumerate(ws.iter_rows(), start=1):
            if row_idx <= 4:  # Skip first 5 rows
                continue
            print(row_idx, row)
            max_row = row_idx

            for cell in row[8:]:  # Skip the first 8 columns (index 0 to 7)
            
                cell.value = None  # Clear the cell value but keep formatting

        # Update only values, keeping formatting intact
        for row_idx, row in enumerate(updated_df.itertuples(index=False), start=5):  
            for col_idx, value in enumerate(row, start=1):
                # print( row_idx, row)
                # print(col_idx, value)
                # break

                ws.cell(row=row_idx, column=col_idx, value=value)


        ws['I'+str(max_row)].value = roi['Manual FTE Required/Day'].sum()
        ws['N'+str(max_row)].value = roi['Total Cost (Per Day)'].sum()
        ws['Q'+str(max_row)].value = roi['BOT FTE Required / Day'].sum()
        ws['R'+str(max_row)].value = roi['FTE Cost After BOT Implementation(Per Day)'].sum()
        ws['Y'+str(max_row)].value = roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ws['Z'+str(max_row)].value = roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ws['AA'+str(max_row)].value = (roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum())/roi['Total Cost (Per Day)'].sum()


        # # do_not write in excel
        # FTE_Saved_Day = roi['FTE Saved/Day'].sum()/len(roi['FTE Saved/Day'])
        # BOT_MHT_AHT_Min = roi['BOT MHT / AHT (Min)'].sum()/len(roi['BOT MHT / AHT (Min)'])
        # # BOT_FTE = roi['BOT FTE'].sum()/len(roi['BOT FTE'])
        # # FTE_Cost= roi['FTE Cost'].sum()/len(roi['FTE Cost'])

        # # ---
        # BOT_FTE_Required_Day=roi['BOT FTE Required / Day'].sum()/len(roi['BOT FTE Required / Day'])
        # FTE_Cost_After_BOT_Implementation=roi['FTE Cost After BOT Implementation(Per Day)'].sum()/len(roi['FTE Cost After BOT Implementation(Per Day)'])
        # Infra_Cost_Per_Day=roi['Infra Cost (Per Day)'].sum()/len(roi['Infra Cost (Per Day)'])
        # License_Cost_Per_Day=roi['Runner'].sum()/len(roi['Runner'])+roi['Creater'].sum()/len(roi['Creater'])+roi['OCR'].sum()/len(roi['OCR'])
        # Development_Support=roi['Development'].sum()/len(roi['Development'])+roi['Support'].sum()/len(roi['Support'])
        # Total_Cost_Excluding_Development_Creater_Cost=(roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum())/roi['Total Cost (Per Day)'].sum()

        # # Save the updated Excel file
        # wb.save("data_modified2.xlsx")
        # print('writtennnnnn')
        # return "data_modified2.xlsx"


        # do_not write in excel
        FTE_Saved_Day = roi['FTE Saved/Day'].sum()/len(roi['FTE Saved/Day'])
        BOT_MHT_AHT_Min = roi['BOT MHT / AHT (Min)'].sum()/len(roi['BOT MHT / AHT (Min)'])
        # BOT_FTE = roi['BOT FTE'].sum()/len(roi['BOT FTE'])
        # FTE_Cost= roi['FTE Cost'].sum()/len(roi['FTE Cost'])

        # ---
        BOT_FTE_Required_Day=roi['BOT FTE Required / Day'].sum()/len(roi['BOT FTE Required / Day'])
        FTE_Cost_After_BOT_Implementation=roi['FTE Cost After BOT Implementation(Per Day)'].sum()
        # FTE_Cost_After_BOT_Implementation=roi['FTE Cost After BOT Implementation(Per Day)'].sum()/len(roi['FTE Cost After BOT Implementation(Per Day)'])
        Infra_Cost_Per_Day=roi['Infra Cost (Per Day)'].sum()/len(roi['Infra Cost (Per Day)'])
        License_Cost_Per_Day=roi['Runner'].sum()/len(roi['Runner'])+roi['Creater'].sum()/len(roi['Creater'])+roi['OCR'].sum()/len(roi['OCR'])
        Development_Support=roi['Development'].sum()/len(roi['Development'])+roi['Support'].sum()/len(roi['Support'])
        Total_Cost_Excluding_Development_Creater_Cost=roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ROI_in_Rupees=roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ROI_In_Percent=(roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum())/roi['Total Cost (Per Day)'].sum()


        
        percentage=calculate_BOT_HT_AHT(roi['Complexity'][0])

        return_dict={

            'FTE_Saved_Day':FTE_Saved_Day,
            'BOT_MHT_AHT_Min':BOT_MHT_AHT_Min,
            # 'BOT_FTE':None,
            # 'FTE_Cost':None,
            'BOT_FTE_Required_Day':BOT_FTE_Required_Day,
            'FTE_Cost_After_BOT_Implementation':FTE_Cost_After_BOT_Implementation,
            'Infra_Cost_Per_Day':Infra_Cost_Per_Day,
            'License_Cost_Per_Day':License_Cost_Per_Day,
            'Development_Support':Development_Support,
            'Total_Cost_Excluding_Development_Creater_Cost':Total_Cost_Excluding_Development_Creater_Cost,
            'ROI_in_Rupees':ROI_in_Rupees,
            'ROI_In_Percent':ROI_In_Percent*100,
            'BOT_FTE_Calculated':roi['Manual FTE Required/Day'].mean()*percentage,
            'Manual_Cost':roi['Total Cost (Per Day)'].sum()
        }

        # Save the updated Excel file
        wb.save("data_modified2.xlsx")
        print('writtennnnnn')
        return ["data_modified2.xlsx",return_dict]

    except:
        return 0

    # except:
    #     return 0
    

def update_user_values(val):
    try:
        dictt={}
        dictt['usecase']= val.usecase
        dictt['frequency']= val.frequency
        dictt['manual_mht_aht']= val.manual_mht_aht
        dictt['avg_transaction_annual']= val.avg_transaction_annual
        dictt['number_of_vms']= val.number_of_vms
        dictt['is_ocr']= val.is_ocr
        dictt['already_in_prod']= val.already_in_prod
        dictt['complexivity']= val.complexivity
        dictt['monthly_vm_cost']= val.monthly_vm_cost
        dictt['monthly_fte_cost']= val.monthly_fte_cost
        dictt['monthly_seat_cost']= val.monthly_seat_cost
        dictt['monthly_fte_other_cost']= val.monthly_fte_other_cost
        dictt['monthly_runner_licence_cost']= val.monthly_runner_licence_cost
        dictt['document_automation_cost']= val.document_automation_cost
        dictt['monthly_creater_licence_cost']= val.monthly_creater_licence_cost
        dictt['support_cost_per_resource']= val.support_cost_per_resource
        dictt['monthly_developer_cost_per_resource']= val.monthly_developer_cost_per_resource
    
        filepath="RPA-ROI-Calculator v4.0 - Copy.xlsx"#r"C:\Users\Krushna_Kadam\Downloads\RPA-ROI-Calculator v4.0 - Copy.xlsx"
        # Load the Excel data into a pandas DataFrame
        #read assumptions sheet

        print('load_workbook')
        wb = load_workbook(filepath)
        print('load_workbook donme')
        ws = wb["ROI"]
        wd = wb["Assumptions"]

        input_dict=dictt

        ws['B' + str(5)].value = input_dict['usecase']
        ws['C' + str(5)].value = input_dict['already_in_prod']
        ws['D' + str(5)].value = input_dict['avg_transaction_annual']
        ws['E' + str(5)].value = input_dict['is_ocr']
        ws['F' + str(5)].value = input_dict['frequency']
        ws['G' + str(5)].value = input_dict['complexivity']
        ws['H' + str(5)].value = input_dict['manual_mht_aht']


        number_of_vms=input_dict['number_of_vms']
        monthly_vm_cost=input_dict['monthly_vm_cost']
        monthly_fte_cost=input_dict['monthly_fte_cost']
        monthly_seat_cost=input_dict['monthly_seat_cost']
        monthly_fte_other_cost=input_dict['monthly_fte_other_cost']
        monthly_runner_licence_cost=input_dict['monthly_runner_licence_cost']
        document_automation_cost=input_dict['document_automation_cost']
        monthly_creater_licence_cost=input_dict['monthly_creater_licence_cost']
        support_cost_per_resource=input_dict['support_cost_per_resource']
        monthly_developer_cost_per_resource=input_dict['monthly_developer_cost_per_resource']


        # over write user input values
        if number_of_vms is not None:
            wd['B16']=number_of_vms

        if monthly_vm_cost is not None:
            wd['B9']=monthly_vm_cost

        if monthly_fte_cost  is not None:
            wd['B11']=monthly_fte_cost

        if monthly_seat_cost  is not None:
            wd['B13']=monthly_seat_cost

        if monthly_fte_other_cost  is not None:
            wd['B14']=monthly_fte_other_cost
        if monthly_runner_licence_cost is not None:
            wd['B21']=monthly_runner_licence_cost
        if document_automation_cost is not None:
            wd['B22']=document_automation_cost
        if monthly_creater_licence_cost is not None:
            wd['B23']=monthly_creater_licence_cost
        if support_cost_per_resource is not None:
            wd['B27']=support_cost_per_resource
        if monthly_developer_cost_per_resource is not None:
            wd['B26']=monthly_developer_cost_per_resource
        
        

        wb.save('data_modified4.xlsx')



        filepath='data_modified4.xlsx'
        # filepath=r"C:\Users\Krushna_Kadam\Downloads\RPA-ROI-Calculator v1.0.xlsx"
        # Load the Excel data into a pandas DataFrame
        roi = pd.read_excel(filepath, sheet_name="ROI", header=[2,3 ])
        roi.columns = pd.MultiIndex.from_tuples([(col[0].strip(), col[1].strip()) for col in roi.columns])

        columns_to_change = ['License Cost (Per Day)','Development & Support']  # Columns to rename using sub-column names

        # Create new column names
        new_columns = [
            col[1] if col[0] in columns_to_change else col[0] 
            for col in roi.columns
        ]

        # Assign new column names
        roi.columns = new_columns

        roi.columns = roi.columns.str.strip()


        input_columns=['#Priority','UseCase','Already in Production','AverageTransactions(Annual)','OCR/NON-OCR','Frequency','Complexity','Manual MHT / AHT in Min /Transaction']

        for i in roi.columns:
            if i not in input_columns:
                roi[i]=None

        #read assumptions sheet
        assumptions = pd.read_excel(filepath, sheet_name="Assumptions")
        assumptions =assumptions.set_index(assumptions.columns[0]).T.reset_index(drop=True)
        assumptions.columns = assumptions.columns.str.strip()

        def calculate_BOT_HT_AHT(Complexity_val):
            percentage=0
            # print(Complexity_val)
            if Complexity_val.lower().strip()=='medium':
                percentage=assumptions['BOT Accuracy - Medium'][0]
            elif Complexity_val.lower().strip()=='complex':
                percentage=assumptions['BOT Accuracy - Complex'][0]
            else:
                percentage=assumptions['BOT Accuracy - Simple'][0]

            print('percentage',percentage)
            return percentage
            
        def check_prod(val):
            if str(val).lower()=='y':
                return 0
            return 1
        
        def check_ocr(val):
            if str(val).lower().strip()=='ocr':
                return 1
            return 0


        roi=roi[:-1]
        roi['Manual FTE Required/Day']=roi['AverageTransactions(Annual)']*roi['Manual MHT / AHT in Min /Transaction']/assumptions['Mins Per Hour'][0]/assumptions['Working Days in a Year'][0]/assumptions['Working Hours per Day'][0]
        roi['FTE Cost (Per Day)']=roi['Manual FTE Required/Day']*(assumptions['Monthly FTE Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Seat Cost (Per Day)']=roi['Manual FTE Required/Day']*(assumptions['Monthly Seat Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Employee Cost (Others - Paid by )* Per Day']=roi['Manual FTE Required/Day']*(assumptions['Monthly FTEOther Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Other OPEX Cost*Per Day']=roi['Manual FTE Required/Day']*(assumptions['Monthly Other OPEX'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Total Cost (Per Day)']=roi['Manual FTE Required/Day']*(((assumptions['Monthly FTE Cost'][0]+assumptions['Monthly Seat Cost'][0]+assumptions['Monthly FTEOther Cost'][0]+assumptions['Monthly Other OPEX'][0]))/assumptions['Monthly Paid Days for FTE'][0])
        roi['calculate_percentage']=roi['Complexity'].apply(calculate_BOT_HT_AHT)
        roi['FTE Saved/Day']=roi['Manual FTE Required/Day']*roi['calculate_percentage']

        roi['BOT MHT / AHT (Min)']=(roi['Manual MHT / AHT in Min /Transaction']-(roi['Manual MHT / AHT in Min /Transaction']*roi['calculate_percentage']))+(roi['Manual MHT / AHT in Min /Transaction']-(roi['Manual MHT / AHT in Min /Transaction']*roi['calculate_percentage']))*.02


        roi['BOT FTE Required / Day']=roi['Manual FTE Required/Day']-roi['FTE Saved/Day']
        roi['FTE Cost After BOT Implementation(Per Day)']=roi['BOT FTE Required / Day']*(assumptions['Monthly FTE Cost'][0]+assumptions['Monthly Seat Cost'][0]+assumptions['Monthly FTEOther Cost'][0]+assumptions['Monthly Other OPEX'][0])/assumptions['BOT Number of Days Per Month'][0]
        roi['Infra Cost (Per Day)']=(((assumptions['Monthly VM Cost'][0]*assumptions["Number of VM's"][0])/assumptions['Number of Processes'][0])/assumptions['BOT Number of Days Per Month'][0])
        roi['Runner']=(((roi['AverageTransactions(Annual)']/assumptions['BOT Days in a Year'][0])*roi['BOT MHT / AHT (Min)'])/assumptions['Mins Per Hour'][0])*(assumptions['Monthly Runner License Cost'][0]/(assumptions['BOT Number of Days Per Month'][0]*assumptions['BOT Number of Hours in a Day'][0]))

        roi['Creater']=(assumptions['Monthly Creater License Cost'][0]/assumptions['BOT Number of Days Per Month'][0])/assumptions['Creater License Average Out Time'][0]
        roi['OCR']=(assumptions['Document Automation Cost'][0]/assumptions['Number Of Documents per OCR License'][0])*(roi['AverageTransactions(Annual)']/assumptions['BOT Days in a Year'][0])
        roi['Support']=(math.ceil(len(roi['UseCase'])/8)*100000/len(roi['UseCase']))/assumptions['BOT Number of Days Per Month'][0]
        # (math.ceil(23/8)*100000/23)/30

        roi['checkProduction']=roi['Already in Production'].apply(check_prod)

        roi['check_ocr']=roi['OCR/NON-OCR'].apply(check_ocr)

        roi['OCR']=roi['OCR']*roi['check_ocr']
        print('OCR COLUMN',roi['OCR'],roi['checkProduction'])

        row_cols_to_sum=['FTE Cost After BOT Implementation(Per Day)',
            'Infra Cost (Per Day)', 'Runner', 'OCR',
            'Support']
        
        # print('sum',roi['FTE Cost After BOT Implementation(Per Day)'][0],roi['Infra Cost (Per Day)'][0],roi['Runner'][0],roi['OCR'][0], roi['Support'][0])

        roi['Total Cost (Excluding Development & Creater Cost)'] =roi[row_cols_to_sum].sum(axis=1, skipna=True)

        roi['ROI in Rupees (Per Day)']=roi['Total Cost (Per Day)']-roi['Total Cost (Excluding Development & Creater Cost)']
        roi['ROI In Percent']=(roi['Total Cost (Per Day)']-roi['Total Cost (Excluding Development & Creater Cost)'])/roi['Total Cost (Per Day)']

        roi.drop(columns=['calculate_percentage','checkProduction','check_ocr'], inplace=True)

        # Load the original Excel file

        print('load_workbook')
        wb = load_workbook(filepath)
        print('load_workbook donme')
        ws = wb["ROI"]

        # Load the transformed data
        updated_df = roi
        max_row=0

        updated_df = updated_df.replace({np.nan: None})


        for row_idx, row in enumerate(ws.iter_rows(), start=1):
            if row_idx <= 4:  # Skip first 5 rows
                continue
            print(row_idx, row)
            max_row = row_idx

            for cell in row[8:]:  # Skip the first 8 columns (index 0 to 7)
            
                cell.value = None  # Clear the cell value but keep formatting

        # Update only values, keeping formatting intact
        for row_idx, row in enumerate(updated_df.itertuples(index=False), start=5):  
            for col_idx, value in enumerate(row, start=1):
                # print( row_idx, row)
                # print(col_idx, value)
                # break

                ws.cell(row=row_idx, column=col_idx, value=value)


        ws['I'+str(max_row)].value = roi['Manual FTE Required/Day'].sum()
        ws['N'+str(max_row)].value = roi['Total Cost (Per Day)'].sum()
        ws['Q'+str(max_row)].value = roi['BOT FTE Required / Day'].sum()
        ws['R'+str(max_row)].value = roi['FTE Cost After BOT Implementation(Per Day)'].sum()
        ws['Y'+str(max_row)].value = roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ws['Z'+str(max_row)].value = roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ws['AA'+str(max_row)].value = (roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum())/roi['Total Cost (Per Day)'].sum()


        # do_not write in excel
        FTE_Saved_Day = roi['FTE Saved/Day'].sum()/len(roi['FTE Saved/Day'])
        BOT_MHT_AHT_Min = roi['BOT MHT / AHT (Min)'].sum()/len(roi['BOT MHT / AHT (Min)'])
        # BOT_FTE = roi['BOT FTE'].sum()/len(roi['BOT FTE'])
        # FTE_Cost= roi['FTE Cost'].sum()/len(roi['FTE Cost'])

        # ---
        BOT_FTE_Required_Day=roi['BOT FTE Required / Day'].sum()/len(roi['BOT FTE Required / Day'])
        # FTE_Cost_After_BOT_Implementation=roi['FTE Cost After BOT Implementation(Per Day)'].sum()/len(roi['FTE Cost After BOT Implementation(Per Day)'])
        FTE_Cost_After_BOT_Implementation=roi['FTE Cost After BOT Implementation(Per Day)'].sum()
        Infra_Cost_Per_Day=roi['Infra Cost (Per Day)'].sum()/len(roi['Infra Cost (Per Day)'])
        License_Cost_Per_Day=roi['Runner'].sum()/len(roi['Runner'])+roi['Creater'].sum()/len(roi['Creater'])+roi['OCR'].sum()/len(roi['OCR'])
        Development_Support=roi['Development'].sum()/len(roi['Development'])+roi['Support'].sum()/len(roi['Support'])
        Total_Cost_Excluding_Development_Creater_Cost=roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ROI_in_Rupees=roi['ROI in Rupees (Per Day)'][0]
        ROI_In_Percent=roi['ROI In Percent'][0]


        #BOT FTE

        percentage=calculate_BOT_HT_AHT(roi['Complexity'][0])
        

        return_dict={

            'FTE_Saved_Day':FTE_Saved_Day,
            'BOT_MHT_AHT_Min':BOT_MHT_AHT_Min,
            # 'BOT_FTE':None,
            # 'FTE_Cost':None,
            'BOT_FTE_Required_Day':BOT_FTE_Required_Day,
            'FTE_Cost_After_BOT_Implementation':FTE_Cost_After_BOT_Implementation,
            'Infra_Cost_Per_Day':Infra_Cost_Per_Day,
            'License_Cost_Per_Day':License_Cost_Per_Day,
            'Development_Support':Development_Support,
            'Total_Cost_Excluding_Development_Creater_Cost':Total_Cost_Excluding_Development_Creater_Cost,
            'ROI_in_Rupees':ROI_in_Rupees,
            'ROI_In_Percent':ROI_In_Percent*100,
            'BOT_FTE_Calculated':roi['Manual FTE Required/Day'][0]*percentage,
            'Manual_Cost':roi['Total Cost (Per Day)'].sum()
        }

        # Save the updated Excel file
        wb.save("data_modified2.xlsx")
        print('writtennnnnn')
        return ["data_modified2.xlsx",return_dict]

        
    except:
        return 0




def modify_excel_fields_v2(filepath,monthly_vm_cost,monthly_fte_cost,monthly_seat_cost,monthly_fte_other_cost,monthly_runner_licence_cost,document_automation_cost,monthly_creater_licence_cost,support_cost_per_resource,monthly_developer_cost_per_resource):
    try:
        dictt={}
        dictt['monthly_vm_cost']= monthly_vm_cost
        dictt['monthly_fte_cost']= monthly_fte_cost
        dictt['monthly_seat_cost']= monthly_seat_cost
        dictt['monthly_fte_other_cost']= monthly_fte_other_cost
        dictt['monthly_runner_licence_cost']= monthly_runner_licence_cost
        dictt['document_automation_cost']= document_automation_cost
        dictt['monthly_creater_licence_cost']= monthly_creater_licence_cost
        dictt['support_cost_per_resource']= support_cost_per_resource
        dictt['monthly_developer_cost_per_resource']=monthly_developer_cost_per_resource

        # filepath="RPA-ROI-Calculator v4.0 - Copy.xlsx"#r"C:\Users\Krushna_Kadam\Downloads\RPA-ROI-Calculator v4.0 - Copy.xlsx"
        # Load the Excel data into a pandas DataFrame
        #read assumptions sheet

        print('load_workbook')
        wb = load_workbook(filepath, data_only=True)
        print('load_workbook donme')
        ws = wb["ROI"]
        wd = wb["Assumptions"]

        input_dict=dictt

        monthly_vm_cost=input_dict['monthly_vm_cost']
        monthly_fte_cost=input_dict['monthly_fte_cost']
        monthly_seat_cost=input_dict['monthly_seat_cost']
        monthly_fte_other_cost=input_dict['monthly_fte_other_cost']
        monthly_runner_licence_cost=input_dict['monthly_runner_licence_cost']
        document_automation_cost=input_dict['document_automation_cost']
        monthly_creater_licence_cost=input_dict['monthly_creater_licence_cost']
        support_cost_per_resource=input_dict['support_cost_per_resource']
        monthly_developer_cost_per_resource=input_dict['monthly_developer_cost_per_resource']

        # over write user input values
        if monthly_vm_cost is not None:
            wd['B9']=monthly_vm_cost

        if monthly_fte_cost  is not None:
            wd['B11']=monthly_fte_cost

        if monthly_seat_cost  is not None:
            wd['B13']=monthly_seat_cost

        if monthly_fte_other_cost  is not None:
            wd['B14']=monthly_fte_other_cost
        if monthly_runner_licence_cost is not None:
            wd['B21']=monthly_runner_licence_cost
        if document_automation_cost is not None:
            wd['B22']=document_automation_cost
        if monthly_creater_licence_cost is not None:
            wd['B23']=monthly_creater_licence_cost
        if support_cost_per_resource is not None:
            wd['B27']=support_cost_per_resource
        if monthly_developer_cost_per_resource is not None:
            wd['B26']=monthly_developer_cost_per_resource

        wb.save('data_modified_5.xlsx')



        filepath='data_modified_5.xlsx'
        
        
        # filepath=r"C:\Users\Krushna_Kadam\Downloads\RPA-ROI-Calculator v1.0.xlsx"
        # Load the Excel data into a pandas DataFrame
        roi = pd.read_excel(filepath, sheet_name="ROI", header=[2,3 ])
        roi.columns = pd.MultiIndex.from_tuples([(col[0].strip(), col[1].strip()) for col in roi.columns])

        columns_to_change = ['License Cost (Per Day)','Development & Support']  # Columns to rename using sub-column names

        # Create new column names
        new_columns = [
            col[1] if col[0] in columns_to_change else col[0] 
            for col in roi.columns
        ]

        # Assign new column names
        roi.columns = new_columns

        roi.columns = roi.columns.str.strip()


        input_columns=['#Priority','UseCase','Already in Production','AverageTransactions(Annual)','OCR/NON-OCR','Frequency','Complexity','Manual MHT / AHT in Min /Transaction']

        for i in roi.columns:
            if i not in input_columns:
                roi[i]=None

        #read assumptions sheet
        assumptions = pd.read_excel(filepath, sheet_name="Assumptions")
        assumptions =assumptions.set_index(assumptions.columns[0]).T.reset_index(drop=True)
        assumptions.columns = assumptions.columns.str.strip()

        def calculate_BOT_HT_AHT(Complexity_val):
            percentage=0
            # print(Complexity_val)
            if Complexity_val.lower().strip()=='medium':
                percentage=assumptions['BOT Accuracy - Medium'][0]
            elif Complexity_val.lower().strip()=='complex':
                percentage=assumptions['BOT Accuracy - Complex'][0]
            else:
                percentage=assumptions['BOT Accuracy - Simple'][0]
            print('percentage',percentage)
            return percentage
            
        def check_prod(val):
            if str(val).lower()=='y':
                return 0
            return 1
        
        def check_ocr(val):
            if str(val).lower().strip()=='ocr':
                return 1
            return 0

        roi=roi[:-1]
        roi['Manual FTE Required/Day']=roi['AverageTransactions(Annual)']*roi['Manual MHT / AHT in Min /Transaction']/assumptions['Mins Per Hour'][0]/assumptions['Working Days in a Year'][0]/assumptions['Working Hours per Day'][0]
        roi['FTE Cost (Per Day)']=roi['Manual FTE Required/Day']*(assumptions['Monthly FTE Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Seat Cost (Per Day)']=roi['Manual FTE Required/Day']*(assumptions['Monthly Seat Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Employee Cost (Others - Paid by )* Per Day']=roi['Manual FTE Required/Day']*(assumptions['Monthly FTEOther Cost'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Other OPEX Cost*Per Day']=roi['Manual FTE Required/Day']*(assumptions['Monthly Other OPEX'][0]/assumptions['Monthly Paid Days for FTE'][0])
        roi['Total Cost (Per Day)']=roi['Manual FTE Required/Day']*(((assumptions['Monthly FTE Cost'][0]+assumptions['Monthly Seat Cost'][0]+assumptions['Monthly FTEOther Cost'][0]+assumptions['Monthly Other OPEX'][0]))/assumptions['Monthly Paid Days for FTE'][0])
        roi['calculate_percentage']=roi['Complexity'].apply(calculate_BOT_HT_AHT)
        roi['FTE Saved/Day']=roi['Manual FTE Required/Day']*roi['calculate_percentage']

        roi['BOT MHT / AHT (Min)']=(roi['Manual MHT / AHT in Min /Transaction']-(roi['Manual MHT / AHT in Min /Transaction']*roi['calculate_percentage']))+(roi['Manual MHT / AHT in Min /Transaction']-(roi['Manual MHT / AHT in Min /Transaction']*roi['calculate_percentage']))*.02



        roi['BOT FTE Required / Day']=roi['Manual FTE Required/Day']-roi['FTE Saved/Day']
        roi['FTE Cost After BOT Implementation(Per Day)']=roi['BOT FTE Required / Day']*(assumptions['Monthly FTE Cost'][0]+assumptions['Monthly Seat Cost'][0]+assumptions['Monthly FTEOther Cost'][0]+assumptions['Monthly Other OPEX'][0])/assumptions['BOT Number of Days Per Month'][0]
        roi['Infra Cost (Per Day)']=(((assumptions['Monthly VM Cost'][0]*assumptions["Number of VM's"][0])/assumptions['Number of Processes'][0])/assumptions['BOT Number of Days Per Month'][0])
        roi['Runner']=(((roi['AverageTransactions(Annual)']/assumptions['BOT Days in a Year'][0])*roi['BOT MHT / AHT (Min)'])/assumptions['Mins Per Hour'][0])*(assumptions['Monthly Runner License Cost'][0]/(assumptions['BOT Number of Days Per Month'][0]*assumptions['BOT Number of Hours in a Day'][0]))

        roi['Creater']=(assumptions['Monthly Creater License Cost'][0]/assumptions['BOT Number of Days Per Month'][0])/assumptions['Creater License Average Out Time'][0]
        roi['OCR']=(assumptions['Document Automation Cost'][0]/assumptions['Number Of Documents per OCR License'][0])*(roi['AverageTransactions(Annual)']/assumptions['BOT Days in a Year'][0])

        roi['checkProduction']=roi['Already in Production'].apply(check_prod)

        roi['check_ocr']=roi['OCR/NON-OCR'].apply(check_ocr)
        roi['OCR']=roi['OCR']*roi['check_ocr']
        roi['Support']=(math.ceil(len(roi['UseCase'])/8)*100000/len(roi['UseCase']))/assumptions['BOT Number of Days Per Month'][0]
        
        # roi['OCR']=roi['OCR']*roi['checkProduction']


        row_cols_to_sum=['FTE Cost After BOT Implementation(Per Day)',
            'Infra Cost (Per Day)', 'Runner', 'OCR', 'Support']

        roi['Total Cost (Excluding Development & Creater Cost)'] = roi[row_cols_to_sum].sum(axis=1)
        roi['ROI in Rupees (Per Day)']=roi['Total Cost (Per Day)']-roi['Total Cost (Excluding Development & Creater Cost)']
        roi['ROI In Percent']=(roi['Total Cost (Per Day)']-roi['Total Cost (Excluding Development & Creater Cost)'])/roi['Total Cost (Per Day)']
        
        # (math.ceil(len(roi['UseCase'])/8*100000)/assumptions['Number of Processes'][0])/assumptions['BOT Number of Days Per Month'][0]

        roi.drop(columns=['calculate_percentage','checkProduction','check_ocr'], inplace=True)

        # Load the original Excel file

        print('load_workbook')
        wb = load_workbook(filepath)
        print('load_workbook donme')
        ws = wb["ROI"]

        # Load the transformed data
        updated_df = roi
        max_row=0

        updated_df = updated_df.replace({np.nan: None})


        for row_idx, row in enumerate(ws.iter_rows(), start=1):
            if row_idx <= 4:  # Skip first 5 rows
                continue
            print(row_idx, row)
            max_row = row_idx

            for cell in row[8:]:  # Skip the first 8 columns (index 0 to 7)
            
                cell.value = None  # Clear the cell value but keep formatting

        # Update only values, keeping formatting intact
        for row_idx, row in enumerate(updated_df.itertuples(index=False), start=5):  
            for col_idx, value in enumerate(row, start=1):
                # print( row_idx, row)
                # print(col_idx, value)
                # break

                ws.cell(row=row_idx, column=col_idx, value=value)


        ws['I'+str(max_row)].value = roi['Manual FTE Required/Day'].sum()
        ws['N'+str(max_row)].value = roi['Total Cost (Per Day)'].sum()
        ws['Q'+str(max_row)].value = roi['BOT FTE Required / Day'].sum()
        ws['R'+str(max_row)].value = roi['FTE Cost After BOT Implementation(Per Day)'].sum()
        ws['Y'+str(max_row)].value = roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ws['Z'+str(max_row)].value = roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ws['AA'+str(max_row)].value = (roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum())/roi['Total Cost (Per Day)'].sum()


        # # do_not write in excel
        # FTE_Saved_Day = roi['FTE Saved/Day'].sum()/len(roi['FTE Saved/Day'])
        # BOT_MHT_AHT_Min = roi['BOT MHT / AHT (Min)'].sum()/len(roi['BOT MHT / AHT (Min)'])
        # # BOT_FTE = roi['BOT FTE'].sum()/len(roi['BOT FTE'])
        # # FTE_Cost= roi['FTE Cost'].sum()/len(roi['FTE Cost'])

        # # ---
        # BOT_FTE_Required_Day=roi['BOT FTE Required / Day'].sum()/len(roi['BOT FTE Required / Day'])
        # FTE_Cost_After_BOT_Implementation=roi['FTE Cost After BOT Implementation(Per Day)'].sum()/len(roi['FTE Cost After BOT Implementation(Per Day)'])
        # Infra_Cost_Per_Day=roi['Infra Cost (Per Day)'].sum()/len(roi['Infra Cost (Per Day)'])
        # License_Cost_Per_Day=roi['Runner'].sum()/len(roi['Runner'])+roi['Creater'].sum()/len(roi['Creater'])+roi['OCR'].sum()/len(roi['OCR'])
        # Development_Support=roi['Development'].sum()/len(roi['Development'])+roi['Support'].sum()/len(roi['Support'])
        # Total_Cost_Excluding_Development_Creater_Cost=(roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum())/roi['Total Cost (Per Day)'].sum()

        # # Save the updated Excel file
        # wb.save("data_modified2.xlsx")
        # print('writtennnnnn')
        # return "data_modified2.xlsx"


        # do_not write in excel
        FTE_Saved_Day = roi['FTE Saved/Day'].sum()/len(roi['FTE Saved/Day'])
        BOT_MHT_AHT_Min = roi['BOT MHT / AHT (Min)'].sum()/len(roi['BOT MHT / AHT (Min)'])
        # BOT_FTE = roi['BOT FTE'].sum()/len(roi['BOT FTE'])
        # FTE_Cost= roi['FTE Cost'].sum()/len(roi['FTE Cost'])

        # ---
        BOT_FTE_Required_Day=roi['BOT FTE Required / Day'].sum()/len(roi['BOT FTE Required / Day'])
        FTE_Cost_After_BOT_Implementation=roi['FTE Cost After BOT Implementation(Per Day)'].sum()
        # FTE_Cost_After_BOT_Implementation=roi['FTE Cost After BOT Implementation(Per Day)'].sum()/len(roi['FTE Cost After BOT Implementation(Per Day)'])
        Infra_Cost_Per_Day=roi['Infra Cost (Per Day)'].sum()/len(roi['Infra Cost (Per Day)'])
        License_Cost_Per_Day=roi['Runner'].sum()/len(roi['Runner'])+roi['Creater'].sum()/len(roi['Creater'])+roi['OCR'].sum()/len(roi['OCR'])
        Development_Support=roi['Development'].sum()/len(roi['Development'])+roi['Support'].sum()/len(roi['Support'])
        Total_Cost_Excluding_Development_Creater_Cost=roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ROI_in_Rupees=roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum()
        ROI_In_Percent=(roi['Total Cost (Per Day)'].sum()-roi['Total Cost (Excluding Development & Creater Cost)'].sum())/roi['Total Cost (Per Day)'].sum()


        
        percentage=calculate_BOT_HT_AHT(roi['Complexity'][0])

        return_dict={

            'FTE_Saved_Day':FTE_Saved_Day,
            'BOT_MHT_AHT_Min':BOT_MHT_AHT_Min,
            # 'BOT_FTE':None,
            # 'FTE_Cost':None,
            'BOT_FTE_Required_Day':BOT_FTE_Required_Day,
            'FTE_Cost_After_BOT_Implementation':FTE_Cost_After_BOT_Implementation,
            'Infra_Cost_Per_Day':Infra_Cost_Per_Day,
            'License_Cost_Per_Day':License_Cost_Per_Day,
            'Development_Support':Development_Support,
            'Total_Cost_Excluding_Development_Creater_Cost':Total_Cost_Excluding_Development_Creater_Cost,
            'ROI_in_Rupees':ROI_in_Rupees,
            'ROI_In_Percent':ROI_In_Percent*100,
            'BOT_FTE_Calculated':roi['Manual FTE Required/Day'].mean()*percentage,
            'Manual_Cost':roi['Total Cost (Per Day)'].sum()
        }

        # Save the updated Excel file
        wb.save("data_modified6.xlsx")
        print('writtennnnnn')
        return ["data_modified6.xlsx",return_dict]

    except:
        return 0
