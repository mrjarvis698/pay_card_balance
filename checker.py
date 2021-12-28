import os
from os import path
import shutil
import json
import tkinter
from tkinter import filedialog
import pandas as pd
import json
from openpyxl.workbook import Workbook
from openpyxl import load_workbook

root = tkinter.Tk()
root.withdraw()

# Open xlsx file
open_sheet = path.exists("cache/opened_sheet.json")
if open_sheet == True :
  opened_sheet_file_path = "cache/opened_sheet.json"
  json_file = open(opened_sheet_file_path)
  data = json.load(json_file)
  xlsx_sheet_check = path.exists(data ['xlsx_file_path'])
  if xlsx_sheet_check == True :
    xlsx_file_path = data ['xlsx_file_path']
  else :
    shutil.rmtree('cache', ignore_errors=True)
    xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
    cache_path = os.path.join(str(os.getcwd()), "cache")
    dictionary = {"xlsx_file_path" : xlsx_file_path}
    json_object = json.dumps(dictionary, indent = 1)
    with open("cache/opened_sheet.json", "w") as outfile:
      outfile.write(json_object)
else :
  xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
  cache_path = os.path.join(str(os.getcwd()), "cache")
  os.mkdir(cache_path)
  dictionary = {"xlsx_file_path" : xlsx_file_path}
  json_object = json.dumps(dictionary, indent = 1)
  with open("cache/opened_sheet.json", "w") as outfile:
    outfile.write(json_object)

# read imported xlsx file path using pandas
input_workbook = pd.read_excel(xlsx_file_path, sheet_name = 'Sheet1', usecols = 'A:F', dtype=str)
input_workbook.head()

# read total number of rows present in xlsx
number_of_rows = len(input_workbook.index)
print('Total Cards = ',number_of_rows)

input_workbook_cc_number = input_workbook['Cards'].values.tolist()
input_workbook_cvv_number = input_workbook['CVV'].values.tolist()
input_workbook_expiry_number = input_workbook['Exp date'].values.tolist()
input_workbook_kit_number = input_workbook['KIT NUMBER'].values.tolist()
input_workbook_mobile_number = input_workbook['Mobile Number'].values.tolist()
input_workbook_balance_number = input_workbook['Balance'].values.tolist()

# get-output sheet to append output
output_sheet = path.exists("Output.xlsx")
if output_sheet == True :
  output_sheet_file_path = "Output.xlsx"
else :
  output_headers= ["Cards", 'CVV', 'Exp date', 'KIT NUMBER','Mobile Number', 'Balance', 'Current Balance']
  overall_output = Workbook()
  page = overall_output.active
  page.append(output_headers)
  overall_output.save(filename = 'Output.xlsx')
  output_sheet_file_path = "Output.xlsx"

def output_save():
  entry_list = [[ input_workbook_cc_number[x], input_workbook_cvv_number[x], input_workbook_cvv_number[x], input_workbook_expiry_number[x], input_workbook_kit_number[x], input_workbook_mobile_number[x], input_workbook_balance_number[x]]]
  output_wb = load_workbook(output_sheet_file_path)
  page = output_wb.active
  for info in entry_list:
      page.append(info)
  output_wb.save(filename='Output.xlsx')

def resume_output():
  global output_cc_number
  global h
  output_load_wb_2 = pd.read_excel(output_sheet_file_path, sheet_name = 'Sheet', usecols = 'A', dtype=int)
  output_load_wb_2.head()
  output_cc_number = output_load_wb_2['Cards'].values.tolist()
  h = len(output_load_wb_2.index)

try:
  resume_output()
except IndexError:
  for x in range (0 , number_of_rows):
      output_save()
else:
  last_txncard = h
  for x in range (last_txncard , number_of_rows):
      output_save()
