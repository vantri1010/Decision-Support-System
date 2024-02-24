import json
import xlrd
import numpy as np
import xlsxwriter
import os, os.path
import win32com.client
from pyahp.hierarchy import AHPModel
from pyahp.methods import EigenvalueMethod
from pyahp.hierarchy import AHPCriterion

with open('risk.json') as json_model:
    model = json.load(json_model)
#Insert excel data
wb = xlrd.open_workbook('inputs.xlsx')
sheetActors = wb.sheet_by_index(4)
sheetStructure = wb.sheet_by_index(6)  
sheetTask = wb.sheet_by_index(5)  
sheetTechnology = wb.sheet_by_index(7)    
Actors = []
Structure = []
Task = []
Technology = []
for i in range(16):
    lists=[]
    for j in range(16):
        lists.append(sheetActors.cell(i,j).value)
    Actors.append(lists)
for i in range(8):
    lists=[]
    for j in range(8):
        lists.append(sheetStructure.cell(i,j).value)
    Structure.append(lists)
for i in range(16):
    lists=[]
    for j in range(16):
        lists.append(sheetTask.cell(i,j).value)
    Task.append(lists)
for i in range(9):
    lists=[]
    for j in range(9):
        lists.append(sheetTechnology.cell(i,j).value)
    Technology.append(lists)
model["preferenceMatrices"]["subCriteria:Actors"] = Actors
model["preferenceMatrices"]["subCriteria:Structure"] = Structure
model["preferenceMatrices"]["subCriteria:Task"] = Task
model["preferenceMatrices"]["subCriteria:Technology"] = Technology

solver=EigenvalueMethod #Use EigenvalueMethod to calculate AHP matrix
ahp_model = AHPModel(model, EigenvalueMethod)
preference_matrices = model['preferenceMatrices'] #Dictionary of subCriteria as key and its matrix as values
criteria_list = model.get('criteria') #List of Dimensions
subCriteria_list = model.get('subCriteria') #Dictionary of Dimension as key and Factors as values
criteria = [AHPCriterion(n, model, solver) for n in criteria_list]

#crit_pm = np.array(preference_matrices['criteria'])
#crit_pr = ahp_model.solver.estimate(crit_pm)

crit_attr_pr = [criterion.get_priorities() for criterion in criteria] #calculate priority for each subCriteria in each dimension
#if all criteria is equal (Actors = Structure = Task = Technology) then multiplier is 0.25, else multiply by "crit_pr"
attr_global_pr = [list(0.25* crit_attr_pr[i]) for i in range(len(criteria))]

print(subCriteria_list) #Print list of dimension and factors
print(attr_global_pr) #Print list of priority values of each factor

#Write priority values to excel file.
workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()
row = -1
col = 0
index = -1
for key in subCriteria_list.keys():
    row += 1
    index += 1
    i = 0
    worksheet.write(row, col, key)
    for item in subCriteria_list[key]:
        worksheet.write(row, col + 1, item)
        worksheet.write(row, col + 2, attr_global_pr[index][i]*100)
        row += 1
        i += 1
    row -= 1
workbook.close()

#Run excel VBA to generate chart image.
if os.path.exists("VBA.xlsm"):
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(os.path.abspath("VBA.xlsm"), ReadOnly=0)
    xl.Application.Run("VBA.xlsm!Module1.CreateRadarChart")
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl
