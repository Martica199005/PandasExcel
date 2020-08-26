import pandas as pd
import os
import sys
import io
import xlsxwriter
from datetime import datetime


link_excel=os.getenv('LINK_EXCEL')

def writetxt(filename): 
  original_stdout = sys.stdout 
  with open(filename, 'w') as f:
    sys.stdout =f 
    for elem in flowId:
      print(elem)
      rows=excel.loc[excel['FlowID'] == elem]
      rows_group=rows.groupby([rows['ImportFolderType'].fillna('0'),rows['ProviderID']])['Number'].unique()
      print(rows_group)
      print('\n')
    sys.stdout = original_stdout 
    # Reset the standard output to its original value


def write_excel(filename, list1):
  writer = pd.ExcelWriter(filename, engine='xlsxwriter')
  workbook = writer.book
  start_row = 0
  for df in list1:
    df.to_excel(writer, sheet_name='Sheet 1', index=False, startrow=start_row)
    start_row = start_row + len(df) + 2
  writer.save()
  workbook.close()  



#print(sys.version) #check python version
#python3 -m pip install xlrd


#print('Link excel file: '+link_excel)

excel=pd.read_excel('Result_numberOfFiles_flows.xlsx')


list_col=list(excel.columns)
#print(list_col)
print(list_col)
print('\n')

list_df=[]



flowId=list(excel['FlowID'].drop_duplicates())

for elem in flowId:
  #print(elem)
  rows=excel.loc[excel['FlowID'] == elem]
  rows_group=rows.groupby([rows['ImportFolderType'].fillna('Na'),rows['ProviderID']])['Number'].unique()
  df1=rows_group.reset_index()
  df1.insert(0,"FlowID", elem, True)
  list_df.append(df1)
  #print(type(df1))

list_df2=[list_df[0],list_df[1],list_df[2]]

# datetime object containing current date and time
now = datetime.now()
 
print("now =", now)

# dd/mm/YY H:M:S
dt_string = now.strftime("%d/%m/%Y_%H-%M-%S")
print("date and time =", dt_string)
filename='output'+dt_string+'.xlsx'
print(filename)
#write_excel(,list_df)

 
 
















