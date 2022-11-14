
import pandas as pd
from datetime import datetime
import datetime
from datetime import timedelta
from xlsxwriter import Workbook
# read by default 1st sheet of an excel file
# df = pd.read_excel('Adaptive Dock Scheduler.xlsx')
df = pd.read_excel('Rishit_file_Testing.xlsx')
#Rishit_file_Testing
df.drop(df.columns[[5,6,7,8]], axis=1, inplace=True)
# print(df.columns)

# Making list of time to added as data in time column
#####################
date_time_str = datetime.datetime.strptime('06:30:00', '%H:%M:%S')
date_time_end = datetime.datetime.strptime('21:00:00', '%H:%M:%S')
timee=[]

while date_time_str < date_time_end :
 date_time_str = date_time_str + timedelta(minutes=30)
 (timee.append (str(date_time_str.time())))
#####################

# Making list of DOCKX
# ############
dock=[]
for x in range(1, 11):
 data=("Dock " + str(x))
 dock.append(data)   
# ############

# Structure of new Dataframe
#################
new_df = pd.DataFrame(timee, columns=['Time'])
new_df[dock] = ''
# #################

# Implementation Started LOgically
#####################################
excle_len=len(df)
print("INSIDE CORE ******************************")
for x in range(0,excle_len):
    hr=df['ETA'][x].time().hour
    str_hr=str(hr).zfill(2)
    ful_time=df['ETA'][x].time()
    first_hlf_con=((str_hr + ':00:00' <= str(ful_time)) and (str_hr + ':30:00' >= str(ful_time)) )
    second_hlf_con=((str_hr + ':30:00' < str(ful_time)) and (str(hr+1) + ':00:00' >= str(ful_time)) )
    dock=1
    loop=True
    if first_hlf_con:
        ind_tim=(new_df[new_df['Time']== str_hr + ':00:00'].index.values)
    
    if second_hlf_con:
        ind_tim=(new_df[new_df['Time']== str_hr + ':30:00'].index.values)
        
    while loop:
      con=(new_df['Dock ' + str(dock)][ind_tim] == "").bool()  & (new_df['Dock ' + str(dock)][ind_tim+1] == "").bool() 
      con2=(ind_tim >0 and ind_tim<11)  and (new_df['Dock ' + str(dock)][ind_tim-1] == "").bool() 
      if(con & con2):
        loop=False
        print(dock)
        new_df['Dock ' + str(dock)][ind_tim] = df['From'][x]
      dock+=1  
      
print("***********CORE Completed********************************************")   

print(new_df)
# new_df.to_excel('pandas_to_excel.xlsx', sheet_name='python',startrow=1,index = False)


#  Beautification Work
writer = pd.ExcelWriter("test.xlsx", engine="xlsxwriter")
new_df.to_excel(writer, startrow=4, startcol=0,index = False)
To=str(df['To'][1]) + " is the 'To' Location"
worksheet = writer.sheets['Sheet1']
workbook = writer.book

merge_format = workbook.add_format({'align': 'center','border': 2})
# worksheet.merge_range(2, 1, 3, 3, 'Merged Cells', merge_format)
worksheet.merge_range('A1:F3',To, merge_format)
writer.save()