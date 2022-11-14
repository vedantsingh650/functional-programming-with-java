
import pandas as pd
from datetime import datetime
import datetime
from datetime import timedelta
import string
 
# read by default 1st sheet of an excel file
# df = pd.read_excel('Adaptive Dock Scheduler.xlsx')
df = pd.read_excel('Rishit_file_Testing.xlsx')
#Rishit_file_Testing
# Below is data cleaning (Kind of Custom Logic not required always and below is hard-coded values)
df.drop(df.columns[[5,6,7,8]], axis=1, inplace=True)
# print(df.columns)

# Making list of time to added as data in time column
#####################
date_time_str = datetime.datetime.strptime('06:00:00', '%H:%M:%S')
date_time_end = datetime.datetime.strptime('21:00:00', '%H:%M:%S')
timee=[]

while date_time_str < date_time_end :
 date_time_str = date_time_str + timedelta(minutes=30)
 (timee.append (str(date_time_str.time())))
#####################

# Making list of DOCKX
# ############
Total_Dock=11  # Actual dock is 10
dock=[]
for x in range(1,Total_Dock):
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
    ful_time=df['ETA'][x].time()
    datetime_obj = datetime.datetime.strptime(str(ful_time), '%H:%M:%S')
    each_time30= datetime_obj + timedelta(minutes=30)
    
    hr=each_time30.time().hour
    str_hr=str(hr).zfill(2)
    first_hlf_con=((str_hr + ':00:00' <= str(each_time30.time())) and (str_hr + ':30:00' >= str(each_time30.time())) )
    second_hlf_con=((str_hr + ':30:00' < str(each_time30.time())) and (str(hr+1) + ':00:00' >= str(each_time30.time())) )
    dock=1
    loop=True
    if first_hlf_con:
        ind_tim=(new_df[new_df['Time']== str_hr + ':00:00'].index.values)
    
    if second_hlf_con:
        ind_tim=(new_df[new_df['Time']== str_hr + ':30:00'].index.values)
        
    while loop:
      con=(new_df['Dock ' + str(dock)][ind_tim] == "").bool()  & (new_df['Dock ' + str(dock)][ind_tim+1] == "").bool() 
      con2=(new_df['Dock ' + str(dock)][ind_tim-1] == "").bool() 
      if(con & con2):
        loop=False
        # print(dock)
        data = '\n'.join([str(df['From'][x]), str(df['Shipment No'][x])])
        new_df['Dock ' + str(dock)][ind_tim] = data 
        new_df['Dock ' + str(dock)][ind_tim+1] = "merge" 
      dock+=1  
      
print("***********CORE Completed********************************************")   

# print(new_df)

#  Beautification Work
startrow=4
writer = pd.ExcelWriter("test2.xlsx", engine="xlsxwriter")

new_df.to_excel(writer, startrow=startrow, startcol=0,index = False,sheet_name='python')
To=str(df['To'][1]) + " is the 'To' Location"
worksheet = writer.sheets['python']
workbook = writer.book

merge_format = workbook.add_format({'align': 'center','border': 2,'bold': True, 'font_color': 'red'})
merge_format.set_align('vcenter')
wrap_format = workbook.add_format({'text_wrap': True})
worksheet.set_column("B:E", None, wrap_format)
worksheet.merge_range('A1:F3',To, merge_format)



# worksheet.merge_range(2, 1, 3, 3, 'Merged Cells', merge_format)

# print(new_df)
# print( (new_df[new_df.eq("merge").any(1)]) )




merge_cell = workbook.add_format({'align': 'center','text_wrap': True})
for x in range(1,Total_Dock):
    data=("Dock " + str(x))
    all_index=(new_df[new_df[data]== 'merge'].index.values)
    for index in all_index:
        Uppercase_col=string.ascii_uppercase[x] 
        sheet_index= startrow+2+index
        mergecell=Uppercase_col + str(sheet_index-1) + ":" + Uppercase_col + str(sheet_index)
        worksheet.merge_range(mergecell,new_df[data][index-1], merge_cell)
        print(data +" "+ str(sheet_index) + " " + new_df[data][index-1])


writer.save()
