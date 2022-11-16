
import pandas as pd
from datetime import datetime
import datetime
from datetime import timedelta
import string
 
# read by default 1st sheet of an excel file
# raw_df = pd.read_excel('Adaptive Dock Scheduler.xlsx')
# raw_df = pd.read_excel('Rishit_file_Testing.xlsx')
raw_df = pd.read_excel('Rishit_file_Testing_huge.xlsx')  

# Below is data cleaning (Kind of Custom Logic not required always and below is hard-coded values)
raw_df.drop(raw_df.columns[[5,6,7,8]], axis=1, inplace=True)
# print(df.columns)

df = raw_df.sort_values(by=['ETA'], ascending=True)
df = df.reset_index(drop = True)

# Making list of time to added as data in time column
#####################
start_op_time='06:30:00' #exclusive in sheet excell
end_op_time='21:00:00'  #inclusive in sheet excell
date_time_str = datetime.datetime.strptime(start_op_time, '%H:%M:%S')
date_time_end = datetime.datetime.strptime(end_op_time, '%H:%M:%S')
timee=[]

while date_time_str < date_time_end :
 date_time_str = date_time_str + timedelta(minutes=30)
 (timee.append (str(date_time_str.time())))
#####################

# Making list of DOCK X
# ############
Total_Dock=11  # Actual dock is 10
dock=[]
for x in range(1,Total_Dock):
 data=("Dock " + str(x))
 dock.append(data)   
# ############

# Structure of new Dataframe
#################
new_df = pd.DataFrame(timee, columns=['Oper_Time'])
new_df[dock] = ''
# #################

dock_assign_index=[]

# print(new_df['Oper_Time'][2])

# Implementation Started LOgically
#####################################
excle_len=len(df)
print("INSIDE CORE ******************************")
for x in range(0,excle_len):
    ful_time=df['ETA'][x].time()
    datetime_obj = datetime.datetime.strptime(str(ful_time), '%H:%M:%S')
    each_time30= datetime_obj + timedelta(minutes=30)#20:31
    
    hr=each_time30.time().hour
    str_hr=str(hr).zfill(2) 
    first_hlf_con=((str_hr + ':00:00' <= str(each_time30.time())) and (str_hr + ':30:00' >= str(each_time30.time())) )
    second_hlf_con=((str_hr + ':30:00' < str(each_time30.time())) and (str(hr+1) + ':00:00' >= str(each_time30.time())) )
    dock=1
    loop=True
    if first_hlf_con:
        ind_tim=(new_df[new_df['Oper_Time']== str_hr + ':00:00'].index.values)
    
    if second_hlf_con:
        ind_tim=(new_df[new_df['Oper_Time']== str_hr + ':30:00'].index.values)
        
    while loop:
      con=(new_df['Dock ' + str(dock)][ind_tim] == "").bool()  # & (new_df['Dock ' + str(dock)][ind_tim+1] == "").bool() 
       #  con2=(new_df['Dock ' + str(dock)][ind_tim-1] == "").bool() 
      if(con):
        loop=False
        # print(dock)
        data = '\n'.join([str(df['From'][x]), str(df['Shipment No'][x])])
        new_df['Dock ' + str(dock)][ind_tim] = data 
        dock_assign_index.append(x) 
      dock+=1

      if(dock>=Total_Dock):
        if ind_tim+1<len(timee):
            dock=1
            ind_tim=ind_tim+1
        if ind_tim+1>=len(timee):
            loop=False

      
print("***********CORE Completed********************************************")   

# print(new_df)

#  Beautification Work
startrow=4
writer = pd.ExcelWriter("hulk.xlsx", engine="xlsxwriter")

new_df.to_excel(writer, startrow=startrow, startcol=0,index = False,sheet_name='python')
To=str(df['To'][1]) + " is the 'To' Location"
worksheet = writer.sheets['python']
workbook = writer.book

merge_format = workbook.add_format({'align': 'center','border': 2,'bold': True, 'font_color': 'red'})
merge_format.set_align('vcenter')
wrap_format = workbook.add_format({'text_wrap': True})
worksheet.set_column("A:L", None, wrap_format)
worksheet.merge_range('A1:F3',To, merge_format)

#####################################
###### Creating Unassigned data sheet###############

unassigned_df = df
for x in dock_assign_index:
    unassigned_df = unassigned_df.drop(x)

if not unassigned_df.empty:
    unassigned_df.to_excel(writer, sheet_name="unassigned", index=False)

###################################

writer.close()

# HTML file creation
##############################################
html = new_df.to_html()
html = html.replace("\\n","<br>")
html = html.replace('class="dataframe"', 'class="table table-striped table-hover"')
boostrap_link = '<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">'

Heading = '<center><h1 style="background-color:powderblue;">{To}</h2></center>'.format(To=To)
text_file = open("work.html", "w")
text_file.write(Heading + '\n' +  boostrap_link + '\n' + html)
text_file.close()
##############################################


# worksheet.merge_range(2, 1, 3, 3, 'Merged Cells', merge_format)

# print(new_df)
# print( (new_df[new_df.eq("merge").any(1)]) )

# Below is for mergring cell in excell for DOCK (2 cells    )
 
# merge_cell = workbook.add_format({'align': 'center','text_wrap': True})
# for x in range(1,Total_Dock):
#     data=("Dock " + str(x))
#     all_index=(new_df[new_df[data]== 'merge'].index.values)
#     for index in all_index:
#         Uppercase_col=string.ascii_uppercase[x] 
#         sheet_index= startrow+2+index
#         mergecell=Uppercase_col + str(sheet_index-1) + ":" + Uppercase_col + str(sheet_index)
#         worksheet.merge_range(mergecell,new_df[data][index-1], merge_cell)
#         print(data +" "+ str(sheet_index) + " " + new_df[data][index-1])



