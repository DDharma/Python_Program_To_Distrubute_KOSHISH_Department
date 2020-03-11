import pandas as pd 
import numpy as np 

#Creating variables 
KOSHISH_ART_n_CRAFT = []
KOSHISH_EVENT_CRAFTER = []
KOSHISH_KURE = []
KOSHISH_FINANCE = []
KOSHISH_PATHSHALA = []
KOSHISH_COLLECTOR_n_DISTRIBUTORS = []
KOSHISH_AIMER = []


#Reading the excel sheet and save the data frame
df = pd.read_excel('depart_pref.xlsx')
df = np.array(df)
'''print(df.shape)
print(df[9][2])'''

#Retriving Elements From the 1st preference
for i in range(35):
    if df[i][1] ==  "KOSHISH ART n CRAFT":
        KOSHISH_ART_n_CRAFT.append(df[i][0])
    elif df[i][1] == "KOSHISH EVENT CRAFTER":
        KOSHISH_EVENT_CRAFTER.append(df[i][0]) 
    elif df[i][1] == "KOSHISH KURE":
        KOSHISH_KURE.append(df[i][0])
    elif df[i][1] == "KOSHISH FINANCE":
        KOSHISH_FINANCE.append(df[i][0])
    elif df[i][1] == "KOSHISH PATHSHALA":
        KOSHISH_PATHSHALA.append(df[i][0])
        print('KOSHISH PATHSHALA                :',KOSHISH_PATHSHALA)
    elif df[i][1] == "KOSHISH COLLECTOR n DISTRIBUTORS":
        KOSHISH_COLLECTOR_n_DISTRIBUTORS.append(df[i][0])
    elif df[i][1] == "KOSHISH AIMER":
        KOSHISH_AIMER.append(df[i][0])


#Retriving Elements From the 2st preference
for j in range(35):
    if df[j][2] ==  "KOSHISH ART n CRAFT":
        KOSHISH_ART_n_CRAFT.append(df[j][0])
    elif df[j][2] == "KOSHISH EVENT CRAFTER":
        KOSHISH_EVENT_CRAFTER.append(df[j][0]) 
    elif df[j][2] == "KOSHISH KURE":
        KOSHISH_KURE.append(df[j][0])
    elif df[j][2] == "KOSHISH FINANCE":
        KOSHISH_FINANCE.append(df[j][0])
    elif df[j][2] == "KOSHISH PATHSALA":
        KOSHISH_PATHSHALA.append(df[j][0])
        print('KOSHISH PATHSHALA                :',KOSHISH_PATHSHALA)
    elif df[j][2] == "KOSHISH COLLECTOR n DISTRIBUTORS":
        KOSHISH_COLLECTOR_n_DISTRIBUTORS.append(df[j][0])
    elif df[j][2] == "KOSHISH AIMER":
        KOSHISH_AIMER.append(df[j][0])

#retriving unique element into the list
KOSHISH_ART_n_CRAFT = list(np.unique(KOSHISH_ART_n_CRAFT))
KOSHISH_COLLECTOR_n_DISTRIBUTORS = list(np.unique(KOSHISH_COLLECTOR_n_DISTRIBUTORS))
KOSHISH_EVENT_CRAFTER = list(np.unique(KOSHISH_EVENT_CRAFTER))
KOSHISH_AIMER = list(np.unique(KOSHISH_AIMER))
KOSHISH_FINANCE = list(np.unique(KOSHISH_FINANCE))
KOSHISH_PATHSHALA = list(np.unique(KOSHISH_PATHSHALA))
KOSHISH_FINANCE = list(np.unique(KOSHISH_FINANCE))


#printig all the retive data
print('KOSHISH ART n CRAFT              :',KOSHISH_ART_n_CRAFT)
print('KOSHISH EVENT CRAFTER            :',KOSHISH_EVENT_CRAFTER)
print('KOSHISH KURE                     :',KOSHISH_KURE)
print('KOSHISH FINANCE                  :',KOSHISH_FINANCE)
print('KOSHISH PATHSHALA                :',KOSHISH_PATHSHALA)
print('KOSHISH COLLECTOR n DISTRIBUTORS :',KOSHISH_COLLECTOR_n_DISTRIBUTORS)
print('KOSHISH AIMER                    :',KOSHISH_AIMER)


#Making data frame for excelsheet
df_write1 = pd.DataFrame({'KOSHISH ART n CRAFT': KOSHISH_ART_n_CRAFT})
df_write2 = pd.DataFrame({'KOSHISH EVENT CRAFTER': KOSHISH_EVENT_CRAFTER})
df_write3 = pd.DataFrame({'KOSHISH KURE':KOSHISH_KURE})
df_write4 = pd.DataFrame({'KOSHISH FINANCE':KOSHISH_FINANCE})
df_write5 = pd.DataFrame({'KOSHISH PATHSHALA':KOSHISH_PATHSHALA})
df_write6 = pd.DataFrame({'KOSHISH COLLECTOR n DISTRIBUTORS':KOSHISH_COLLECTOR_n_DISTRIBUTORS})
df_write7 = pd.DataFrame({'KOSHISH AIMER':KOSHISH_AIMER})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('koshish_department.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df_write1.to_excel(writer, sheet_name='Sheet1', startcol=1)
df_write2.to_excel(writer, sheet_name='Sheet1', startcol=3)
df_write3.to_excel(writer, sheet_name='Sheet1', startcol=5)
df_write4.to_excel(writer, sheet_name='Sheet1', startcol=7)
df_write5.to_excel(writer, sheet_name='Sheet1', startcol=9)
df_write6.to_excel(writer, sheet_name='Sheet1', startcol=11)
df_write7.to_excel(writer, sheet_name='Sheet1', startcol=13)

# Close the Pandas Excel writer and output the Excel file.
writer.save()