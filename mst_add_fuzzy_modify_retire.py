# -*- coding: utf-8 -*-
"""
Created on Tue Mar  5 23:43:26 2019

@author: p.d.ramakrishnan
"""
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import difflib as d

mesu=pd.read_excel(r"Updates",sheet_name='Add',header=1)

mesu_add=mesu.loc[:,['ID','Name ','Definition']]


mesu_add_name=pd.DataFrame(mesu_add["Name "].str.lower())

mst_aliases=pd.read_excel(r"Lookup",sheetname='ALIASES')

mst_aliases=mst_aliases.loc[:,['Name']]

mst_aliases=mst_aliases["Name"].str.lower()

mst_aliases_1=mst_aliases.tolist()



def to_get_close_matches(name):
    print(name)
    #name='flutter sdk'
    a=d.get_close_matches(name,aliases_1,cutoff=0.5)
    a=pd.DataFrame(a)
    a=a.rename(columns={0:"Matches"})
    a["New value"]=""
    a.at[:,"New value"]=name
    a=a.loc[:,['New value','Matches']]
    return(a)
        


    
master=pd.DataFrame()
for index,row in mesu_add_name.iterrows():
    master=master.append(to_get_close_matches(row["Name "]))
    
mst_aliases=pd.read_excel(r"Lookup",sheet_name='ALIASES')


mst_aliases=mst_aliases.loc[:,["ID","Type","Name"]]
mst_aliases["Name"]=mst_aliases["Name"].str.lower()
master_1=master.merge(mst_aliases,left_on="Matches",right_on="Name",how='left')

master_1=master_1.drop_duplicates()

#analyse the matches received


#MSLT
from datetime import date

mslt=pd.read_excel(r"Lookup.xlsx",sheet_name='Lookup')

# Add tab after analysing the results from the fuzzy lookup

#Remove NHS  ids
mslt=mslt[mslt.ID!='']





#Decommission existing ids
mslt.at[mslt["ID"]==,"Action Type"]='Decommisioned'
mslt.at[mslt["ID"]==,["De-duplicated active names"]]='Do not consider'
mslt.at[mslt["ID"]==,["Updates Date"]]=str(date.today())
mslt.at[mslt["ID"]==,["Update Comments"]]='Decommisioned as per EOM report for Aug'




#Add all the new skills
for index,row in mesu.iterrows():
    mslt=mslt.append({"ID":row["ID"],"Active Name":row["Name "],
                 "Definition":row["Definition"],"Release Date":row["Release Date"],
                 "De-duplicated active names":"Consider","Action Type":"Add","Updates Date":date.today(),
                 "Update Comments":"Add","Skill Type":"Skill","Level":5},ignore_index=True)








# Modify tab


mesu_modify=pd.read_excel(r"Updates",sheet_name='Modify',header=1)

mesu_modify_nameanddef=mesu_modify.loc[mesu_modify["Action"]=="Modify Name & Definition"]

mesu_modify_name=mesu_modify.loc[mesu_modify["Action"]=="Modify Name"]

mesu_modify_definition=mesu_modify.loc[mesu_modify["Action"]=="Modify Definition"]

# Modify Definition
for index,row in mesu_modify_definition.iterrows():
    mslt.at[mslt["ID"]==row["ID"],"Definition"]=row["Definition"]
    mslt.at[mslt["ID"]==row["ID"],"Updates Date"]=date.today()
    mslt.at[mslt["ID"]==row["ID"],"Update Comments"]="Modify Definition as per EOM report"
    mslt.at[mslt["ID"]==row["ID"],"Action Type"]="Modify"

#Modify Name
for index,row in mesu_modify_name.iterrows():
    mslt.at[mslt["ID"]==row["ID"],"Active Name"]=row["New Value Name"]
    mslt.at[mslt["ID"]==row["ID"],"Updates Date"]=date.today()
    mslt.at[mslt["ID"]==row["ID"],"Update Comments"]="Modify Name as per EOM report"
    mslt.at[mslt["ID"]==row["ID"],"Action Type"]="Modify"
           
    
        
#Modify name and definition
for index,row in mesu_modify_nameanddef.iterrows():
    mslt.at[mslt["ID"]==row["ID"],"Active Name"]=row["New Value Name"]
    mslt.at[mslt["ID"]==row["ID"],"Skill Definition"]=row["New Definition"]
    mslt.at[mslt["ID"]==row["ID"],"Updates Date"]=date.today()
    mslt.at[mslt["ID"]==row["ID"],"Update Comments"]="Modify Name and Definition as per EOM report"
    mslt.at[mslt["ID"]==row["ID"],"Action Type"]="Modify"
    
  
        
        
#checks
test=mslt.loc[mslt['Updates Date']==date.today()]



#Retire
mesu_retire=pd.read_excel(r"Updates",sheet_name='Retire',header=1)

for index,row in mesu_retire.iterrows():
    mslt.at[mslt["ID"]==row["ID"],"Updates Date"]=date.today()
    mslt.at[mslt["ID"]==row["ID"],"Update Comments"]="Retire"
    mslt.at[mslt["ID"]==row["ID"],"Action Type"]="Retire"

#checks
test=mslt.loc[(mslt['Updates Date']==date.today()) & (mslt['Action Type']=='Retire')]
###Alias tab
mst_alias_tab1=mslt.loc[:,["ID","Active Name","Action Type","Level"]]
mst_alias_tab1["Name Type"]="Active Name"

mst_alias_tab1=mst_alias_tab1.rename(columns={'Active Name':'Name'})

mst_alias_tab2=mslt.loc[:,["ID", 'Alternate Old Names',
 'Alternate Old Names 2',
 'Alternate Old Names 3',
 'Alternate Old Names 4',
 'Alternate Old Names 5',
 'Alternate Old Names 6',
 'Alternate Old Names 7',
 'Alternate Old Names 8',
 'Alternate Old Names 9',
 'Alternate Old Names 10',
 'Alternate Old Names 11',
 'Alternate Old Names 12',
 'Alternate Old Names 13',
 'Alternate Old Names 14',
 'Alternate Old Names 15',
 'Alternate Old Names 16',
 'Alternate Old Names 17',
 'Alternate Old Names 18',
 'Alternate Old Names 19',
 'Alternate Old Names 20']]



a=pd.melt(mst_alias_tab2,id_vars=["ID"],value_vars=['Alternate Old Names',
 'Alternate Old Names 2',
 'Alternate Old Names 3',
 'Alternate Old Names 4',
 'Alternate Old Names 5',
 'Alternate Old Names 6',
 'Alternate Old Names 7',
 'Alternate Old Names 8',
 'Alternate Old Names 9',
 'Alternate Old Names 10',
 'Alternate Old Names 11',
 'Alternate Old Names 12',
 'Alternate Old Names 13',
 'Alternate Old Names 14',
 'Alternate Old Names 15',
 'Alternate Old Names 16',
 'Alternate Old Names 17',
 'Alternate Old Names 18',
 'Alternate Old Names 19',
 'Alternate Old Names 20'])

a=a[pd.notnull(a["value"])]

a=a.rename(columns={'variable':'Name Type','value':'Skill Name'})
a["Name Type"]="Alias"

mslt_level_actiontype=mslt.loc[:,["ID","Level","Action Type"]]

mst_alias_tab3=a.merge(mslt_level_actiontype,left_on="ID",right_on="ID",how='left')

mst_alias_final=mst_alias_tab1.append(mst_alias_tab3)


###
##MST_Original
mst_original=mslt.loc[:,['ID',
 'Action Type',
 'Active Name',
 'Name',
 'Skill Type',
 'Skill Flag',
 'Level',
 'Source',
 'Updates Date',
 'Update Comments',
 'Definition',
 'Release Date',
 'Alternate Old Names',
 'Alternate Old Names 2',
 'Alternate Old Names 3',
 'Alternate Old Names 4',
 'Alternate Old Names 5',
 'Alternate Old Names 6',
 'Alternate Old Names 7',
 'Alternate Old Names 8',
 'Alternate Old Names 9',
 'Alternate Old Names 10',
 'Alternate Old Names 11',
 'Alternate Old Names 12',
 'Alternate Old Names 13',
 'Alternate Old Names 14',
 'Alternate Old Names 15',
 'Alternate Old Names 16',
 'Alternate Old Names 17',
 'Alternate Old Names 18',
 'Alternate Old Names 19',
 'Alternate Old Names 20']]


## 
##MST_NHS
mst_nhs=mslt.loc[mslt["De-duplicated active names"]=="Consider"]
mst_nhs=mst_nhs.loc[:,['ID',
 'Action Type',
 'Active Name',
 'Security Skill Name',
 'Skill Type',
 'Skill Flag',
 'Level',
 'Source',
 'Updates Date',
 'Update Comments',
 'Skill Definition',
 'Release Date',
 'Alternate Old Names',
 'Alternate Old Names 2',
 'Alternate Old Names 3',
 'Alternate Old Names 4',
 'Alternate Old Names 5',
 'Alternate Old Names 6',
 'Alternate Old Names 7',
 'Alternate Old Names 8',
 'Alternate Old Names 9',
 'Alternate Old Names 10',
 'Alternate Old Names 11',
 'Alternate Old Names 12',
 'Alternate Old Names 13',
 'Alternate Old Names 14',
 'Alternate Old Names 15',
 'Alternate Old Names 16',
 'Alternate Old Names 17',
 'Alternate Old Names 18',
 'Alternate Old Names 19',
 'Alternate Old Names 20']]

##
##MST_DL
mst_dl=mst_nhs.loc[:,['ID',
 'Action Type',
 'Active Name',
 'BG Data Available',
 'Skill Type',
 'Skill Flag',
 'Level',
 'Source',
 'Skill Definition',
 'Alternate Old Names',
 'Alternate Old Names 2',
 'Alternate Old Names 3',
 'Alternate Old Names 4',
 'Alternate Old Names 5',
 'Alternate Old Names 6',
 'Alternate Old Names 7',
 'Alternate Old Names 8',
 'Alternate Old Names 9',
 'Alternate Old Names 10',
 'Alternate Old Names 11',
 'Alternate Old Names 12',
 'Alternate Old Names 13',
 'Alternate Old Names 14',
 'Alternate Old Names 15',
 'Alternate Old Names 16',
 'Alternate Old Names 17',
 'Alternate Old Names 18',
 'Alternate Old Names 19',
 'Alternate Old Names 20']]

mst_dl["Alternate Old Names 21"]=''
mst_dl["Alternate Old Names 22"]=''
mst_dl["Alternate Old Names 23"]=''
mst_dl["Alternate Old Names 24"]=''
mst_dl["Alternate Old Names 25"]=''





writer = pd.ExcelWriter(r'Lookup', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
mslt.to_excel(writer, sheet_name='Lookup',index=False)
mst_alias_final.to_excel(writer, sheet_name='ALIASES',index=False)
mst_original.to_excel(writer, sheet_name='Original',index=False)
mst_dl.to_excel(writer,sheet_name='DL',index=False)
mst_nhs.to_excel(writer,sheet_name='NHS',index=False)

writer.save()










