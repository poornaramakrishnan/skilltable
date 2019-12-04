# -*- coding: utf-8 -*-
"""
Created on Tue Mar  5 23:43:26 2019

@author: p.d.ramakrishnan
"""
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import difflib as d

mesu=pd.read_excel(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\MST October\MST November\Month End Skill Updates October 2019.xlsx",sheet_name='Add',header=1)

mesu_add=mesu.loc[:,['Skill SAP ID','Skill Name ','Skill Definition']]


mesu_add_name=pd.DataFrame(mesu_add["Skill Name "].str.lower())

mst_aliases=pd.read_excel(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\MST October\MST November\NHS Team Master Skill Lookup_Oct2019.xlsx",sheetname='MST_ALIASES')

mst_aliases=mst_aliases.loc[:,['Skill Name']]

mst_aliases=mst_aliases["Skill Name"].str.lower()

mst_aliases_1=mst_aliases.tolist()



#a=d.get_close_matches('.net',['microsoft .net architecture',	'microsoft .net compact framework',	'microsoft .net framework',	'.net programming language',	'asp.net',	'asp.net ajax',	'k2.net',	'microsoft-asp.net',	'ado.net',	'asp.net',	'visual basic (excluding vb.net & vba)',	'visual basic.net programming language',	'asp.net mvc',	'openlink rightangle .net',	'openlink endur oc.net',	'odp.net',	'spring.net',	'alm.net',	'aforge.net',	'csla .net',	'microsoft ml.net',	'visual basic .net',	'asp.net-ajax',	'vb .net',	'.net development',	'.net architecture',	'.net compact framework',	'.net framework',	'.net programming',	'introscope for .net',	'.net',	'sap-.net connector',	'visual basic.net',	'asp.net platform',	'mobility applications backend server side (java/.net)',	'microsoft- .net architecture',	'.net operation',	'vb.net',	'microsoft.net',	'microsoft .net compact framework',	'portal integration frameworks (java portlet/.net web parts/wsrp/others)',	'mobility applications - backend server side (java/.net)',	'microsoft .net',	'microsoft .net framework',	'.net architect',	'asp .net',	'asp.net5',	'.net framework 3.5',	'.net framework 2.0',	'cruisecontrol.net',	'asp.net-mvc',	'.net framework 1.1'],cutoff=0.5)


def to_get_close_matches(name):
    print(name)
    #name='flutter sdk'
    a=d.get_close_matches(name,mst_aliases_1,cutoff=0.5)
    a=pd.DataFrame(a)
    a=a.rename(columns={0:"Matches"})
    a["New value"]=""
    a.at[:,"New value"]=name
    a=a.loc[:,['New value','Matches']]
    return(a)
        

#global master
#master=pd.DataFrame()
#master=master.append(mesu_add_name["Value Name"].apply(to_get_close_matches))
#
#m=mesu_add_name.apply(to_get_close_matches,axis=1).apply()
#
#m=mesu_add_name["Value Name"].apply(to_get_close_matches)
#m=master.T
    
master=pd.DataFrame()
for index,row in mesu_add_name.iterrows():
    master=master.append(to_get_close_matches(row["Skill Name "]))
    
mst_aliases=pd.read_excel(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\MST October\MST November\NHS Team Master Skill Lookup_Oct2019.xlsx",sheet_name='MST_ALIASES')


mst_aliases=mst_aliases.loc[:,["ID","Action Type","Skill Name"]]
mst_aliases["Skill Name"]=mst_aliases["Skill Name"].str.lower()
master_1=master.merge(mst_aliases,left_on="Matches",right_on="Skill Name",how='left')

master_1=master_1.drop_duplicates()

#analyse the matches received
master_1.to_csv(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\Master_1.csv")

#MSLT
from datetime import date

mslt=pd.read_excel(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\MST October\MST November\NHS Team Master Skill Lookup_Oct2019.xlsx",sheet_name='Master Skill Lookup Table')

# Add tab after analysing the results from the fuzzy lookup

#Remove NHS  ids
mslt=mslt[mslt.ID!='NHS00157']
mslt=mslt[mslt.ID!='NHS03659']




#Decommission existing ids
mslt.at[mslt["ID"]==80013990,"Action Type"]='Decommisioned'
mslt.at[mslt["ID"]==80013990,["De-duplicated active names"]]='Do not consider'
mslt.at[mslt["ID"]==80013990,["Updates Date"]]=str(date.today())
mslt.at[mslt["ID"]==80013990,["Update Comments"]]='Decommisioned as per EOM report for Aug'

mslt.to_csv(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\test.csv")

#print(mslt.loc[mslt["ID"]=='80011377'])
#Add all the new skills
for index,row in mesu.iterrows():
    mslt=mslt.append({"ID":row["Skill SAP ID"],"Active Name":row["Skill Name "],
                 "Skill Definition":row["Skill Definition"],"Release Date":row["Release Date"],
                 "De-duplicated active names":"Consider","Action Type":"Add","Updates Date":date.today(),
                 "Update Comments":"GSL Add for Nov 2019","Skill Type":"Skill","Level":5},ignore_index=True)


#Add alternate name for any NHS to GSL ID conversion
mslt.at[mslt["Active Name"]=="Prometheus Event Monitoring System","Alternate Old Names"]="prometheus (monitoring system)"
mslt.at[mslt["Active Name"]=="Prometheus Event Monitoring System","Alternate NHS or GSL Skill Id"]="NHS03758"


#mslt.at[mslt["Active Name"]=="Apache Mesos","Alternate Old Names"]="prometheus (monitoring system)"
mslt.at[mslt["Active Name"]=="","Alternate NHS or GSL Skill Id"]="NHS00076"



#mslt.at[mslt["Active Name"]=="Prometheus Event Monitoring System","Alternate Old Names"]="prometheus (monitoring system)"
mslt.at[mslt["Active Name"]=="Bitbucket","Alternate NHS or GSL Skill Id"]="NHS00157"
mslt.at[mslt["Active Name"]=="WireMock","Alternate NHS or GSL Skill Id"]="NHS03659"



##Add newly detected skills for the quarter
new_skills=pd.read_excel(r"C:\Users\p.d.ramakrishnan\Documents\MST for march\Skill_detected_FY19Q2_v3.xlsx")


# find the largest NHS id
nhs_ids=mslt[["ID"]]
nhs_ids["is_nhs"]=""
def check_for_nhs_id(name):
    name=str(name)
    if(name.startswith('NHS')):
        return("Yes")
    else:
        return("No")

nhs_ids["is_nhs"]=nhs_ids["ID"].apply(check_for_nhs_id)
nhs_ids=nhs_ids[nhs_ids["is_nhs"]=="Yes"]
largest_id=nhs_ids["ID"].max()   


for index,row in new_skills.iterrows():
    largest_id=int(largest_id[3:])+1
    largest_id="NHS0"+str(largest_id)
    mslt=mslt.append({"ID":largest_id,"Active Name":row["Skill name"],"Skill Definition":row["Comment"],
    "De-duplicated active names":"Consider","Updates Date":date.today(),
    "Update Comments":"Newly detected skills for FY19 Q2"},ignore_index=True)
    

mslt.to_csv(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\mslt_test.csv")

###########


# Modify tab


mesu_modify=pd.read_excel(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\MST October\MST November\Month End Skill Updates October 2019.xlsx",sheet_name='Modify',header=1)

mesu_modify_nameanddef=mesu_modify.loc[mesu_modify["Action"]=="Modify Name & Definition"]

mesu_modify_name=mesu_modify.loc[mesu_modify["Action"]=="Modify Name"]

mesu_modify_definition=mesu_modify.loc[mesu_modify["Action"]=="Modify Definition"]

# Modify Definition
for index,row in mesu_modify_definition.iterrows():
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Skill Definition"]=row["Skill New Definition"]
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Updates Date"]=date.today()
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Update Comments"]="Modify Definition as per EOM report"
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Action Type"]="Modify"

#Modify Name
for index,row in mesu_modify_name.iterrows():
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Active Name"]=row["Skill  New Value Name"]
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Updates Date"]=date.today()
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Update Comments"]="Modify Name as per EOM report"
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Action Type"]="Modify"
           
    
    if(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 2"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 2"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 3"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 3"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 4"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 4"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 5"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 5"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 6"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 6"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 12"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 12"]=row["Skill Value Name"]
    else:
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 19"]=row["Skill Value Name"]
    
        
#Modify name and definition
for index,row in mesu_modify_nameanddef.iterrows():
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Active Name"]=row["Skill  New Value Name"]
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Skill Definition"]=row["Skill New Definition"]
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Updates Date"]=date.today()
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Update Comments"]="Modify Name and Definition as per EOM report"
    mslt.at[mslt["ID"]==row["Skill SAP ID"],"Action Type"]="Modify"
    
    if(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 2"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 2"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 3"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 3"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 4"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 4"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 5"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 5"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 6"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 6"]=row["Skill Value Name"]
    elif(str((mslt.loc[mslt["ID"]==row["Skill SAP ID"]]["Alternate Old Names 12"]).values)=='[nan]'):
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 12"]=row["Skill Value Name"]
    else:
        mslt.at[mslt["ID"]==row["Skill SAP ID"],"Alternate Old Names 19"]=row["Skill Value Name"]
        
        
#checks
test=mslt.loc[mslt['Updates Date']==date.today()]



#Retire
mesu_retire=pd.read_excel(r"C:\Users\p.d.ramakrishnan\Documents\MST_August\MST October\MST November\Month End Skill Updates October 2019.xlsx",sheet_name='Retire',header=1)

for index,row in mesu_retire.iterrows():
    mslt.at[mslt["ID"]==row["Value SAP ID"],"Updates Date"]=date.today()
    mslt.at[mslt["ID"]==row["Value SAP ID"],"Update Comments"]="Retire as per EOM report for Oct"
    mslt.at[mslt["ID"]==row["Value SAP ID"],"Action Type"]="Retire"

#checks
test=mslt.loc[(mslt['Updates Date']==date.today()) & (mslt['Action Type']=='Retire')]
###Alias tab
mst_alias_tab1=mslt.loc[:,["ID","Active Name","Action Type","Level"]]
mst_alias_tab1["Name Type"]="Active Name"

mst_alias_tab1=mst_alias_tab1.rename(columns={'Active Name':'Skill Name'})

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
 'Alternate NHS or GSL Skill Id',
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
##MST_NHS
mst_nhs=mslt.loc[mslt["De-duplicated active names"]=="Consider"]
mst_nhs=mst_nhs.loc[:,['ID',
 'Action Type',
 'Active Name',
 'BG Data Available',
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


##

#mslt.to_excel(r"C:\Users\p.d.ramakrishnan\Documents\MST for march\NHS Team Master Skill Lookup_March2019.xlsx",sheet_name="Master Skill Lookup Table",index=False)

##MST Deduplication



###


writer = pd.ExcelWriter(r'C:\Users\p.d.ramakrishnan\Documents\MST_August\MST October\MST November\NHS Team Master Skill Lookup_Nov2019.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
mslt.to_excel(writer, sheet_name='Master Skill Lookup Table',index=False)
mst_alias_final.to_excel(writer, sheet_name='MST_ALIASES',index=False)
mst_original.to_excel(writer, sheet_name='MST_Original',index=False)
mst_dl.to_excel(writer,sheet_name='MST_DL',index=False)
mst_nhs.to_excel(writer,sheet_name='MST_NHS',index=False)

writer.save()










