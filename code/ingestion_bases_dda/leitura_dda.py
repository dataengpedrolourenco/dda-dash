# -*- coding: utf-8 -*-
"""


@author: DQ623AA
"""

# -*- coding: utf-8 -*-
"""
@author: DQ623AA
"""

import pandas as pd 
import numpy as np
import datetime

path_read = 'C:\\Users\\DQ623AA\\EmployeeSchedule_-_13_weeks2021-4-20-8-27-8-134169936.xlsx'#input

path_write = 'C:\\Users\\DQ623AA\\AlocaçãoDDA_Retain.xlsx'#output
print('Ingestão OK')
cont2=0
name = []
gpn = []
rank = []
client = []
project = []
client_id = []
eng_number = []
torre = []
h1 = []
h2 = []
h3 = []
h4 = []
h5 = []
h6 = []
h7 = []
h8 = []
h9 = []
h10 = []
h11 = []
h12 = []
h13 = []
h14 = []
h15 = []
h16 = []
h17 = []
h18 = []
h19 = []
h20 = []
h21 = []
h22 = []
h23 = []
h24 = []
h25 = []
h26 = []

ckp_employee = 0
ckp_projeto = 0
ckp_torre = 1

cont = 0
data = 0
ckp_employee = 0
ckp_projeto = 0
ckp_torre = 1

df = pd.read_excel(path_read, header = None)
list = df.values.tolist()

for x in list:   
    
    if x[0] == 'Schedule Week Start Date:': 
        data = x[1]
        
    if x[0] == 'Employee Name:':
        ckp_employee = 1
        i = x[1].index('-')
        nome = x[1][:i-1]
        codigo = x[1][i+2:]
        nivel = x[7]
    
    if ckp_employee == 1 and ckp_projeto == 1 and ckp_torre == 1:          
        name.append(nome)
        gpn.append(codigo)
        rank.append(nivel)
        torre.append(tower)
    
    
    if ckp_projeto == 1:
        client.append(x[0])
        client_id.append(x[1])
        project.append(x[3])
        eng_number.append(x[4])
        h1.append(x[7])
        h2.append(x[8])
        h3.append(x[9])
        h4.append(x[10])
        h5.append(x[11])
        h6.append(x[12])
        h7.append(x[13])
        h8.append(x[14])
        h9.append(x[15])
        h10.append(x[16])
        h11.append(x[17])
        h12.append(x[18])
        h13.append(x[19])
        h14.append(x[20])
        h15.append(x[21])
        h16.append(x[22])
        h17.append(x[23])
        h18.append(x[24])
        h19.append(x[25])
        h20.append(x[26])
        h21.append(x[27])
        h22.append(x[28])
        h23.append(x[29])
        h24.append(x[30])
        h25.append(x[31])
        h26.append(x[32])
    
    if x[0] == 'Client Name':
        ckp_projeto = 1 
        
    if x[0] == 'Emp Cost Centre:':
        ckp_torre = 1
        tower = x[7]
        
    if pd.isnull(x[0:33]).all():
        nome = ''
        codigo = ''
        nivel = ''
        ckp_employee = 0
        ckp_projeto = 0
        ckp_torre = 0

list_of_lists = [h1,h2,h3,h4,h5,h6,h7,h8,h9,h10,h11,h12,h13,h14,h15,h16,h17,h18,h19,h20,h21,h22,h23,h24,h25,h26]

date_object = data

a_dict = {"Name": name, "GPN":gpn, "Department":torre, "Rank":rank, "Client":client, "Project":project,"Client ID": client_id, "Eng #":eng_number}

for i in list_of_lists:
    string_test = i
    data_test = ""+str(datetime.date.strftime(date_object + datetime.timedelta(7*cont), "%d/%m/%Y"))
    a_dict[data_test] = i
    cont = cont+1

df2 = pd.DataFrame.from_dict(a_dict, orient='index')

df2 = df2.loc[:, df2.isnull().sum() < 0.8*df2.shape[0]]
df2 = df2.transpose()
df2 = (df2.set_index(['Name','GPN','Department','Rank','Client','Project','Client ID','Eng #']).stack().reset_index(name='Hours of Week').rename(columns={'level_8':'Date'}))

df2.to_excel(path_write, index=False)

print('Gravacao concluida!')
