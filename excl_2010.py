import pandas as pd
from pandas import ExcelWriter
import xlrd
import time
start_time = time.time()

list_variable_result=[]
list_variable_colonne=[]
wb = xlrd.open_workbook('C:\FichierSource\REACTION QUAD 2014.xls')
print('sheets', wb.sheet_names())# nom des feuilles
sheets= wb.sheet_names()
sh = wb.sheet_by_name(sheets[0])#choisir le nom du feuille à triater
#print('sh=',sh)
#print('sh.nrows=',sh.nrows)
#listeColonne=sh.row_values(0)
for sheet in sheets:
    resultPanda=[]
    listeColonne=[]
    print('sheet=',sheet)
    sh = wb.sheet_by_name(sheet)
    for rownum in range(sh.nrows):
        #print('line=',sh.row_values(rownum))
        if rownum == 0:
           listeColonne=sh.row_values(rownum)
        else:
            resultPanda.append(sh.row_values(rownum))
    if not resultPanda:
        resultPanda.append(['']* len(listeColonne))
        print('resultPanda=',resultPanda)
    #print('resultPanda=',resultPanda)
    print('listeColonne=',listeColonne)
    exec('result_'+sheet+'='+str(resultPanda))
    list_variable_result.append('result_'+sheet)

    exec('listColonne_'+sheet+'='+str(listeColonne))
    list_variable_colonne.append('listColonne_'+sheet)
    #list_variable.append('df_'+sheet)
    
    print('############################################################')
print('list_variable_result=',list_variable_result)
print('list_variable_colonne=',list_variable_colonne)

for i, j in  zip (list_variable_result,list_variable_colonne ):
    print(i)
    print(j)
if len(list_variable_result) != len(list_variable_colonne):
    print('Erreur')
elif len(list_variable_result)==7:
    df0=pd.DataFrame(data=eval(list_variable_result[0]),columns=eval(list_variable_colonne[0]))

    df1=pd.DataFrame(data=eval(list_variable_result[1]),columns=eval(list_variable_colonne[1]))
    
    df2=pd.DataFrame(data=eval(list_variable_result[2]),columns=eval(list_variable_colonne[2]))
    df3=pd.DataFrame(data=eval(list_variable_result[3]),columns=eval(list_variable_colonne[3]))
    df4=pd.DataFrame(data=eval(list_variable_result[4]),columns=eval(list_variable_colonne[4]))
    df5=pd.DataFrame(data=eval(list_variable_result[5]),columns=eval(list_variable_colonne[5]))
    df6=pd.DataFrame(data=eval(list_variable_result[6]),columns=eval(list_variable_colonne[6]))
    
    writer = pd.ExcelWriter('output.xlsx')

    df0.to_excel(writer,sheets[0],index=False)
    df1.to_excel(writer,sheets[1],index=False)
    df2.to_excel(writer,sheets[2],index=False)
    df3.to_excel(writer,sheets[3],index=False)
    df4.to_excel(writer,sheets[4],index=False)
    df5.to_excel(writer,sheets[5],index=False)
    df6.to_excel(writer,sheets[6],index=False)
    
writer.save()

'''        
colonne1 = sh.col_values(3)#recupérer une colonne
print('resultPanda=',resultPanda)
print('colonne1=',colonne1)
print('listeColonne=',listeColonne)
df=pd.DataFrame(data= resultPanda,columns=listeColonne)
df.to_csv('result_csv.csv')
'''
'''
col=[['a', '20a'],['lo52','3625'],['fr2','ik8']]
col=[['', '']]
    
df1 = pd.DataFrame(data=col,columns=['a','b'])
df2 = pd.DataFrame(data=[['a', '20a'],['lo52','3625'],['fr2','ik8']],columns=['a','b'])
print('df2=',df2)
writer = pd.ExcelWriter('output.xlsx')

df1.to_excel(writer,'sheet1')
df2.to_excel(writer,'sheet2')
writer.save()
'''
etime = time.time() - start_time
print('Time=',etime)

