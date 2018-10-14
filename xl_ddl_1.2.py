import os
from xlrd import *
import  itertools
from easygui import *
import sys
import threading
import time





def check_none(var):
    if var is None:
        sys.exit(0)
    else:
        pass


ret_val = msgbox("Welcome to Postgres & Phoenix DDL generator","【D】【D】【L】  Generator ☺",ok_button="START")
if ret_val is None: 
    sys.exit(0)

m=msgbox("-Make sure Data Model Template matches with the input file.\n\n-In case of commenting on sheets,do comment after an empty row. \n\n-Make sure there is no empty sheets.\n\n-'Column Name','Data Type','Allow null','Primary Key','Foreign key','Referenced Table' are mandatory columns in same order.\n\n-'Unique Key' column  and Hbase DM sheets are optional.\n\n-Phoenix data model sheet should contain 'Row Key' column.","Prerequisite",ok_button="PROCEED")
check_none(msgbox)


sch=enterbox("Enter the name of the schema","Schema Name")
check_none(sch)
xl_path=fileopenbox(msg="Select the Data Model to be Converted",title="Data Model",filetypes='*.xls')
check_none(xl_path)
op_path_1=fileopenbox(msg="Select the Output file for generating Postgres DDL",title="Postgres SQL file ",filetypes= '*.sql')
check_none(op_path_1)
op_path_2=fileopenbox(msg="Select the Output file for generating Phoenix DDL",title="Phoenix SQL file ",filetypes= '*.sql')
check_none(op_path_2)


def value_from_key(sheet, key):
    for row_index, col_index in itertools.product(range(sheet.nrows), range(sheet.ncols)):
        if sheet.cell(row_index, col_index).value == key:
            return (row_index, col_index)
        else:
            no_val = 1
            return no_val

book = open_workbook(xl_path)
hb=[]
n=[]
y=[]
pk=[]
uk=[]
z=[]
m=[]
rk=[]
fk=0
rk=[]
fk_str = []
uk_list = []
    # print number of sheets
#print  (book.nsheets)

    # print sheet names
#print (book.sheet_names())
b_s = book.sheet_names()

for i in range (0,book.nsheets):
    z.append(i)
#print(z)

for i in range (0,book.nsheets):
    c =  str(b_s[i]) 
    if (c.lower() == 'index' or c.lower() == 'version history' or c.lower() == 'version_history'):
        n.append(i)



if(len(n) != 0):
	y = list(set(z) - set(n))
else:
	y = z 

    # get the first worksheet
fo=open(op_path_1,'w')

fo.write("------------------------------------------------------------------------------------------------------------------------------------------------\n")
fo.write ("-----------------------------------------------------------DDL scripts-------------------------------------------------------------------------\n")
fo.write("--------------------------------------------------------------------------------------------------------------------------------------------\n\n\n")

fo.close()

for i in range(0,book.nsheets):
    sheet = book.sheet_by_index(i)
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            a = str(sheet.cell_value(row,col))
            if  a.lower() == 'row key':
                hb.append(i)
#print(hb)





if(len(hb) !=0 ):
    m = list(set(y) - set(hb))
else:
    m = y   
for i in range(0,len(m)):
        fo=open(op_path_1,'a')  
        sheet = book.sheet_by_index(m[i])
        c = tuple(sheet.col_values(0))
        #print(m[i])
        if('' not in c):
            d=len(c)
            
        else:
            d=c.index('')
        #print(d)
        if ('Column Name' not in c):
            fo.write("------------------------------------------------------Check sheet "+str(b_s[m[i]])+": 'Column Name' cell is missing----------------------------")
            continue
        else:
            position = c.index('Column Name')
        heading = sheet.row_values(position)
        if('Unique Key' in heading):
            ukval = heading.index('Unique Key')
            #print ("uk column in  sheet" + str(m[i]) +" is "  ,ukval)
            uk_list.append(m[i])
        #print(position)
        x = value_from_key(sheet, 'Table Name')
        if x == 1:
            fo.write("------------------------------------------------------Check sheet "+str(b_s[m[i]])+": 'Table Name' cell is missing----------------------------")
            continue
        else:
            table_name_value = list(value_from_key(sheet, 'Table Name'))
            table_name = sheet.cell(table_name_value[0],table_name_value[1]+1).value
        #print("Table name is :",table_name)
        fo.write("\n\n-----------------------------------------------------"+str(table_name)+"--------------------------------------------------------------\n\n")
        fo.write("CREATE TABLE " +str(sch).upper()+"."+str(table_name) +"\n ( \n")       
        value = value_from_key(sheet, 'Column Name')
        for j in range(position+1,d):
            a = sheet.row_values(j)
            b = tuple(a)
            #print(b)
            if(b[2] == 'N'):
                nullable = 'NOT NULL'
            else:
                nullable = ''

            fo.write(str(b[0]) + "  " +  str(b[1]) + "  " + str(nullable) +"," + "\n")
            
            if(b[3] == 'Y'):
                pk.append(b[0])
                
            if (m[i] in uk_list):
                if(b[ukval] != ''):
                    uk.append(b[0])
                

            if(b[4] == 'Y'):
                if(len(fk_str) == 0):
                    fk = 1
                else:
                    fk = len(fk_str) + 1
                fk_str.append("CONSTRAINT "+str(table_name)+"_FK"+str(fk) +" FOREIGN KEY (" +str(b[0])+") REFERENCES " + str(b[5]) + "(" +str(b[0])+")")
                
        #print ("string l",len(fk_str))
        if(len(pk) != 0 ):
            if(len(pk) > 1):
                s = str(tuple(pk))
                t = s.replace("'","")
                fo.write("CONSTRAINT "+str(table_name)+"_PK  PRIMARY KEY " + t + ",\n")
                #fo.write("\n"  + "); \n \n ")
            else:
                fo.write("CONSTRAINT "+str(table_name)+"_PK  PRIMARY KEY  (" + str(pk[0]) + "),\n")
                #fo.write("\n"  + "); \n \n")
        


            
        if(len(fk_str) > 1):
            for i in range(0,len(fk_str)-1):
                fo.write(fk_str[i] + ",\n")
            fo.write(fk_str[len(fk_str) - 1] +",\n" )
        elif(len(fk_str) == 1):
            fo.write(fk_str[0] + ",\n")

     
        
        if(len(uk) > 1):
            w = str(tuple(uk))
            x = w.replace("'","")
            fo.write("CONSTRAINT "+str(table_name)+"_UK  UNIQUE " + x + ",\n")
        elif(len(uk) == 1):
            fo.write("CONSTRAINT "+str(table_name)+"_UK  UNIQUE  (" + str(uk[0]) + "),\n")
        fo.close()

        with open(op_path_1,'rb+') as filehandle:
            filehandle.seek(-3, os.SEEK_END)
            filehandle.truncate()
        filehandle.close()

        fo=open(op_path_1,'a')
        fo.write(");\n\n")
        fo.close()
        
        del uk[0:]    
        del pk[0:]
        del fk_str[0:]

fo=open(op_path_1,'a')
fo.write("\n\n\n------------------------------------------------------------------------------------------------------------------------------------------------\n")
fo.write ("----------------------------------------------------------END----------------------------------------------------------------------------------------\n")
fo.write("--------------------------------------------------------------------------------------------------------------------------------------------------\n\n\n")
fo.close()



if(len(hb) != 0):
    fo=open(op_path_2,'w')   
    fo.write("---------------------------------------------------------------------------------------------------------------------------------------------\n")
    fo.write ("----------------------------------------------------PHOENIX DDL scripts---------------------------------------------------------------------\n")
    fo.write("-----------------------------------------------------------------------------------------------------------------------------------------\n\n\n")
    fo.close()
    for i in range(0,len(hb)):
        fo=open(op_path_2,'a')   
        sheet = book.sheet_by_index(hb[i])
        c = tuple(sheet.col_values(0))
        #print(hb[i])
        if('' not in c):
            d=len(c)
            
        else:
            d=c.index('')
        #print(d)
        if ('Column Name' not in c):
            fo.write("------------------------------------------------------Check sheet "+str(b_s[hb[i]])+": 'Column Name' cell is missing----------------------------")
            continue
        else:
            position = c.index('Column Name')

        if ('Row Key' not in list(sheet.row_values(position))):
            fo.write("------------------------------------------------------Check sheet "+str(b_s[hb[i]])+": 'Row Key' cell is missing----------------------------")
            continue
        else:
            rk_position = list(sheet.row_values(position)).index('Row Key')

        heading = sheet.row_values(position)
        x = value_from_key(sheet, 'Table Name')
        if x == 1:
            fo.write("------------------------------------------------------Check sheet "+str(b_s[hb[i]])+": 'Table Name' cell is missing----------------------------")
            continue
        else:
            table_name_value = list(value_from_key(sheet, 'Table Name'))
            table_name = sheet.cell(table_name_value[0],table_name_value[1]+1).value
        #print("Table name is :",table_name)
        fo.write("\n\n--------------------------------------------------"+str(table_name)+"-------------------------------------------------------------\n\n")
        fo.write("CREATE TABLE " +str(sch).upper()+"."+str(table_name).upper() +"\n ( \n")    
        fo.write("PK VARCHAR NOT NULL,\n")
        for j in range(position+1,d):
            a = sheet.row_values(j)
            b = tuple(a)
            if(b[rk_position] == 'Y'):
                rk.append(b[0])
                


            fo.write(str(b[0]) + " VARCHAR," + "\n")
        fo.write("CONSTRAINT "+str(table_name)+"_PK  PRIMARY KEY (PK))  default_column_family='cf',column_encoded_bytes=0;\n\n")
        if(len(rk) != 0 ):
            if(len(rk) > 1):
                s = str(tuple(rk))
                t = s.replace("'","").replace(","," |").replace("(","'").replace(")","'")
                fo.write("--PK columns : "+ t + "\n\n")
                #fo.write("\n"  + "); \n \n ")
            else:
                fo.write("--PK column : "+ str(rk[0]) + "\n\n")
        del rk[0:]        
        fo.close()

    fo=open(op_path_2,'a')
    fo.write("\n\n\n------------------------------------------------------------------------------------------------------------------------------------------------\n")
    fo.write ("----------------------------------------------------------END----------------------------------------------------------------------------------------\n")
    fo.write("--------------------------------------------------------------------------------------------------------------------------------------------------\n\n\n")
    fo.close()

def animate():
    for c in itertools.cycle(['|', '/', '-', '\\']):
        if done:
            break
        sys.stdout.write('\rConverting ' + c)
        sys.stdout.flush()
        time.sleep(0.1)
        sys.stdout.write('\rFinished    ')
        
done = False
t = threading.Thread(target=animate)
t.start()
time.sleep(10)
done = True
msgbox("Please Check the output file","Thank You ☺")
msgbox("Application Developed by Seshadri of TCS Optumera Base Product DB Team",title="About",ok_button='Close')

