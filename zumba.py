import os
import pathlib
import pandas as pd
from datetime import date
import numpy as np




today = date.today()
#dd/mm/YY
d1 = today.strftime("%Y-%m-%d")


current_file = str(pathlib.Path(__file__).parent.absolute())

docs_zoom = current_file.split("\\attendance_V1")[0]
#print(docs_zoom)
temp_file_today = docs_zoom + "\\attendance_V1\\today.csv"
home = docs_zoom + "\\attendance_V1"
year_17 = home + "\\17\\BE EEE 2021.xlsx"
year_18 = home + "\\18\\BE EEE 2022.xlsx"
year_19 = home + "\\19\\BE EEE 2023.xlsx"
mainn = docs_zoom + "\\zoom"
os.chdir(mainn)




def do_it(paath):
    #print(paath)
    name_ =[]
    no_ = []
    file = open(paath)
    data = file.readlines()
    #print(data)
    for t in range(0,len(data)):
        msg = data[t].split(':')[-1]
        try:
            if (int(msg)> 100517000000):
                its_roll = msg
                valid_name = (data[t].split("From")[-1]).split(":")[0]
                valid_no = (data[t].split("From")[-1]).split(":")[-1]
                print(valid_no, valid_name)
                name_.append(valid_name)
                no_.append(valid_no)
                
            
          
                
        except:
            continue
    return name_,no_



    
    

cont = os.listdir()
#print(cont)
r = []
aname_ = []
ano_ = []
for x in range (0,len(cont)):
    
    r.append(cont[x].split(" "))
    date = r[x][0]
    if(d1 == str(date)):
        #print("its true", d1) 
        time = r[x][1]
        topic = r[x][2]
        meeting_id = r[x][3]
        os.chdir(mainn + "\\" + cont[x])
        sub_files = os.listdir()

        try:
            
            for y in range (0,len(sub_files)):
        
                if(sub_files[y].split('.')[-1]=="txt"):
                    chat_f = sub_files[y]
                    #print(chat_f)
                    nameee, nooo = do_it((os.getcwd()+ "\\" + chat_f))
                    #print(nameee,nooo)
                    for xy in range(0, len(nameee)):
                        aname_.append(nameee[xy])
                        ano_.append(nooo[xy])





            #print(aname_)
            #print(ano_)
            dic = { 'name' : aname_,
                    'number' : ano_
                   }
            temp_df = pd.DataFrame(dic)
              
              
            # sorting by first name 
            temp_df.sort_values("number",ascending=False, inplace = True) 
              
            # dropping ALL duplicte values 
            temp_df.drop_duplicates(subset ="name", 
                                 keep = 'first', inplace = True) 
              
            # displaying data 

            temp_df.to_csv(temp_file_today,index = False)

            #now adding the attendence to org file
            
            today_df = pd.read_csv(temp_file_today)
            #print("ok bbo")
            if(today_df["number"][0]<100518000000):
                big_df = pd.read_excel(year_17)
                main_save_loc = year_17
            elif(today_df["number"][0]<100519000000):
                big_df = pd.read_excel(year_18)
                main_save_loc = year_18
            elif(today_df["number"][0]<100520000000):
                big_df = pd.read_excel(year_19)
                main_save_loc = year_19
            
            

            if(big_df.columns[-1]!= date):
                big_df[date] = np.zeros((big_df.shape[0],1))# if date is same wont add date twice
            #print(today_df.shape)
            for x in range(0,today_df.shape[0]):
                r_no_ins = today_df["number"][x]
                indx = big_df[big_df['Roll Number']==r_no_ins ].index.values
                #print(x,r_no_ins)
                big_df.at[indx,date]=1

            big_df.to_excel(main_save_loc,index = False)

            #clc percentage from here
            dayss = big_df.shape[1] - 6
            ra = pd.read_excel(main_save_loc, usecols = range(6,big_df.shape[1]))
            ra['sum'] = ra.sum(1)
            for per in range(0, big_df.shape[0]):
                attended = ra['sum'][per]
                percentage_att = attended/dayss * 100
                big_df.at[per, "percentage"] = percentage_att
            big_df.to_excel(main_save_loc,index = False)
        except:
            continue
            
        
        
        

        




































