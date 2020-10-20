import tkinter 
from tkinter import messagebox
import tkinter
import pathlib
from tkinter import messagebox
from tkinter import *
from PIL import Image
from datetime import date
from PIL import ImageTk
from tkinter import ttk

import random
from colormap import rgb2hex



hex_number = rgb2hex(random.randint(0,255),random.randint(0,255),random.randint(0,255))
hex_number2 = rgb2hex(random.randint(0,255),random.randint(0,255),random.randint(0,255))
hex_number3 = rgb2hex(random.randint(0,255),random.randint(0,255),random.randint(0,255))
hex_number4 = rgb2hex(random.randint(0,255),random.randint(0,255),random.randint(0,255))

today = date.today()
#dd/mm/YY
d12 = today.strftime("%d-%m-%Y")

window=tkinter.Tk()
window.title("Online Attendance Management")


current_fileed = str(pathlib.Path(__file__).parent.absolute())


window.geometry('800x500')
bgImage=PhotoImage(file=current_fileed +"\\Background.png")



Label(window,image=bgImage).place(relwidth=1,relheight=1)


window.resizable(height=FALSE,width=FALSE)
window.iconbitmap(current_fileed + "\\logo.ico")


def recordattendance():
    
    import zumba
   # label2=tkinter.Label(window,text=hello).pack() 
    
    
    messagebox.showinfo("Record Success!","Record successful")
   
   
def mail2():
    
   # label2=tkinter.Label(window,text=hello).pack() 
    
    import smtplib
    import pandas as pd
    import os

    #print(os.getcwd())

    docs_zoom = os.getcwd().split("\\attendance_V1")[0]
    
    temp_file_today = docs_zoom + "\\attendace_V1\\today.csv"
    home = docs_zoom + "\\attendace_V1"
    #year_17 = home + "\\17\\BE EEE 2021.xlsx"
    
    #year_18 = home + "\\18\\BE EEE 2022.xlsx"
    #year_19 = home + "\\19\\BE EEE 2023.xlsx"
    mainn = r"C:\Users\Abdul\Documents" + "\\zoom"
    demo = home + "\\demo.xlsx"


    year_19 = current_file = str(pathlib.Path(__file__).parent.absolute()) + "\\19\\BE EEE 2023.xlsx"

    ##year_17 = r"C:\Users\Abdul\Documents\attendence\17\BE EEE 2021.xlsx"
    ##year_18 = r"C:\Users\Abdul\Documents\attendence\18\BE EEE 2022.xlsx"
    ##year_19 = r"C:\Users\Abdul\Documents\attendence\19\BE EEE 2023.xlsx"



    # change these as per use   
    #mail_pass_file = home + "\\f.txt"
    #your_email_file = open(os.getcwd() + "\\mail_pass.txt")


    #secure = your_email_file.readlines()
    
    your_email = "eeduceouams@gmail.com"
    ##
    your_password = "uceou.edu"

    # establishing connection with gmail 
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465) 
    server.ehlo() 
    server.login(your_email, your_password)
    #print(os.getcwd())
    # reading the spreadsheet 
    email_list = pd.read_excel(year_19) 
      
    # getting the names and the emails 
    names = email_list['First Name'] 
    emails = email_list['Email Address']
    percentage = email_list['percentage']
      
    # iterate through the records 
    for i in range(len(emails)): 
      
        # for every record get the name and the email addresses 
        name = names[i] 
        email = emails[i]
        cent = percentage[i]
        print("now of mails sent " + str(i))
      
        # the message to be emailed 
        message = """\
Subject: Attendance Report

Dear """ + name + " your cumulative attendence percentage till " + str(d12)+" is " + str(cent) + "%"
      
        # sending the email 
        server.sendmail(your_email, [email], message) 
      
    # close the smtp server 
    server.close()
    messagebox.showinfo("Mail Success!","Mail to 2nd years Successful")  
   

 
def mail3():
    import smtplib
    import pandas as pd
    import os

    #print(os.getcwd())

    docs_zoom = os.getcwd().split("\\attendance_V1")[0]
    
    temp_file_today = docs_zoom + "\\attendace_V1\\today.csv"
    home = docs_zoom + "\\attendace_V1"
    #year_17 = home + "\\17\\BE EEE 2021.xlsx"
    
    #year_18 = home + "\\18\\BE EEE 2022.xlsx"
    #year_19 = home + "\\19\\BE EEE 2023.xlsx"
    #mainn = r"C:\Users\Abdul\Documents" + "\\zoom"
    #demo = home + "\\demo.xlsx"



    year_18 = current_file = str(pathlib.Path(__file__).parent.absolute()) + "\\18\\BE EEE 2022.xlsx"



    
    your_email = "eeduceouams@gmail.com"
    ##
    your_password = "uceou.edu"

    # establishing connection with gmail 
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465) 
    server.ehlo() 
    server.login(your_email, your_password)
    print(os.getcwd())
    # reading the spreadsheet 
    email_list = pd.read_excel(year_18) 
      
    # getting the names and the emails 
    names = email_list['First Name'] 
    emails = email_list['Email Address']
    percentage = email_list['percentage']
      
    # iterate through the records 
    for i in range(len(emails)): 
      
        # for every record get the name and the email addresses 
        name = names[i] 
        email = emails[i]
        cent = percentage[i]
        print("now of mails sent " + str(i))
      
        # the message to be emailed 
        #message = "Hello " + name + " your attendence percentage is " + str(cent) + "%"
        message = """\
Subject: Attendance Report

Dear """ + name + " your cumulative attendence percentage till " + str(d12)+" is " + str(cent) + "%"
      
        # sending the email 
        server.sendmail(your_email, [email], message) 
      
    # close the smtp server 
    server.close()
    
    
    
    
    
    messagebox.showinfo("Mail Success!","Mail to 3rd years Successful")    
    
def mail4():
    
   # label2=tkinter.Label(window,text=hello).pack()
    import smtplib
    import pandas as pd
    import os

    print(os.getcwd())

    docs_zoom = os.getcwd().split("\\attendance_V1")[0]
    
    temp_file_today = docs_zoom + "\\attendace_V1\\today.csv"
    home = docs_zoom + "\\attendace_V1"
    year_17 = current_file = str(pathlib.Path(__file__).parent.absolute()) + "\\17\\BE EEE 2021.xlsx"
    
    #year_18 = home + "\\18\\BE EEE 2022.xlsx"
    #year_19 = home + "\\19\\BE EEE 2023.xlsx"
    #mainn = r"C:\Users\Abdul\Documents" + "\\zoom"
    #demo = home + "\\demo.xlsx"






    
   
    your_email = "eeduceouams@gmail.com"
    ##
    your_password = "uceou.edu"
    
    # establishing connection with gmail 
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465) 
    server.ehlo() 
    server.login(your_email, your_password)
    #print(year_17)
    # reading the spreadsheet 
    email_list = pd.read_excel(year_17) 
      
    # getting the names and the emails 
    names = email_list['First Name'] 
    emails = email_list['Email Address']
    percentage = email_list['percentage']
      
    # iterate through the records 
    for i in range(len(emails)): 
      
        # for every record get the name and the email addresses 
        name = names[i] 
        email = emails[i]
        cent = percentage[i]
        print("now of mails sent " + str(i))
      
        # the message to be emailed
        message = """\
Subject: Attendance Report

Dear """ + name + " your cumulative attendence percentage till " + str(d12)+" is " + str(cent) + "%"
        #message = "Hello " + name + " your attendence percentage is " + str(cent) + "%"
        
      
        # sending the email 
        server.sendmail(your_email, [email], message) 
      
    # close the smtp server 
    server.close()
    
    
    messagebox.showinfo("Mail Success!","Mail to 4th years Successful")




bt=tkinter.Button(window,text="Record Attendance",bg=hex_number,fg="white",command=recordattendance)
bt.pack()
bt.place(height=40, width=150,relx=0.75,rely=0.28+0.05) 

bt=tkinter.Button(window,text="Mail 2nd Year Attendance",bg=hex_number2,fg="white",command=mail2)
bt.pack()
bt.place(height=40, width=150,relx=0.75,rely=0.39+0.05)

bt=tkinter.Button(window,text="Mail 3rd Year Attendance",bg=hex_number3,fg="white",command=mail3)
bt.pack()
bt.place(height=40, width=150,relx=0.75,rely=0.50+0.05)

bt=tkinter.Button(window,text="Mail 4th Year Attendance",bg=hex_number4,fg="white",command=mail4)
bt.pack()
bt.place(height=40, width=150,relx=0.75,rely=0.61+0.05)


    
label=tkinter.Label(window,text="Record attendance for\n \n " ,bg="white",fg="black",font=("Arial Bold",12))
label.pack()
label.place(relx=0.01+0.05,rely=0.33)

labe=tkinter.Label(window,text=str(d12) ,bg="white",fg="black",font=("Arial Bold Italic",24))
labe.pack()
labe.place(relx = 0.0588,rely = 0.4)

window.mainloop()
