import smtplib
import pandas as pd
import os


docs_zoom = r"C:\Users\Abdul\Documents"
temp_file_today = docs_zoom + "\\attendace_V1\\today.csv"
home = docs_zoom + "\\attendace_V1"
#year_17 = home + "\\17\\BE EEE 2021.xlsx"
#year_18 = home + "\\18\\BE EEE 2022.xlsx"
year_19 = home + "\\19\\BE EEE 2023.xlsx"
mainn = r"C:\Users\Abdul\Documents" + "\\zoom"
demo = home + "\\demo.xlsx"




##year_17 = r"C:\Users\Abdul\Documents\attendence\17\BE EEE 2021.xlsx"
##year_18 = r"C:\Users\Abdul\Documents\attendence\18\BE EEE 2022.xlsx"
##year_19 = r"C:\Users\Abdul\Documents\attendence\19\BE EEE 2023.xlsx"


mail_pass_file = home + "\\f.txt"
your_email_file = open(os.getcwd() + "\\mail_pass.txt")


secure = your_email_file.readlines()
your_mail = secure[0].split(',')[0]
##
your_password = secure[0].split(',')[1]

# establishing connection with gmail 
server = smtplib.SMTP_SSL('smtp.gmail.com', 465) 
server.ehlo() 
server.login(your_email, your_password)

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
  
    # the message to be emailed 
    message = "Hello " + name + " your attendence percentage is " + str(cent) + "%"
    
  
    # sending the email 
    server.sendmail(your_email, [email], message) 
  
# close the smtp server 
server.close()
