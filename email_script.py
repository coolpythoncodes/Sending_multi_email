#import required libraries

import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#filepath of Excel workbook that has the email
#To know how to get filepath of your excel file check: https://coolpythoncodes.com/openpyxl-python-tutorial/
file_path = "C:\\Users\\RAPTURE C. GODSON\\Documents\\openpyxl\\my_blog_email_list.xlsx"
wb = openpyxl.load_workbook(file_path)
#Get the worksheet where your email is
sheet = wb["USA"]
#create a list of all the email in the worksheet
email = [sheet.cell(row=i,column=1).value for i in range(14,25)]


host = "smtp.gmail.com"
port = 587
username = "Your gmail email"
password = "your gmail password"
s = smtplib.SMTP(host,port)
s.ehlo()
s.starttls()
s.login(username,password)

message = MIMEMultipart("alternative")
message["Subject"] = "Testing"
message["From"] = username


plain_text = """"
Hi,
How are you?
Real Python has many great tutorials:
www.realpython.com
"""
html_text ="""
<html>
  <body>
    <p>Hi,<br>
       How are you?<br>
       <a href="http://www.realpython.com">Real Python</a> 
       has many great tutorials.
    </p>
  </body>
</html>
"""
part_1 = MIMEText(plain_text,"plain")
part_2 = MIMEText(html_text,"html")
message.attach(part_1)
message.attach(part_2)
i = 0
while i < len(email):
    for name in email:
        if name==None:
            pass
        else:
            s.sendmail(username,email[i],message.as_string())  
        i+=1
        
s.quit()