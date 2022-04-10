"""
@author: sourpi
"""

import pandas as pd
import smtplib
import time
from email.message import EmailMessage

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()  #transport layer security  
server.login('@account' , 'password of @account') #email from which we will send the emails.

def send_email(content: str, recipient: str):
    
    msg = EmailMessage()
    msg.set_content(content, subtype="plain")
    msg['Subject'] = 'Application for This_Worskshop'
    msg['From'] = '@account'
    msg['To'] = recipient
    server.send_message(msg) 
    

path =  r"...absolute path to the excel file with people's data..."
answers = pd.read_excel(path, sheet_name = 'Name of the Excel Datasheet we want to read from')
#answers.info() 
#print(answers)


#creating a series of type pandas into new lists. 
#the specified (by name) columns of the excel will be converted into lists.
Receivers = answers['Email Address'].tolist()
Nationality = answers['Country'].tolist()
NameSurname = answers['Name and surname'].tolist()

# we will send emails according to the different data and conditions we set 
for i , Country in enumerate(Nationality):
    
    
    #send an acceptance email to those that are from Greece aka GR
   if Country == 'GR':
           acceptance_text= "Dear "  + str(NameSurname[i]) + ",\n\nCongratulations! We read your application and we are more than happy to announce you that you will participate in this Workshop ..."
           send_email(acceptance_text , Receivers[i])
           #print( Receivers[i])
           print("Acceptance mail sent")
           #print(acceptance_text)
           time.sleep(10) # Sleep for 10 seconds
                                            
    # send a rejection email to those that don't have enough qualifications to participate
    # we will set the Country in the excel as NQ(non-qualified) for those people    
   elif Country == 'NQ':
           rejection_text1 = "Dear " + str(NameSurname[i]) + ",\n\nThank you for taking the time to apply for the  workshop. Try again next year...."
           send_email(rejection_text1 , Receivers[i])
           #print( Receivers[i])
           print("NQ rejection email") 
           #print(rejection_text2)
           time.sleep(10) 
      
    # if the participant is not from Greece and doesn't have the qualifications needed for this event send a rejection mail.
   else:
           rejection_text2 =  "Dear " + str(NameSurname[i]) +   ",\n\nThank you for taking the time to apply for the QSilver Workshop!\nUnfortunately, we won't be proceeding with your application at this time. This workshop will be held in Greek and as a result, you will not be able to understand the lectures and therefore participate.\n\nIf you know the Greek language in a Proficiency Level let us know with an email reply and we will add you to the workshop! In any case, we strongly encourage you to apply for a QSilver held by QGreece or by another QCousin of QWorld in the future which will be given in English or your native language.\n\nSincerely,\nThe QGreece Team" 
           send_email(rejection_text2 , Receivers[i])
           #print( Receivers[i])
           print("Rejection mail sent")
           #print(rejection_text1)
           #print(i)
           time.sleep(10)
           
           
server.quit()
