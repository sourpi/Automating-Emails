# Automating-Emails
A python script that will pull data (such as a person's Name, Country, Email Address) from an Excel sheet and then compose and send emails based on their qualifications and other conditions. 

Τhis code can be adapted to several projects such as recruitment or newsletter mail handling

In order to login to the gmail account of the sender you need to allow access to less secure apps in your Google Settings first. 
The order of the message configuration (Subject, Recepient, Content etc.) in the send_email function and the ten seconds delay after every sent email will guarantee 
that the email won't go to the spam folder of the recipient.
