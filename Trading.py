# import necessary libraries

import webbrowser
import pretty_html_table
from pretty_html_table import build_table
import pandas as pd
import smtplib
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from nsetools import Nse
nse = Nse()
now = datetime.now()
# source: https://docs.python.org/2/library/datetime.html#strftime-and-strptime-behavior
dt_string = now.strftime("%a %d %b %Y %I:%M %p")
# Example is Sun 21 Mar 2021 09:51 AM

# Import CSV into dataframe
filepath = "C:/temp/Stock/stocklist.csv"
# df = pd.read_csv(filepath,encoding='windows-1252')
df = pd.read_csv(filepath)

# drop unneeded column from dataframe
# source: Stackoverflow, Unable to drop a column from pandas dataframe [duplicate]
df.drop(['Number'], axis=1, inplace=True)
print("Done-drop unneeded column from dataframe")

# convert list of stock names from a dataframe into a list
stocks = df["Name"].tolist()
print("Done-convert list of stock names from a dataframe into a list")

# get necessary metric from nse library
# source=https://stackoverflow.com/questions/62300845/how-to-get-multiple-stock-quote-with-symbol-in-python
lastPrice = {stock: nse.get_quote(stock)['lastPrice'] for stock in stocks}
companyName = {stock: nse.get_quote(stock)['companyName'] for stock in stocks}
high52 = {stock: nse.get_quote(stock)['high52'] for stock in stocks}
low52 = {stock: nse.get_quote(stock)['low52'] for stock in stocks}
print("Done- get necessary metric from nse library")

# Store in dataframes
lp = pd.DataFrame(list(lastPrice.items()), columns=['Name', 'listPrice'])
cname = pd.DataFrame(list(companyName.items()), columns=['Name', 'companyName'])
h52 = pd.DataFrame(list(high52.items()), columns=['Name', 'high52'])
l52 = pd.DataFrame(list(low52.items()), columns=['Name', 'low52'])
print("Done-Store in dataframes")

# Merge dataframes on primary key column, Name to create 1 final dataframe
a = pd.merge(cname, lp, on='Name')
b = pd.merge(a, l52, on='Name')
c = pd.merge(b, h52, on='Name')
print("Done- Merge dataframes on primary key column, Name to create 1 final dataframe")

# Merge input file and merged dataframe to create a final dataframe
final = pd.merge(c, df, on='Name')
print("Done- Merge input file and merged dataframe to create a final dataframe")

# Create bins based on list price
# source: https://pbpython.com/pandas-qcut-cut.html
final['Bin'] = pd.cut(final['listPrice'], bins=[0, 300, 500, 1000, 2000, 5000, 1000000], include_lowest=True,
                      labels=['Under 300', '300-500', '500-1000', '1000-2000', '2000-5000', '5000+'])

# perform calculations for change in stock price.
final['Perc'] = round(((final['listPrice'] - final['BuyingPrice'])/final['BuyingPrice'])*100, 2)
final['AbsDelta'] = final['listPrice'] - final['BuyingPrice']

# filter by stocks which are declining
final_buy = final[final['Perc'] < 0]
final_buy_Under300 = final_buy[final_buy['Bin'] == 'Under 300']
final_buy_300to500 = final_buy[final_buy['Bin'] == '300-500']

# sort dataframe based on perc column, sort in descending order
final_buy_Under300 = final_buy_Under300.sort_values(by=['AbsDelta', 'Perc'], ascending=True)
final_buy_300to500 = final_buy_300to500.sort_values(by=['AbsDelta', 'Perc'], ascending=True)

# The mail addresses and password
# source: https://www.tutorialspoint.com/send-mail-from-your-gmail-account-using-python
sender_address = 'anoop.mahajan78@gmail.com'
sender_pass = 'QnKeRN7818'

# send to multiple recepients
# source: https://stackoverflow.com/questions/8856117/how-to-send-email-to-multiple-recipients-using-python-smtplib
receiver_address = 'anoop.mahajan@gmail.com, patilrupali@gmail.com, ramchandra.n.mahajan@gmail.com'

# Setup the MIME
message = MIMEMultipart()
message['From'] = sender_address
message['To'] = receiver_address
# The subject line
message['Subject'] = 'Stock Buy Alert at : ' + dt_string

# The body and the attachments for the mail
# source, sending html: https://stackoverflow.com/questions/50564407/pandas-send-email-containing-dataframe-as-a-visual-table
# source, sending multiple tables: https://stackoverflow.com/questions/56208515/python-how-to-send-multiple-dataframe-as-html-tables-in-the-email-body

html = """\
<html>
      <head></head>
      <body>
        <p>Hi!<br>
           This is list of stocks with list price less than 300 Rs.:<br>
           {0}
           <br>This is list of stocks with list price between 300 to 500:<br>
           {1}

           Regards,
        </p>
      </body>
    </html>

""".format(final_buy_Under300.to_html(), final_buy_300to500.to_html())

part1 = MIMEText(html, 'html')
message.attach(part1)

# Create SMTP session for sending the mail
# use gmail with port
session = smtplib.SMTP('smtp.gmail.com', 587)
# enable security
session.starttls()
# login with mail_id and password
session.login(sender_address, sender_pass)
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()
print('Done-Mail Sent')

# Extract data into a csv
final_buy.to_csv("C:/temp/Stock/output.csv")

# open csv for validation
webbrowser.open('output.csv')


#useful links
#how to create a EC2 instance on AWS: https://praneeth-kandula.medium.com/launching-and-connecting-to-an-aws-ec2-instance-6678f660bbe6
#scp -i C:/temp/Stock/Trading.pem C:/temp/Stock/Trading.py ec2-user@ec2-54-234-82-254.compute-1.amazonaws.com:/home/ec2-user
