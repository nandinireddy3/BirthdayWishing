import pandas as pd
import datetime
import smtplib
import os

# VIA GMAIL [------------------------------]

os.chdir(r"C:\Users\dell\Desktop\Projects\BdayWisher_Mail")
#os.mkdir("p")

# Enter Your authentication details
GMAIL_ID = ''
GMAIL_PSWD = ''


def sendEmail(to,sub,msg):
    print(f"Email to {to} sent with subject: {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)

    s.sendmail(GMAIL_ID,to,f"Subject: {sub}\n\n{msg}")
    s.quit()

"""
# via whatsapp
def sendwhatmsg(to,msg):
         pywhatkit.sendwhatmsg(to,msg,18,46)
"""

if __name__ == "__main__":
    df = pd.read_excel("data.xlsx")
    #print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    yearnow = datetime.datetime.now().strftime("%Y")
    #print(today)
    
    writeInd = []
    for index, item in df.iterrows():
        #print(index,item['Birthday'])
        bday = item['Birthday'].strftime("%d-%m")
        #print(bday)
        if(today == bday) and yearnow not in str(item['Year']):
            sendEmail(item['Email'],"Happy Birthday",item['Dialogue'])
            #sendwhatmsg(item['Number'],item['Dialogue'])
            writeInd.append(index)

    for i in writeInd:
        yr = df.loc[i,'Year']
        #print(yr)
        df.loc[i,'Year'] = str(yr) + ',' + str(yearnow)

    #print(df)
    df.to_excel("data.xlsx",index=False)


