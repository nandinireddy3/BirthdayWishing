import pandas as pd
import datetime
import pywhatkit

# via whatsapp
def sendwhatmsg(to,msg):
    #print(str("+91")+to)
    #print(type(str("+91")+to))
    pywhatkit.sendwhatmsg(str("+91")+to,msg,0,20)


#pywhatkit.sendwhatmsg("+917995968782","GOOD EVENING",18,37)

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
            sendwhatmsg(str(item['Number']),item['Dialogue'])
            writeInd.append(index)

    for i in writeInd:
        yr = df.loc[i,'Year']
        #print(yr)
        df.loc[i,'Year'] = str(yr) + ',' + str(yearnow)

    #print(df)
    df.to_excel("data.xlsx",index=False)
