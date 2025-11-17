import pandas as pd
from datetime import datetime, timedelta, date
import win32com.client as win32


filename_format = datetime.now()
filename = r"F:\Trade Support\Corporate Actions\CA check\CA Raw file\CA_Active_Summary_"+filename_format.strftime("%b")+"_"+filename_format.strftime("%d")+"_"+filename_format.strftime("%Y")+".csv"

df = pd.read_csv(filename)
#df = pd.read_excel(filename)
Output = df[["Security ID","Security Name","Event Type","Client Deadline Date","Response Status(ELIG)","Client"]]
Output = Output.dropna()
Output['Client Deadline Date'] = pd.to_datetime(Output['Client Deadline Date'].str[:11], format='%d %b %Y')
    
startdate = date.today()
enddate = startdate + timedelta(days=15)

FinalOutput = Output[Output['Response Status(ELIG)'].str.contains("RESPONSE REQUIRED") &
                    #(Output['Client'].str.contains("CIF") &
                    (Output['Client Deadline Date'].dt.date >= startdate) & 
                    (Output['Client Deadline Date'].dt.date < enddate) &
                    ~Output['Event Type'].isin(["OPTIONAL DIVIDEND", "CASH DISTRIBUTIONS", "DIVIDEND REINVESTMENT               "])]

FinalOutput["Comments"]=""
FinalOutput = FinalOutput.sort_values(by='Client Deadline Date')
#FinalOutput.drop(columns=FinalOutput.columns[0], axis=1,  inplace=True)

FinalOutput = FinalOutput.to_html(index=False)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.SentOnBehalfOfName = 'vvasanthakumar@ci.com'
mail.To = ''
mail.Subject = 'DAILY CA CHECK'
mail.HTMLBody = (FinalOutput)
mail.display()

# html_table = FinalOutput.to_html(index=False)

#subject = "CA to be sent"
#body = f"<p>{html_table}</p>"

#recipients = "#zhanxiong.li@ci.com"

#send_email(subject, body, recipients=recipients, display=True)
