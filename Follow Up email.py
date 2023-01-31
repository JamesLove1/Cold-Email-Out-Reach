#%%
import pandas as pd
import win32com.client 
notes = "Prospecting v3"

#%%
df = pd.read_excel("Master_prospecting_sheet.xlsx")
df = df.astype({
    "Company Name": "string",
    "Website":"string",
    "Company Email":"string",
    "Company phone number":"string",
    "Linkedin Company URL":"string",
    "Contact Person First Name":"string",
    "Contact Person Second Name":"string",
    "Company position":"string",
    "Contact Email":"string",
    "Contact Persons Phone Number":"string",
    "Location":"string",
    "Contact Sorces URL e.g. LinkedIn":"string",
    "notes":"string",
    "cold email ":"string",
    "Telemarketing":"string",
    "Follow up email":"string",
    "Cold Email 23/11/2022":"string",
    "Telersales/Marketing 21/11/2022":"string",
    "Folllow up 01/12/2022":"string",
    "Cold Email 29/01/2023":"string",
    "Telemarketing 29/01/2023":"string",
    "Follow Up Email":"string",
    })
#print(df.dtypes)
# %%
df1 = df[df["notes"] == notes]
send_follow = [df1.iloc[16], df1.iloc[27]]
print(send_follow)



#print(df1["Website"].iloc[0])
# %%

outlook = win32com.client.Dispatch('Outlook.Application')
#email = outlook.CreateItem(0)
#email.To = 'xyz@gmail.com'
#email.CC = 'xyz@gmail.com'
#email.Subject = "testing "
#email.Body= 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'
#email.Attachments.Add()
#email.Display()

# %%
for emails in send_follow:
    #print(email["Company Name"])
    
    email = outlook.CreateItem(0)
    email.To = emails["Contact Email"]
    #email.CC = 'xyz@gmail.com'
    email.Subject = "Aspiratio / {client} {second} - Follow Up".format(client = emails["Contact Person First Name"], second= emails["Contact Person Second Name"])
    
    email.HTMLBody = """
    
    <style>
        .body{ 
            font-family: Calibri,Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif; 
            font-size: 16px; 
            font-style: normal; 
            font-variant: normal; 
            font-weight: 400; 
            line-height: 13.2px; 
        } 
    </style>

    <div class="body">
    Hey """+emails["Contact Person First Name"]+""",<br>
    <br>
    Thank you so much for taking the time to speak to my college. He mentaioned that you would be intrested in talking further.<br> 
    <br>
    Please click on the <a href="https://calendly.com/aspiratio/30min">link</a> to my calendar to book a meeting.<br>
    <br>
    Kind regards,<br>
    James Love<br>
    <br>
    James Love <br>
    Managing Director Aspiratio<br>
    <a href="https://aspiratio.uk/">https://aspiratio.uk/</a><br>
    <a href="https://calendly.com/aspiratio/30min">Book a Meeting link</a><br>
    </div>
    """

    #email.Attachments.Add()
    #email.Display()
    email.Send()


# %%
