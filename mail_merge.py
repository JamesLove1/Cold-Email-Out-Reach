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
    "Cold Email ":"string",
    "Telemarketing ":"string",
    "Follow Up Email":"string",
    })
#print(df.dtypes)
# %%
df1 = df[df["notes"] == notes]

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
for i in range(len(df1["Company Name"])):
    #df1["Company Name"].iloc[i]
    
    email = outlook.CreateItem(0)
    email.To = df1["Contact Email"].iloc[i]
    #email.CC = 'xyz@gmail.com'
    email.Subject = "Aspiratio / {client} {second} - Follow Up".format(client = df1["Contact Person First Name"].iloc[i], second= df1["Contact Person Second Name"].iloc[i])
    
    email.HTMLBody = """
    
    <style>
        .body{ 
            font-family: Calibri,Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif; font-size: 16px; font-style: normal; font-variant: normal; font-weight: 400; line-height: 13.2px; 
        } 
            
        /*h3 { font-family: Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif; font-size: 14px; font-style: normal; font-variant: normal; font-weight: 700; line-height: 15.4px; 
        } 
        
        p { font-family: Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif; font-size: 14px; font-style: normal; font-variant: normal; font-weight: 400; line-height: 20px; 
        }
        
        blockquote { font-family: Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif; font-size: 21px; font-style: normal; font-variant: normal; font-weight: 400; line-height: 30px; 
        } 
        
        pre { font-family: Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif; font-size: 13px; font-style: normal; font-variant: normal; font-weight: 400; line-height: 18.5714px; 
        }*/
    </style>

    <div class="body">
    Hey """+df1["Contact Person First Name"].iloc[i]+""",<br>
    <br>
    I love how you’re changing lives with """+df1["Company Name"].iloc[i]+""" - your innovative business model is awesome!<br>
    <br>
    Being in the web3 and cryptocurrency market sure has some great perks, but there’s a great threat you might be ignoring.<br>
    <br>
    And it can be doom and gloom for your business.<br>
    <br>
    I’d love to solve that problem for you. I’m James from Aspiratio, and we specialize in connecting businesses like yours with trustworthy and skilled insurers.<br>
    <br>
    See, the market you’re working in is highly volatile… but insurance keeps you safe from malpractice and nasty lawsuits.<br>
    <br>
    So if you’re interested in a tailor-made insurance solution that fits the needs of your business, let’s put a meeting in the diary.<br>
    <br>
    <a href="https://calendly.com/aspiratio/30min">⇒ Please see a link to my here!</a>\n
    <br>
    Remember, there’s nothing at risk. We’ve helped dozens of businesses with web3 & crypto insurance… and we can do the same for you!<br>
    <br>
    Looking forward to hearing from you,<br>
    James Love<br>
    <br>
    James Love <br>
    Managing Director Aspiratio<br>
    https://aspiratio.uk/<br>
    <a href="https://calendly.com/aspiratio/30min">Book a Meeting link</a><br>
    </div>
    """

    #email.Attachments.Add()
    #email.Display()
    #email.Send()


# %%
