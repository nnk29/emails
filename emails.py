import pandas as pd 
from email.message import EmailMessage
import re
import ssl
import smtplib
from datetime import datetime
import pytz  

# Automates sending emails to companies listed on a google sheet, using the email template below

def gettime(): 
    est = pytz.timezone('US/Eastern')

    current_time = datetime.now(est)
    hour = int(current_time.strftime("%H"))
    descriptor = "day"
    if hour >= 0 and hour < 12: 
         descriptor = "morning"
    elif hour >= 12 and hour < 18: 
         descriptor = "afternoon"
    elif hour >= 18: 
         descriptor = "evening"

    return descriptor
    
     

def send_email(business_name, email): 
    
    email_sender = 'ruraag@gmail.com'
    email_password = '' # Did not include password for security reasons
    email_receiver = email; 

    time = gettime()

    subject = 'Partner With Rutgers Asian A Capella!'
    body = """ Good """ + time + """ ,

I hope this email finds you well. My name is Naren, and I am reaching out on behalf of the Rutgers Asian A Cappella Group (RAAG). Our team, composed of talented and passionate college students, has performed and competed in singing competitions nationwide. As we gear up for an exciting competition season, we are actively seeking sponsors to support our journey.

Last year, we won 3rd place at the South Asian A Cappella Nationals (A3) and were semifinalists for the International Championship of Collegiate A Cappella (ICCA), which brought in audiences of over 40,000. These achievements have earned us recognition and admiration for our musical prowess and cultural diversity.

We believe that """ + business_name + """ would be a perfect fit as a sponsor for our upcoming season. Your establishment's dedication to providing authentic and delicious Indian cuisine through your locations and catering services resonates with our commitment to South Asian cultural expression. We see a fantastic opportunity to help grow awareness of """ + business_name + """'s culinary delights within our diverse and engaged audience. In return for your generous support, we offer various sponsorship tiers, each providing unique advertising opportunities:

**Bronze Sponsor:** Social media shoutouts, event acknowledgments, and verbal thank yous to our sponsor at all RAAG events and performances.

**Silver Sponsor:** All benefits from the Bronze package, inclusion in our intro video for competitions, a link on the RAAG social media, and discount codes for the sponsor's product shared through social media posts and stories.

**Gold Sponsor:** All Silver package benefits, event mentions, sponsor's logo on event materials, dedicated social media story and post, and inclusion in exclusive event content.

**Platinum Sponsor:** All benefits from the Gold package, exclusive promotional video, dedicated social media post, longer-form video with sponsor integration, event sponsorship opportunities, presence at RAAG events, merch collaboration, premium placement, and a complimentary performance at a sponsor event.

We are confident that this partnership will be mutually beneficial. Your establishment will be prominently featured in our promotional materials, including posters, social media campaigns, and event programs. Your support will help us pursue our passion for music and cultural expression while providing """ + business_name + """ with valuable exposure to an appreciative South Asian audience.

If you are interested in discussing this opportunity further or have any questions, please feel free to reach out to us by email at ruraag@gmail.com, or through our Instagram. We would be delighted to set up a Zoom, phone call, or in-person meeting to provide more details and explore how we can tailor this partnership to align with your marketing objectives.

Thank you for considering our proposal. We look forward to the possibility of collaborating with """ + business_name + """ and creating a partnership that resonates with audiences nationwide.

Warm regards,

Naren  
Rutgers Asian A Cappella Group  
ruraag@gmail.com """

    em = EmailMessage()
    em['From'] = email_sender 
    em['To'] = email_receiver
    em['Subject'] = subject
    em.set_content(body)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(email_sender, email_password)
            smtp.sendmail(email_sender, email_receiver, em.as_string())

url = 'https://docs.google.com/spreadsheets/d/12yUKeAG7kp0_UZuLFMkcPfjCzRc3f_zEK4SC5XNzxRI/edit?gid=0#gid=0'

def convert_google_sheet_url(url):
    # Regular expression to match and capture the necessary part of the URL
    pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)(/edit#gid=(\d+)|/edit.*)?'

    # Replace function to construct the new URL for CSV export
    # If gid is present in the URL, it includes it in the export URL, otherwise, it's omitted
    replacement = lambda m: f'https://docs.google.com/spreadsheets/d/{m.group(1)}/export?' + (f'gid={m.group(3)}&' if m.group(3) else '') + 'format=csv'

    # Replace using regex
    new_url = re.sub(pattern, replacement, url)

    return new_url


# Replace with modified URL 
url = 'https://docs.google.com/spreadsheets/d/12yUKeAG7kp0_UZuLFMkcPfjCzRc3f_zEK4SC5XNzxRI/edit?gid=0#gid=0'
new_url = convert_google_sheet_url(url)

print(new_url)
# https://docs.google.com/spreadsheets/d/1mSEJtzy5L0nuIMRlY9rYdC5s899Ptu2gdMJcIalr5pg/export?gid=1606352415&format=csv

dataframe1 = pd.read_csv(new_url)

print(dataframe1) 


for index, row in dataframe1.iterrows():
    if pd.isna(row["Initial correspondence sent:"]):  # Check if the correspondence is NaN or empty
        if not pd.isna(row["Name of Business:"]) and not pd.isna(row["Email:"]):  # Check if there is a business name in that row
            print(f"Missing correspondence for business: {row['Name of Business:']}")
            send_email(row["Name of Business:"], row['Email:']) 











