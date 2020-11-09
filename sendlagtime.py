#!/usr/bin/env python3

from lagreport import *
from emailer import *
from dotenv import load_dotenv

# set python environment
if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
else:
    # we are running in a normal Python environment
    bundle_dir = os.path.dirname(os.path.abspath(__file__))

# load environmental variables
load_dotenv(bundle_dir + "/.env")


def processAndEmail():
    """
    if (datetime.date.today().isocalendar()[1] % 2 == 0 ):
       return
   """

    lagReport()

    htmlText = """Hi Patrick,
<p>Here is the Lag Time Report for %s.</p>
<font color="#888888"><br>
<br>
-- <br>
North River Boats "For Waters Less Traveled"<br>
Fred Warren - Computer Support Specialist<br>
Phone: <a href="tel:541-673-2438x140" value="+15416732438" target="_blank">541-673-2438x140</a>
Fax: <a href="tel:541-679-2818" value="+15416792818" target="_blank">541-679-2818</a><br>
Email: <a href="mailto:fredw@northriverboats.com" target="_blank"><span class="il">fredw@northriverboats.com</span></a><br>
<br>
</font>"""%(datetime.date.today())

    plainText = """Hi Dennis,

Here is the Lag Time Report for %s.

--
North River Boats "For Waters Less Traveled"
Fred Warren - Computer Support Specialist
Phone: 541-673-2438x140 Fax: 541-67902818
Email: fredw@northriverboats.com"""%(datetime.date.today())

    m = Email(os.getenv('MAIL_SERVER'))
    m.setFrom(os.getenv('MAIL_FROM'))

    if os.getenv('DEBUG').upper() == 'TRUE':
        m.addRecipient(os.getenv('MAIL_FROM'))
        print(os.getenv('MAIL_FROM'))
    else:
        mTo = os.getenv('MAIL_TO')
        for email in mTo.split(','):
            m.addRecipient(email)
            print(email)

        mCc = os.getenv('MAIL_CC')
        if mCc:
            for email in mCc.split(','''):
                m.addCC(email)
                print(email)

        mBcc = os.getenv('MAIL_BCC')
        if mBcc:
            for email in mBcc.split(','):
                m.addBCC(email)
                print(email)

    m.setSubject('Lag Time Report %s'%(datetime.date.today()))
    m.setTextBody(plainText)
    m.setHtmlBody(htmlText)
    m.addAttachment('/tmp/LagReport-%s.xlsx'%(datetime.date.today()))
    m.send()



if __name__ == "__main__":
    processAndEmail()

