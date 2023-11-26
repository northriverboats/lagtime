#!/usr/bin/env python3

import click
import os
import sys
from lagreport import *
from dotenv import load_dotenv
from envelopes import Envelope
from smtplib import SMTPException # allow for silent fail in try exception


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """

    try:
        base_path = sys._MEIPASS  # pylint: disable=protected-access
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def split_address(email_address):
    """Return a tuple of (address, name), name may be an empty string
       Can convert the following forms
         exaple@example.com
         <example@exmaple.con>
         Example <example@example.com>
         Example<example@example.com>
    """
    address = email_address.split('<')
    if len(address) == 1:
        return (address[0], '')
    if address[0]:
        return (address[1][:-1], address[0].strip())
    return (address[1][:-1], '')

def mail_results(subject, body, attachment=''):
    """ Send emial with html formatted body and parameters from env"""
    envelope = Envelope(
        from_addr=split_address(os.environ.get('MAIL_FROM')),
        subject=subject,
        html_body=body
    )

    # add standard recepients
    tos = os.environ.get('MAIL_TO','').split(',')
    if tos[0]:
        for to in tos:
            envelope.add_to_addr(to)

    # add carbon coppies
    ccs = os.environ.get('MAIL_CC','').split(',')
    if ccs[0]:
        for cc in ccs:
            envelope.add_cc_addr(cc)

    # add blind carbon copies recepients
    bccs = os.environ.get('MAIL_BCC','').split(',')
    if bccs[0]:
        for bcc in bccs:
            envelope.add_bcc_addr(bcc)

    if attachment:
        envelope.add_attachment(attachment)

    # send the envelope using an ad-hoc connection...
    try:
        _ = envelope.send(
            os.environ.get('MAIL_SERVER'),
            port=os.environ.get('MAIL_PORT'),
            login=os.environ.get('MAIL_LOGIN'),
            password='zcrkyqvgbxkxnjdg',
            tls=True
        )
    except SMTPException:
        print("SMTP EMail error")


def processAndEmail():
    """
    if (datetime.date.today().isocalendar()[1] % 2 == 0 ):
       return
   """

    lagReport()

    htmlText = """<p>Here is the Lag Time Report for %s.</p>
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

    plainText = """Here is the Lag Time Report for %s.

--
North River Boats "For Waters Less Traveled"
Fred Warren - Computer Support Specialist
Phone: 541-673-2438x140 Fax: 541-67902818
Email: fredw@northriverboats.com"""%(datetime.date.today())

    mail_results(
        'Lag Time Report %s'%(datetime.date.today()),
        htmlText,
        attachment='/tmp/LagReport-%s.xlsx'%(datetime.date.today()))


@click.command()
def main():
    load_dotenv(dotenv_path=resource_path(".env"))
    processAndEmail()
    sys.exit(0)

if __name__ == "__main__":
    main()
