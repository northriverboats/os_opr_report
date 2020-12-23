#!/usr/bin/env python3
import pprint
import xlwt
import bgtunnel
import MySQLdb
import MySQLdb.cursors
import re
import sys
import os
import click
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
from hashlib import md5
from titlecase import titlecase
from emailer import *
from dotenv import load_dotenv

fields = [
    'submitted',
    'dealership',
    'model',
    'hull_serial_number',
    'date_delivered',
    'agency',
    'first_name',
    'last_name',
    'phone_home',
    'email',
    'mailing_address',
    'mailing_city',
    'mailing_state',
    'mailing_zip',
]

states = {
    'Alaska': 'AK',
    'Alabama': 'AL',
    'Arkansas': 'AR',
    'American Samoa': 'AS',
    'Arizona': 'AZ',
    'California': 'CA',
    'Colorado': 'CO',
    'Connecticut': 'CT',
    'District of Columbia': 'DC',
    'Delaware': 'DE',
    'Florida': 'FL',
    'Georgia': 'GA',
    'Hawaii': 'HI',
    'Iowa': 'IA',
    'Idaho': 'ID',
    'Illinois': 'IL',
    'Indiana': 'IN',
    'Kansas': 'KS',
    'Kentucky': 'KY',
    'Louisiana': 'LA',
    'Massachusetts': 'MA',
    'Maryland': 'MD',
    'Maine': 'ME',
    'Michigan': 'MI',
    'Minnesota': 'MN',
    'Missouri': 'MO',
    'Mississippi': 'MS',
    'Montana': 'MT',
    'National': 'NA',
    'North Carolina': 'NC',
    'North Dakota': 'ND',
    'Nebraska': 'NE',
    'New Hampshire': 'NH',
    'New Jersey': 'NJ',
    'New Mexico': 'NM',
    'Nevada': 'NV',
    'New York': 'NY',
    'Ohio': 'OH',
    'Oklahoma': 'OK',
    'Oregon': 'OR',
    'Pennsylvania': 'PA',
    'Puerto Rico': 'PR',
    'Rhode Island': 'RI',
    'South Carolina': 'SC',
    'South Dakota': 'SD',
    'Tennessee': 'TN',
    'Texas': 'TX',
    'Utah': 'UT',
    'Virginia': 'VA',
    'Virgin Islands': 'VI',
    'Vermont': 'VT',
    'Washington': 'WA',
    'Wisconsin': 'WI',
    'West Virginia': 'WV',
    'Wyoming': 'WY'
}


"""
Levels
0 = no output
1 = minimal output
2 = verbose outupt
3 = very verbose outupt
"""
dbg = 0
def debug(level, text):
    if dbg > (level -1):
        print(text)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def mail_results(subject, body, attachment=None):
    mFrom = os.getenv('MAIL_FROM')
    mTo = os.getenv('MAIL_TO')
    m = Email(os.getenv('MAIL_SERVER'))
    m.setFrom(mFrom)
    for email in mTo.split(','):
      m.addRecipient(email)
    # m.addCC(os.getenv('MAIL_FROM'))

    m.setSubject(subject)
    m.setTextBody("You should not see this text in a MIME aware reader")
    m.setHtmlBody(body)
    if (attachment):
        m.addAttachment(attachment)
    m.send()


def fetch_oprs(report_start):
    # connect to mysql on the server
    silent = dbg < 1
    forwarder = bgtunnel.open(
        ssh_user=os.getenv('SSH_USER'),
        ssh_address=os.getenv('SSH_HOST'),
        ssh_port=os.getenv('SSH_PORT'),
        host_port=3306,
        bind_port=3308,
        silent=silent
    )
    conn= MySQLdb.connect(
        host='127.0.0.1',
        port=3308,
        user=os.getenv('DB_USER'),
        passwd=os.getenv('DB_PASS'),
        db=os.getenv('DB_NAME'),
        cursorclass=MySQLdb.cursors.DictCursor
    )

    cursor = conn.cursor()

    # select all records from the OPR table
    sql = """
          SELECT  submitted, model, dealership, hull_serial_number, date_delivered,
                  agency, first_name,last_name, phone_home, email,
                  mailing_address, mailing_city, mailing_state, mailing_zip
            FROM  wp_nrb_opr
            WHERE  submitted > '{:%Y-%m-%d}'
        ORDER BY  submitted DESC
    """.format(report_start)

    total = cursor.execute(sql) # not used
    oprs = cursor.fetchall()

    cursor.close()
    conn.close()
    forwarder.close()

    return oprs



def write_sheet(oprs, xlsfile):
    wb = load_workbook(filename = xlsfile)
    ws = wb.active
    for row, opr in enumerate(oprs, start=2):
        opr['submitted'] = opr['submitted'].date()
        for column, field in enumerate(fields, start=1):
            _ = ws.cell(row=row, column=column, value=opr[field])
    title = datetime.now().strftime("%b %-d, %Y")
    filename = datetime.now().strftime("OPR Sales %Y-%m-%d.xlsx")
    ws.title = title
    wb.save(filename = resource_path(filename))
    return filename



@click.command()
@click.option('--debug', '-d', is_flag=True, help='show debug output')
@click.option('--verbose', '-v', default=1, type=int, help='verbosity level 0-3')
def main(debug, verbose):
    global dbg
    if debug:
        dbg = verbose

    # load environmental variables
    load_dotenv(resource_path(".env-local"))
    xlsfile = resource_path(os.getenv('XLSFILE'))

    report_date = datetime.now()
    report_start = datetime.now() - timedelta(days=int(os.getenv('INTERVAL')))

    try:
        oprs = fetch_oprs(report_start)
        filename = write_sheet(oprs, xlsfile)
        mail_results(
            filename[:-5],
            '<p>Here is the ' + os.getenv('INTERVAL_TITLE') + ' OS OPR Sales Report.</p>',
            attachment = filename
        )
        os.remove(resource_path(filename))
    except Exception as e:
        mail_results(
            'OS OPR Sales Processing Error',
            '<p>Spreadsheet can not be updated due to script error:<br />\n' + str(e) + '</p>'
        )

if __name__ == "__main__":
    main()
