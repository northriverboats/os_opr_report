#!/usr/bin/env python3
import pprint
from xlrd import open_workbook, XLRDError, xldate_as_tuple
import xlwt
import bgtunnel
import MySQLdb
import MySQLdb.cursors
import re
import sys
import os
import click
from hashlib import md5
from xlutils.copy import copy
from titlecase import titlecase
from emailer import *
from dotenv import load_dotenv

xlsfile = ''

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

def read_workbook():
    # Read boat/dealer/model from spreadsheet    use  book.release_resources() before saving
    book = open_workbook(xlsfile, formatting_info=True, on_demand=True)

    # dont assume it is the first sheet, scan thru all sheets looking for 'DEALER'
    sx = 0
    for i in range(0,len(book.sheet_names())-1):
        if (book.sheet_names()[i] == 'DEALER'):
            sx = i

    sh = book.sheet_by_index(sx) # read-only copy    sh.cell(row,col).value sh.cell_value(row,col)
    wb = copy(book)             # to write to file  wb.save('filename')
    ws = wb.get_sheet(sx)        # write-only copy   ws.write(row,col,'value')


    # build dictionary of hullserial to row
    hulls = {}
    nulls = 0

    for rx in range(sh.nrows):
        hull = sh.cell_value(rowx=rx, colx=0)

        if (hull[:3] != 'NRB'):
            nulls += 1
            if nulls > 6:
                break
            else:
                continue
        hulls[hull] = rx

    return book, hulls, sh, wb, ws


def fetch_oprs():
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
           WHERE  submitted > '2020-12-01'
        ORDER BY  submitted DESC
    """

    total = cursor.execute(sql) # not used
    oprs = cursor.fetchall()

    cursor.close()
    conn.close()
    forwarder.close()

    return oprs


def process_sheet(data, hulls, col, sh, ws):
    """
    For logic see gist:
        https://gist.github.com/northriverboats/bd05796844dee5ecdb493880b5e5e01d
    """
    font_size_style = xlwt.easyxf('font: name Garmond, bold off, height 240;')
    date_font_size_style = xlwt.easyxf('font: name Garmond, bold off, height 240;')
    date_font_size_style.num_format_str = 'mm/dd/yyyy'
    changed = 0

    output = "\n\n---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n"
    output += "| Hull         | Lastname        | Firstname  | Phone                | Mailing                                            | Street                                             | Purchased  |\n"
    output += "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n"

    pp = pprint.PrettyPrinter(indent=4)

    #               0      1      2      3      4     5      6     7
    truth_table = [True, False, False, False, True, False, True, False ]

    for datum in data:
        rx = hulls.get(datum.get('hull_serial_number')[:3] + datum.get('hull_serial_number')[4:9] + datum.get('hull_serial_number')[10:],0)
        if (rx):  # rx is row on sheet where datum's hull_serial_number shows up
            opr_char = sh.cell_value(rx, 19)
            css_char = sh.cell_value(rx, 20)
            opr_flag = (0, 2)[opr_char != '']
            css_flag = (0, 1)[css_char != '']
            flag = (col * 4) + opr_flag + css_flag
            if flag == 1:
                changed += 1
                ws.write(rx, 19 + col, 'X', font_size_style)
            if truth_table[flag]:
                changed += 1
                homephone = str(datum.get('phone_home','')).upper()
                workphone = str(datum.get('phone_work','')).upper()
                if (homephone == 'NA' or homephone == 'N/A' or homephone == 'NONE'):
                    homephone = ''
                if (workphone == 'NA' or workphone == 'N/A' or workphone == 'NONE'):
                    workphone = ''

                ws.write(rx,  1, titlecase(datum.get('last_name','')), font_size_style)
                ws.write(rx,  2, titlecase(datum.get('first_name','')), font_size_style)
                ws.write(rx,  3, (workphone, homephone)[bool(homephone)], font_size_style)
                ws.write(rx,  4, titlecase(datum.get('mailing_address','')), font_size_style)
                ws.write(rx,  5, titlecase(datum.get('mailing_city','')), font_size_style)
                ws.write(rx,  6, states.get(datum.get('mailing_state',''),''), font_size_style)
                ws.write(rx,  7, datum.get('mailing_zip','').upper(), font_size_style)
                if col == 0:
                    ws.write(rx,  8, titlecase(datum.get('street_address','')), font_size_style)
                    ws.write(rx,  9, titlecase(datum.get('street_city','')), font_size_style)
                    ws.write(rx, 10, states.get(datum.get('street_state',''),''), font_size_style)
                    ws.write(rx, 11, datum.get('street_zip','').upper(), font_size_style)
                ws.write(rx, 12, datum.get('email_address', datum.get('email', '')), font_size_style)
                ws.write(rx, 13, datum.get('date_purchased','01/01/01'), date_font_size_style )
                ws.write(rx, 19 + col, 'X', font_size_style)

                output1 = "| %-12s | %-15s | %-10s | %-20s | %-50s | %-50s | %-10s | %s\n" % (
                    datum.get('hull_serial_number',''),
                    titlecase(datum.get('last_name',''))[:15],
                    titlecase(datum.get('first_name',''))[:10],
                    (workphone, homephone)[bool(homephone)][:20],
                    titlecase(
                        datum.get('mailing_address','') + ', ' + datum.get('mailing_city','')
                    ) + ', ' + states.get(datum.get('mailing_state',''),'') + ', ' + datum.get('mailing_zip', '').upper(),
                    titlecase(
                        datum.get('street_address','') + ', ' + datum.get('street_city','')
                    ) + ', ' + states.get(datum.get('street_state',''),'') + ', ' + datum.get('street_zip', '').upper(),
                    datum.get('date_purchased','01/01/01'), rx
                )
                debug(1, output1.replace('\n',''))
                output += output1


    return output, changed

def mail_results(subject, body):
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
    m.send()



@click.command()
@click.option('--debug', '-d', is_flag=True, help='show debug output')
@click.option('--verbose', '-v', default=1, type=int, help='verbosity level 0-3')
def main(debug, verbose):
    global xlsfile
    global dbg
    if debug:
        dbg = verbose

    # set python environment
    if getattr(sys, 'frozen', False):
        bundle_dir = sys._MEIPASS
    else:
        # we are running in a normal Python environment
        bundle_dir = os.path.dirname(os.path.abspath(__file__))

    # load environmental variables
    load_dotenv(bundle_dir + "/.env-local")

    xlsfile = os.getenv('XLSFILE')

    try:
        oprs = fetch_oprs()
        print(oprs)
        """
        book, hulls, sh, wb, ws = read_workbook()
        output_1, changed_1 = process_sheet(oprs, hulls, 0, sh, ws)
        output_2, changed_2 = process_sheet(csss, hulls, 1, sh, ws)
        output = "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n"
        output += "| Hull         | Lastname        | Firstname  | Phone                | Mailing                                            | Street                                             | Purchased  |\n"
        output += "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n"
        output = output_1 + output_2
        changed = changed_1 + changed_2

        if (changed and not dbg):
            wb.save(xlsfile)
            mail_results('OPR to Warranty Spreadsheet Update', '<pre>' + output + '</pre>')
        """
    except OSError:
        mail_results(
            'OPR to Warranty Spreadsheet is open',
            'OPR to Warranty Spreadsheet is open, spreadsheet can not be updated'
        )
    except Exception as e:
        mail_results(
            'OPR to Warranty Spreadsheet Processing Error',
            '<p>Spreadsheet can not be updated due to script error:<br />\n' + str(e) + '</p>'
        )

if __name__ == "__main__":
    main()
