#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#   aw2gd.py - Last modified: Tue 03 Jan 2012 11:03:05 AM CST
#
#   Copyright (C) 2011 Richard A. Johnson <nixternal@gmail.com>
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""Parses Active Works registration report CSV file into a Google Docs
Spreadsheet"""

import csv
import getopt
import os
import re
import sys

try:
    from progressbar import ProgressBar, Bar, SimpleProgress
except ImportError:
    print "You need 'python-progressbar' installed"
    sys.exit(3)

try:
    import gdata.service
    import gdata.spreadsheet
    import gdata.spreadsheet.service
except ImportError:
    print "Either 'python-gdata' is not installed or it isn't configured."
    sys.exit(3)

# DEBUG - Enabled prints to stdout, Disabled goes to GDocs
DEBUG = 0

# Google Docs Spreadsheet Name
SPREADSHEET = '2012_TTSeries_Reg_Results'

# Event year for calculating racing age
YEAR = 2012

# Categories
CATS = {
    # Age Categories
    '10-14': (10, 14),
    '15-19': (15, 19),
    '20-24': (20, 24),
    '25-29': (25, 29),
    '30-34': (30, 34),
    '35-39': (35, 39),
    '40-44': (40, 44),
    '45-49': (45, 49),
    '50-54': (50, 54),
    '55-59': (55, 59),
    '60-64': (60, 64),
    '65-69': (65, 69),
    '70-74': (70, 74),
    '75-79': (75, 79),
    '80+'  : (80, 120),
    # Ability Categories
    '1/2/3': ('Category 1/2/3 (Elite Level)'),
    '4/5'  : ('Category 4/5 (Amateur Level)')
}

# Time Slots
TIMESLOTS = {
    '1': '9:30am or earlier',
    '2': '9:30am to 11:00am',
    '3': '11:00am to 1:00pm',
    '4': '1:00pm to 2:30pm',
    '5': '2:30pm or later (if available)'
}


def read_csv_file(fname):
    """Read the CSV file and return a list"""
    reader = csv.DictReader(open(fname, 'rb'), delimiter=',', quotechar='"')
    csvlist = []
    for row in reader:
        csvlist.append(row)
    return csvlist


def get_events(events):
    """Parse the data and return the events rider is participating in."""
    if 'Series 4 Race Deal' in events:
        return str('1'), str('1'), str('1'), str('1')
    elif '#1' in events:
        return str('1'), None, None, None
    elif '#2' in events:
        return None, str('1'), None, None
    elif '#3' in events:
        return None, None, str('1'), None
    elif '#4' in events or 'John Fraser Memorial Time Trial' in events:
        return None, None, None, str('1')


def get_category(cattype, catdata):
    """Parse the data and return the correct category. Either age-based, or
    ability based if racing twice.
    @cattype If true then age, else parse category 1/2/3 and 4/5 data
    """
    if cattype:
        racingage = YEAR - int(catdata.split('/')[2])
        for cat in CATS.keys():
            if '1/2/3' not in cat and '4/5' not in cat:
                if racingage in range(int(CATS[cat][0]), int(CATS[cat][1])+1):
                    return cat
    for cat in CATS.keys():
        if catdata in CATS[cat]:
            return cat


def cleanup_rider_list(riderlist):
    """Cleans up the rider list. If rider racing twice, adds new rider to the
    list. Makes writing to GDocs easier in the end."""
    riders = []
    for rider in riderlist:
        nrider = {}
        print rider['Name: Last name'], rider['Name: First name']
        nrider['tt1'], nrider['tt2'], nrider['tt3'], nrider['tt4'] = get_events(
                rider['Registration category'])
        nrider['rcvddate'], nrider['rcvdtime'] = rider['Registration time'].split(' ')
        nrider['namefirst'] = rider['Name: First name'].title()
        nrider['namelast'] = rider['Name: Last name'].title()
        nrider['bdate'] = rider['Date of birth']
        nrider['gender'] = rider['Gender'][0]
        nrider['email'] = rider['Email']
        nrider['phoneday'] = rider['Day phone']
        nrider['phoneeve'] = rider['Evening phone']
        nrider['phonecell'] = rider['Cell phone']
        nrider['street'] = rider['Contact address: Address1'].title()
        nrider['apt'] = rider['Contact address: Address2'].title()
        nrider['city'] = rider['Contact address: City'].title()
        nrider['state'] = rider['Contact address: State Province Region']
        nrider['zip'] = rider['Contact address: ZIP/Postal code']
        nrider['emcontact'] = rider['Emergency contact name'].title()
        nrider['emphone'] = rider['Emergency contact phone'].title()
        nrider['abrlicense'] = rider['ABR License Number']
        nrider['club'] = rider['Cycling club'].title()
        nrider['catprimary'] = '%s%s' % (nrider['gender'],
                get_category(True, rider['Date of birth']))
        nrider['timeprefstart'] = rider['Desired Start Time']
        for key in TIMESLOTS.keys():
            if nrider['timeprefstart'] == TIMESLOTS[key]:
                nrider['timeprefstart'] = key
        if rider['Category']:
            nrider['catsecondary'] = '%s%s' % (nrider['gender'],
                    get_category(False, rider['Category']))
            #nrider['catsecondary'] = get_category(False, rider['Category'])
            nrider_ = nrider.copy()
            nrider_['catprimary'] = nrider_['catsecondary']
            del nrider_['catsecondary']
            try:
                nrider_['timebetweenraces'] = rider['Time Between Races']
            except:
                nrider_['timebetweenraces'] = 45
            riders.append(nrider_)
        riders.append(nrider)
    return riders


def send_to_gdocs(user, pw, riders):
    """Sends the riders to Google Docs spreadsheet."""
    gdclient = gdata.spreadsheet.service.SpreadsheetsService()
    gdclient.email = user
    gdclient.password = pw
    gdclient.source = 'ActiveWorks Registration to Google Spreadsheet'
    gdclient.ProgrammaticLogin()
    keyfeed = gdclient.GetSpreadsheetsFeed()
    spreads = []
    for i in keyfeed.entry:
        spreads.append(i.title.text)
    snum = None
    for i, j in enumerate(spreads):
        if j == SPREADSHEET:
            snum = i
    if snum is None:
        sys.exit(1)
    keyparts = keyfeed.entry[snum].id.text.split('/')
    currkey = keyparts[len(keyparts)-1]
    wkshtfeed = gdclient.GetWorksheetsFeed(currkey)
    wkshtparts = wkshtfeed.entry[0].id.text.split('/')
    currwkshtid = wkshtparts[len(wkshtparts)-1]
    pbar = ProgressBar(widgets=[SimpleProgress(), Bar()],
            maxval=len(riders)).start()
    counter = 0
    for rider in riders:
        for key in rider.keys():
            if not rider[key] or 'None' in rider[key]:
                del rider[key]
        entry = gdclient.InsertRow(rider, currkey, currwkshtid)
        if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            counter += 1
            pbar.update(counter)
        else:
            print 'FAILURE! %s, %s - %s did not insert into spreadsheet' % (
                rider['namelast'], rider['namefirst'], rider['catprimary'])
    pbar.finish()


def main():
    """Main function that runs the program"""
    user = None
    pw = None
    ifile = None

    try:
        opts, args = getopt.getopt(sys.argv[1:], "", ["user=", "pw=",
            "ifile="])
    except getopt.error, msg:
        print 'aw2gd.py --user [google username] --pw [google password] --ifile [CSV file]'
        sys.exit(2)
    for o, a in opts:
        if o == "--user":
            user = a
        elif o == "--pw":
            pw = a
        elif o == "--ifile":
            ifile = a
    if not user or not pw or not ifile:
        print 'aw2gd.py --user [google username] --pw [google password] --ifile [CSV file]'
        sys.exit(2)

    sys.stdout.write(
        '\x1b[31m%s\x1b[0m\n' % 'Run in (d)ebug mode, (p)rint debug mode, or '
        '(r)elease mode?')
    input_ = raw_input('d/p/r: ')
    stdoutprint = False
    releasemode = False
    if input_ == 'p':
        stdoutprint = True
    if input_ == 'r':
        releasemode = True

    csvlist = read_csv_file(ifile)
    riders = cleanup_rider_list(csvlist)

    if stdoutprint:
        from pprint import PrettyPrinter
        printer = PrettyPrinter(indent=4)
        for rider in riders:
            printer.pprint(rider)

        print '##### DEBUG INFO #####'
        counter = 0
        for rider in riders:
            try:
                if rider['catsecondary']:
                    counter += 1
            except KeyError:
                pass
        print 'Total registrants parsed: %s ' % (len(riders)-counter)
        print 'Total racing twice: %s' % counter
        print 'Total racing: %s' % len(riders)
        print '\n~~~~~ Categories for people racing twice ~~~~~'
        counter = 0
        for rider in riders:
            try:
                if rider['catsecondary']:
                    counter += 1
                    print '%s - %s - %s - %s - %s, %s' % (counter, rider['bdate'],
                        rider['catprimary'], rider['catsecondary'], rider['namelast'],
                        rider['namefirst'])
            except KeyError:
                pass
        print '\n~~~~~ Riders Preferred Timeslots ~~~~~'
        ts1 = 0
        ts2 = 0
        ts3 = 0
        ts4 = 0
        ts5 = 0
        for rider in riders:
            if int(rider['timeprefstart']) == 1: ts1+=1
            elif int(rider['timeprefstart']) == 2: ts2+=1
            elif int(rider['timeprefstart']) == 3: ts3+=1
            elif int(rider['timeprefstart']) == 4: ts4+=1
            elif int(rider['timeprefstart']) == 5: ts5+=1
        print '9:30am or earlier:\t%s' % ts1
        print '9:30am to 11:00am:\t%s' % ts2
        print '11:00am to 1:00pm:\t%s' % ts3
        print '1:00pm to 2:30pm:\t%s' % ts4
        print '2:30pm or later:\t%s' % ts5

    if releasemode:
        send_to_gdocs(user, pw, riders)


if __name__ == '__main__':
    main()
