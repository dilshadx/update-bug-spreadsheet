#!/usr/bin/python

"""Script to extract information from Bugzilla and populate a Google Spreadsheet."""

import urllib
import urllib2
import cookielib
import datetime
import argparse

from gdata.spreadsheet import service
from bs4 import BeautifulSoup

import user_config


class UpdateBugSpreadsheet(object):
    """Script to extract information from Bugzilla and populate a Google Spreadsheet."""

    def __init__(self, args):
        """Initialise variables and store data from user_config."""
        self.worksheet_data = {}
        self.loginurl = user_config.user_details['bugzilla']
        self.username = args.username if args.username else user_config.user_details['username']
        self.foundry_pass = args.found_pass if args.found_pass else user_config.user_details['foundry-password']
        self.email = args.email if args.email else user_config.user_details['email']
        self.bug_pass = args.bugPass if args.bugPass else user_config.user_details['bugzilla-password']
        self.google_pass = args.google_pass if args.google_pass else user_config.user_details['google-password']
        self.user_details = user_config.user_details['urls']
        self.terminal_only = args.terminalOnly
        self.spreadsheet_args = args.spreadsheet if args.spreadsheet else [user_config.user_details['spreadsheet']]
        self.worksheet_args = args.worksheet
        self.column_args = args.column

    def access_bugzilla_site(self):
        """Login to Bugzilla with details from user_config."""
        cookie_jar = cookielib.CookieJar()
        pass_manager = urllib2.HTTPPasswordMgrWithDefaultRealm()
        pass_manager.add_password(None, self.loginurl, self.username, self.foundry_pass)
        handler = urllib2.HTTPBasicAuthHandler(pass_manager)
        self.opener = urllib2.buildopener(handler, urllib2.HTTPCookieProcessor(cookie_jar))
        urllib2.installopener(self.opener)
        login_data = urllib.urlencode({'Bugzilla_login': self.email, 'Bugzilla_password': self.bug_pass})
        self.loginpage = self.opener.open(self.loginurl, login_data)

    def extract_data_from_sites(self):
        """Open bugzilla webpages and extract data from them."""
        worksheets = self.worksheet_args if self.worksheet_args else self.user_details.keys()
        for worksheet in worksheets:
            row_data = {}
            print('\n...Worksheet: %s...' % worksheet)
            row_data['date'] = (datetime.datetime.now().strftime('%d/%m/%Y'))
            print('Date: %s' % row_data['date'])
            column_headings = self.user_details[worksheet].items()

            # Filter column headings for user selection
            if self.column_args:
                column_headings = [heading for heading in column_headings if heading[0] in self.column_args]

            for url in sorted(column_headings):
                page = self.opener.open(url[1])
                soup = BeautifulSoup(page.read())

                # Variables for information needed to extract
                whole_element = soup.find_all('span', {'class': 'bz_result_count'})
                split_element = (whole_element[0].string).split(' ')[0]

                # Check for a Zarro return and convert to 0
                if split_element == 'Zarro': split_element = '0'
                # Check for a \n return and convert to 1
                elif split_element == '\n': split_element = '1'

                # Print information for user
                print('%s: %s' % (url[0], split_element))

                # Make characters lowercase and remove whitespace ( needed for google docs )
                row_data[(url[0].lower().replace(' ', ''))] = split_element
                self.worksheet_data[worksheet] = row_data
                page.close()

    def google_docs_log_on(self):
        """Log on to the google docs service with user_config data."""
        self.spr_client = service.SpreadsheetsService()
        self.spr_client.email = self.email
        self.spr_client.password = self.google_pass
        self.spr_client.source = 'Update Bug Spreadsheet Script'
        self.spr_client.ProgrammaticLogin()

    def access_googlespreadsheet(self):
        """Access spreadsheet and worksheets named in the user_config and populate with data."""
        spr_sheet = service.DocumentQuery()
        for spreadsheet in self.spreadsheet_args:
            spr_sheet['title'] = spreadsheet
            spr_sheet['title-exact'] = 'true'
            spr_feed = self.spr_client.GetSpreadsheetsFeed(query=spr_sheet)
            spreadsheet_key = spr_feed.entry[0].id.text.rsplit('/', 1)[1]
            work_feed = self.spr_client.GetWorksheetsFeed(spreadsheet_key)

            for worksheet in work_feed.entry:
                if worksheet.title.text in self.worksheet_data.keys():
                    worksheet_data = self.worksheet_data[worksheet.title.text]
                    worksheet_id = worksheet.id.text.rsplit('/', 1)[1]
                    self.spr_client.InsertRow(row_data=worksheet_data, key=spreadsheet_key, wksht_id=worksheet_id)

            print('\nSpreadsheet %s Updated' % spreadsheet)

    def run(self):
        """Run through each function."""
        self.access_bugzilla_site()
        self.extract_data_from_sites()
        if not self.terminal_only:
            self.google_docs_log_on()
            self.access_googlespreadsheet()
            self.loginpage.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-u', '--username')
    parser.add_argument('-fp', '--found_pass')
    parser.add_argument('-e', '--email')
    parser.add_argument('-bp', '--bugPass')
    parser.add_argument('-gp', '--google_pass')
    parser.add_argument('-s', '--spreadsheet', nargs='*')
    parser.add_argument('-w', '--worksheet', nargs='*')
    parser.add_argument('-c', '--column', nargs='*')
    parser.add_argument('-to', '--terminal_only', action='store_true')
    args = parser.parse_args()

    UpdateBugSpreadsheet(args).run()
