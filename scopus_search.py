from pybliometrics.scopus import ScopusSearch
from pybliometrics.scopus import AuthorRetrieval
from pybliometrics.scopus import AffiliationRetrieval
from pybliometrics.scopus.exception import Scopus404Error

import json
import re
import argparse

import sys
import xlsxwriter

from datetime import date

from bs4 import BeautifulSoup
import requests

from email_scraper import scrape_emails





parser = argparse.ArgumentParser()
group = parser.add_mutually_exclusive_group()
group.add_argument("-k", "--keywords", help="keywords (AND/OR separated) to search in paper title, abstract, keywords")
group.add_argument("-j", "--references-json", help="JSON file with list of references (from anystyle)")
group.add_argument("-t", "--references-txt", help="TXT file with list of references (plain title)")
parser.add_argument('--min-recent-years', default=4, type=int, help="How many years ago the author should have published some papers")
parser.add_argument('--min-recent-papers', default=3, type=int, help="How many papers the author should have been published in recent years")
parser.add_argument('--min-h-index', default=3, type=int, help="The H-index of the author should be higher than a minimum")
parser.add_argument('--max-h-index', default=20, type=int, help="The H-index of the author should be lower than a maximum")
parser.add_argument('--max-reviewers', default=50, type=int, help="How many reviewers to search for")
parser.add_argument('--skip-first-results', default=-1, type=int, help="How many results from the Scopus query should be skipped")
parser.add_argument('--query-years', default=5, type=int, help="How many years ago the query should look in the past")
parser.add_argument('--journal-only', action=argparse.BooleanOptionalAction, default=True, help="Query should only look for journal papers")
parser.add_argument('--publisher', help="Publishers to be considered (comma separated)")
parser.add_argument('--cs-only', action=argparse.BooleanOptionalAction, default=True, help="Query should only look for Computer Science papers")
parser.add_argument('--conflicts', help="Affiliations to be excluded from the results (comma separated)")
parser.add_argument('-e', '--email-lookup', action=argparse.BooleanOptionalAction, default=True, help="Look-up for email addresses (requires Chrome running with remote debugging, and logged into Scopus)")


args = parser.parse_args()



from selenium import webdriver
from selenium.webdriver.chrome.options import Options

browser = None

if args.email_lookup is True:

    opt = Options()
    opt.add_experimental_option("debuggerAddress", "localhost:8989")

    try:
        browser = webdriver.Chrome(options=opt)
        browser.set_page_load_timeout(30)

    except:
        print("Unable to connect to Chrome with remote debugging. Email lookup will not work.\n")
        print("Run Chrome with remote debugging and logged-in into Scopus, then try again.\n\n")
        print("/Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=8989  www.scopus.com\n\n")
        exit(1)

    # Check if logged-in

    browser.get('https://www.scopus.com/')
    page = browser.page_source

    soup = BeautifulSoup(page, "lxml")

    email = None

    for x in soup.findAll('script'):
        match = re.search(r'ScopusUser\s*=\s*{(.*?)};\s*\n', str(x), flags=re.DOTALL)

        if not match is None:
            email = re.search(r'email:\s"(.*)"', str(match[1]))[1]

    if email is None or email == "":
        print("Chrome browser session must be logged-in into Scopus, email lookup will not work")
        exit(1)


reviewer_results = []
scopus_results = []


query_options = ""

if args.query_years > 0:
    query_options += ' AND PUBYEAR AFT {}'.format(date.today().year - args.query_years)

if args.journal_only is True:
    query_options += ' AND SRCTYPE(j)'

if args.cs_only is True:
    query_options += ' AND SUBJAREA(COMP)'

if args.publisher is not None:
    query_options += ' AND ('
    query_options += " OR ".join(['PUBLISHER({})'.format(publisher) for publisher in args.publisher.split(',')])
    query_options += ')'

if args.keywords:

    print("Seaching on Scopus: {}".format(args.keywords))

    s = ScopusSearch('TITLE-ABS-KEY ( {} ) {}'.format(args.keywords, query_options), verbose=True)

    if s.results is None:
        print("No results found for: {}".format(args.keywords))
        exit(1)

    scopus_results = s.results

elif args.references_txt:

    with open(args.references_txt) as f:

        refs = f.readlines()

        for ref in refs:

            query = re.sub("[^a-zA-Z0-9-\s]+", "", ref)

            print("Seaching on Scopus: {}".format(query))

            s = ScopusSearch('TITLE-ABS-KEY ( {} ) {}'.format(query, query_options), verbose=True)

            try:
                if s.results == 0:
                    print("No results found for: {}".format(query))
                    continue

                scopus_results.extend(s.results)
            except TypeError:
                pass

elif args.references_json:

    with open(args.references_json) as f:

        data = json.load(f)

        for paper in data:

            if 'date' in paper and 'title' in paper:

                year = int(paper['date'][0])

                if year >= (date.today().year - args.min_recent_years):

                    title = paper['title'][-1]

                    title = re.sub(r'\s\d+\s', ' ', title)

                    print("Reference found: " + title)

                    s = ScopusSearch('TITLE ( {} ) '.format(title))

                    try:
                        scopus_results.extend(s.results)
                    except:
                        pass

else:
    parser.print_help()
    sys.exit(0)


conflicts = []

if args.conflicts is not None:
    conflicts = args.conflicts.split(',')
    scopus_results = [paper for paper in scopus_results if (paper.affilname is not None) and not any(conflict.lower() in paper.affilname.lower() for conflict in conflicts)]


result_num = 0

for scopus_paper in scopus_results:

    if args.max_reviewers != -1 and len(reviewer_results) >= args.max_reviewers:
        print("\nMax number of reviewers found ({}), exiting".format(len(reviewer_results)))
        break


    result_num = result_num+1

    if args.skip_first_results != -1 and result_num <= args.skip_first_results:
        print("Skipping result {}".format(result_num))
        continue

    print("")
    print("--- Result no.: {} ---".format(result_num))



    if scopus_paper.title is None:
        print("No title found, skipping")
        continue

    paper = scopus_paper.title


    if scopus_paper.author_ids is None:
        print("No author IDs found, skipping")
        continue

    author_ids = scopus_paper.author_ids.split(';')

    author_names = scopus_paper.author_names.split(';')



    paper_link = 'https://www.scopus.com/record/display.uri?origin=resultslist&eid='+scopus_paper.eid


    author_emails = []

    if args.email_lookup is True:

        browser.get(paper_link)
        page = browser.page_source

        soup = BeautifulSoup(page, "lxml")

        author_list_tag = soup.find("div", {"data-testid": "author-list"})

        if author_list_tag is not None:

            for author_item in author_list_tag.findAll("li"):

                scraped = scrape_emails(str(author_item))

                if scraped is not None and len(scraped) > 0:
                    author_emails.append(list(scraped)[0])
                else:
                    author_emails.append("")

        else:
            author_emails = [''] * len(author_ids)



    for author_idx in range(len(author_ids)):

        auid = author_ids[author_idx]

        au = None

        try:
            au = AuthorRetrieval(auid)
        except:
            print("Author could not be retrieved, skipping")
            continue

        #indexed_name = au.indexed_name

        if au.given_name is None or au.surname is None:
            print("No author name and surname found, skipping")
            continue

        name = au.given_name
        surname = au.surname

        if au.name_variants is not None:
            for name_var in au.name_variants:
                if name_var.given_name is not None and len(name) < len(name_var.given_name):
                    name = name_var.given_name

                if name_var.surname is not None and len(surname) < len(name_var.surname):
                    surname = name_var.surname

        print("")
        print("Name: {}".format(name))
        print("Surname: {}".format(surname))



        if au.h_index is None:
            print("No h-index found, skipping")
            continue

        h_index = au.h_index
        au_id = au.eid
        au_link = au.self_link

        print("H-index: "+str(au.h_index))
        print("Self-link: "+au.self_link)

        if int(h_index) < args.min_h_index or int(h_index) > args.max_h_index:
            print("H-index out of range, skipping")
            continue



        email = author_emails[author_idx]

        if args.email_lookup is True and email == '':
            print("No email found, skipping")
            continue

        print("Email: "+email)





        docs = ScopusSearch('AU-ID({}) AND PUBYEAR > {}'.format(auid, date.today().year - args.min_recent_years), download=False)

        recent_docs = docs.get_results_size()

        print("Recent docs: {}".format(recent_docs))


        if recent_docs < args.min_recent_papers:
            print("Recent docs out of range, skipping")
            continue


        domain = None
        affiliation_name = None
        j = 0

        while domain is None and j < len(au.affiliation_current):
            try:
                affiliation = AffiliationRetrieval(au.affiliation_current[j].id)
                domain = affiliation.org_domain
                affiliation_name = affiliation.affiliation_name
            except Scopus404Error:
                pass

            j = j+1

        print("Domain: "+str(domain))




        result = {}

        result["Name"] = name
        result["Surname"] = surname
        result["H-index"] = h_index
        result["Author page"] = au_id
        result["Author page link"] = au_link
        result["Domain"] = domain
        result["Affiliation"] = affiliation_name
        result["Email"] = email
        result["Recent docs"] = recent_docs
        result["Recent paper"] = paper
        result["Recent paper link"] = paper_link

        reviewer_results.append(result)




workbook = xlsxwriter.Workbook('scopus_results.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

keys = ("Name", "Surname", "Email", "H-index", "Author page", "Domain", "Recent docs", "Recent paper")

for key in keys:
    worksheet.write(row, col, key)
    col = col + 1

row = 1

for result in reviewer_results:

    col = 0

    for key in keys:

        if key+" link" in result:

            worksheet.write_url(row, col, result[key+" link"], string=result[key])

        else:

            worksheet.write(row, col, result[key])

        col = col + 1

    row = row + 1

workbook.close()

