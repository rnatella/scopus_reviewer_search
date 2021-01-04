from re import search
from pybliometrics.scopus import ScopusSearch
from pybliometrics.scopus import AuthorRetrieval
from pybliometrics.scopus import ContentAffiliationRetrieval

import sys
import xlsxwriter

from email_finder import email_finder



cur_year = 2020
recent_years = 3
min_recent_papers = 3
min_h_index = 3
max_h_index = 20


search_results = []

if len(sys.argv) < 2:
    print("Error: {} <search AND/OR-separated strings>".format(sys.argv[0]))
    exit(1)

s = ScopusSearch('TITLE-ABS-KEY ( {} ) '.format(sys.argv[1]))

if s.results is None:
    print("No results found for: {}".format(sys.argv[1]))
    exit(1)


i = 0

for item in s.results:
    i = i+1


    if i == 100:
        break


    if item.title is None:
        continue

    paper = item.title


    if item.author_ids is None:
        continue

    author_ids = item.author_ids.split(';')

    for auid in author_ids:
        au = AuthorRetrieval(auid)

        #indexed_name = au.indexed_name

        if au.given_name is None or au.surname is None:
            continue

        name = au.given_name
        surname = au.surname

        if au.name_variants is not None:
            for name_var in au.name_variants:
                if len(name) < len(name_var.given_name):
                    name = name_var.given_name
                
                if len(surname) < len(name_var.surname):
                    surname = name_var.surname

        print("Name: {}".format(name))
        print("Surname: {}".format(surname))



        if au.h_index is None:
            continue

        h_index = au.h_index
        au_link = au.self_link

        print("H-index: "+au.h_index)
        print("Self-link: "+au.self_link)

        if int(h_index) < min_h_index or int(h_index) > max_h_index:
            continue


        docs = ScopusSearch('AU-ID({}) AND PUBYEAR > {}'.format(auid, cur_year - recent_years), download=False)

        recent_docs = docs.get_results_size()

        print("Recent docs: {}".format(recent_docs))


        if recent_docs < min_recent_papers:
            continue


        domain = None
        j = 0

        while domain is None and j < len(au.affiliation_current):
            affiliation = ContentAffiliationRetrieval(au.affiliation_current[j].id)

            domain = affiliation.org_domain

            j = j+1

        print("Domain: "+str(domain))



        #email = ""

        #if name is not None and surname is not None and domain is not None:
        #    email = email_finder(name, surname, domain)
        #    print("Email: {}".format(email))


        result = {}

        result["Name"] = name
        result["Surname"] = surname
        result["H-index"] = h_index
        result["Author Link"] = au_link
        result["Domain"] = domain
        #result["Email"] = email
        result["Recent docs"] = recent_docs
        result["Recent paper"] = paper

        search_results.append(result)
        


workbook = xlsxwriter.Workbook('scopus_results.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

keys = ("Name", "Surname", "H-index", "Author Link", "Domain", "Recent docs", "Recent paper")

for key in keys:
    worksheet.write(row, col, key)
    col = col + 1

row = 1

for result in search_results:

    col = 0

    for key in keys:
        worksheet.write(row, col, result[key])
        col = col + 1
    
    row = row + 1

workbook.close()
    