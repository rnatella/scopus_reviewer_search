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
max_reviewers = 50
skip_first_results = -1


search_results = []

if len(sys.argv) < 2:
    print("Error: {} <search AND/OR-separated strings>".format(sys.argv[0]))
    exit(1)

s = ScopusSearch('TITLE-ABS-KEY ( {} ) '.format(sys.argv[1]))

if s.results is None:
    print("No results found for: {}".format(sys.argv[1]))
    exit(1)


result_num = 0

for item in s.results:

    if max_reviewers != -1 and len(search_results) == max_reviewers:
        print("\nMax number of reviewers found ({}), exiting".format(max_reviewers))
        break


    result_num = result_num+1

    if skip_first_results != -1 and result_num <= skip_first_results:
        print("Skipping result {}".format(result_num))
        continue

    print("")
    print("--- Result no.: {} ---".format(result_num))



    if item.title is None:
        print("No title found, skipping")
        continue

    paper = item.title


    if item.author_ids is None:
        print("No author IDs found, skipping")
        continue

    author_ids = item.author_ids.split(';')

    for auid in author_ids:
        au = AuthorRetrieval(auid)

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
        au_link = au.self_link

        print("H-index: "+au.h_index)
        print("Self-link: "+au.self_link)

        if int(h_index) < min_h_index or int(h_index) > max_h_index:
            print("H-index out of range, skipping")
            continue


        docs = ScopusSearch('AU-ID({}) AND PUBYEAR > {}'.format(auid, cur_year - recent_years), download=False)

        recent_docs = docs.get_results_size()

        print("Recent docs: {}".format(recent_docs))


        if recent_docs < min_recent_papers:
            print("Recent docs out of range, skipping")
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
    