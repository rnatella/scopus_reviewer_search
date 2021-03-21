from pybliometrics.scopus import ScopusSearch
from pybliometrics.scopus import AuthorRetrieval
from pybliometrics.scopus import ContentAffiliationRetrieval

import json
import re
import argparse

import sys
import xlsxwriter

#from email_finder import email_finder



cur_year = 2020
recent_years = 3
min_recent_papers = 3
min_h_index = 3
max_h_index = 20
max_reviewers = 30
skip_first_results = -1



parser = argparse.ArgumentParser()
group = parser.add_mutually_exclusive_group()
group.add_argument("-k", "--keywords", help="keywords (AND/OR separated) to search in paper title, abstract, keywords")
group.add_argument("-j", "--references-json", help="JSON file with list of references (from anystyle)")
group.add_argument("-t", "--references-txt", help="TXT file with list of references (plain title)")

args = parser.parse_args()


reviewer_results = []
scopus_results = []

if args.keywords:

    print("Seaching on Scopus: {}".format(args.keywords))

    s = ScopusSearch('TITLE-ABS-KEY ( {} ) '.format(args.keywords))

    if s.results is None:
        print("No results found for: {}".format(args.keywords))
        exit(1)

    scopus_results = s.results

elif args.references_txt:

    with open(args.references_txt) as f:

        refs = f.readlines()

        for ref in refs:

            print("Seaching on Scopus: {}".format(ref))

            s = ScopusSearch('TITLE-ABS-KEY ( {} ) '.format(ref))

            try:
                if s.results == 0:
                    print("No results found for: {}".format(ref))
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

                if year >= (cur_year - recent_years):

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

result_num = 0

for scopus_paper in scopus_results:

    if max_reviewers != -1 and len(reviewer_results) >= max_reviewers:
        print("\nMax number of reviewers found ({}), exiting".format(max_reviewers))
        break


    result_num = result_num+1

    if skip_first_results != -1 and result_num <= skip_first_results:
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
        au_id = au.eid
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
        result["Author page"] = au_id
        result["Author page link"] = au_link
        result["Domain"] = domain
        #result["Email"] = email
        result["Recent docs"] = recent_docs
        result["Recent paper"] = paper
        result["Recent paper link"] = 'https://www.scopus.com/record/display.uri?origin=resultslist&eid='+scopus_paper.eid

        reviewer_results.append(result)

        


workbook = xlsxwriter.Workbook('scopus_results.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

keys = ("Name", "Surname", "H-index", "Author page", "Domain", "Recent docs", "Recent paper")

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
    
