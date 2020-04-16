# Processes a batch of ILLiad requests and checks different online databases
# to see if there's an available copy. Export ILLiad requests into folder
# python file is located. Rename file 'requests.xlsx.'

import requests, bs4,  xlrd, re, datetime

s = requests.Session()

#import your illiad requests in this spreadsheet
requests_workbook = xlrd.open_workbook('requests.xlsx')
books = requests_workbook.sheet_by_index(0)

headings = books.row_values(0)
title_col = headings.index('Loan Title')
author_col = headings.index('Loan Author')

#UNC catalog function searches proquest
#maybe add ISBN
def proquest_access(title,author):
    title_unc = title.replace(" ", "+")
    author_no_end_space = author.strip(" ")
    author_unc = author_no_end_space.replace(" ", "+")
    query = author_unc + "+" + title_unc
    cat_url = "https://catalog.lib.unc.edu/?utf8=%E2%9C%93&search_field=all_fields&q={query}&f%5Baccess_type_f%5D%5B%5D=Online".format(
        query=query)
    cat_get = s.get(cat_url)

    catSoup = bs4.BeautifulSoup(cat_get.text, 'html.parser')

    catSoup_tags = catSoup.select('#documents')
    if len(catSoup_tags) > 0:
        catSoup_ugly = catSoup_tags[0].getText()
        regex = re.compile('[^a-zA-Z]')
        catSoup_caps = regex.sub('',catSoup_ugly)
        search_text = catSoup_caps.lower()

        #title cleanup
        title_caps = regex.sub('', title)
        title_test = title_caps.lower()
        #author cleanup
        author_caps = regex.sub('', author)
        author_test = author_caps.lower()

        if title_test in search_text and author_test in search_text and 'proquest' in search_text:
            return 'Found in ProQuest',cat_url
        else:
            return "Not Found in ProQuest"
    else:
        return "Not Found in Proquest"

#cambridge access searcher
def cambridge_access(title, author):
    title_unc = title.replace(" ", "+")
    author_no_end_space = author.strip(" ")
    author_unc = author_no_end_space.replace(" ", "+")
    query = author_unc + "+" + title_unc
    cat_url = "https://catalog.lib.unc.edu/?utf8=%E2%9C%93&search_field=all_fields&q={query}&f%5Baccess_type_f%5D%5B%5D=Online".format(
                query=query)
    cat_get = s.get(cat_url)

    catSoup = bs4.BeautifulSoup(cat_get.text, 'html.parser')

    catSoup_tags = catSoup.select('#documents')
    if len(catSoup_tags) > 0:
        catSoup_ugly = catSoup_tags[0].getText()
        regex = re.compile('[^a-zA-Z]')
        catSoup_caps = regex.sub('', catSoup_ugly)
        search_text = catSoup_caps.lower()

        # title cleanup
        title_caps = regex.sub('', title)
        title_test = title_caps.lower()
        # author cleanup
        author_caps = regex.sub('', author)
        author_test = author_caps.lower()

        if title_test in search_text and 'cambridge' in search_text and author_test in search_text:
            return 'Available through Cambridge Press Online',cat_url

        else:
            return "Not Found in Cambridge Press."
    else:
        return "Not Found in Cambridge Press."


def hathi_access(title, author):
    title_unc = title.replace(" ", "+")
    author_no_end_space = author.strip(" ")
    author_unc = author_no_end_space.replace(" ", "+")
    query = author_unc + "+" + title_unc
    hathi_url = "https://catalog.lib.unc.edu/?utf8=%E2%9C%93&search_field=all_fields&q={query}&f%5Baccess_type_f%5D%5B%5D=Online".format(
        query=query)
    hathi_get = s.get(hathi_url)

    hathiSoup = bs4.BeautifulSoup(hathi_get.text, 'html.parser')
    hathi_elements = hathiSoup.find_all("a", {"class": "link-type-fulltext link-open-access"})
    author_tags = hathiSoup.select('#facet-author_facet_f')
    title_tags = hathiSoup.select('#documents')
    if len(author_tags) > 0 and len(title_tags) > 0:
        # string cleanup, test to see if author is there
        author_ugly = author_tags[0].getText()
        regex = re.compile('[^a-zA-Z]')
        authors_ugly_caps = regex.sub('', author_ugly)
        author_caps = regex.sub('', author)
        authors = authors_ugly_caps.lower()
        author_test = author_caps.lower()

        # string cleanup, test to see if title is there
        titles_ugly = title_tags[0].getText()
        titles_ugly_caps = regex.sub('', titles_ugly)
        title_caps = regex.sub('', title)
        titles = titles_ugly_caps.lower()
        title_test = title_caps.lower()

        # print(authors)
        # print(author_test)
        if len(hathi_elements) > 0 and author_test in authors and title_test in titles:
            hathi_list = [{elem.getText(): elem.get('href')} for elem in hathi_elements]
            return "Temporary access in Hathi.", hathi_url

        else:
            return "Not found in Hathi."

    else:
        return "Not found in Hathi"


###function for red_shelf_access
###pretty much works
def red_shelf_access(title, author):
    query = title + ' ' + author
    red_query = query.replace(" ", "%20")
    red_shelf_url = 'https://studentresponse.redshelf.com/search/?terms=%s' % red_query
    red_shelf = s.get(red_shelf_url)

    redSoup = bs4.BeautifulSoup(red_shelf.text, 'html.parser')
    red_items = redSoup.select(".price-content")
    if len(red_items) > 0:
        if "Borrow through" in red_items[0].getText():
            title_items = redSoup.select(".title-row")
            # print(title_items[0].getText())
            t1 = title_items[0].getText()
            t2 = "".join(t1.split())
            title_nospace = "".join(title.split())
            if title_nospace in t2:

                return "Found in Red Shelf.", red_shelf_url

            else:
                return "Not found in Red Shelf."

        else:
            return "Not found in Red Shelf"
    else:
        return "Not found in Red Shelf"


# open libray function
def open_library_access(title, author):
    title_open = title.replace(" ", "+")
    author_no_end_space = author.strip(" ")
    author_open = author_no_end_space.replace(" ", "+")
    open_lib_url = 'https://openlibrary.org/search?title=%s&author=%s' % (title_open, author_open)
    open_lib = s.get(open_lib_url)
    openSoup = bs4.BeautifulSoup(open_lib.text, 'html.parser')
    if 'No results found.' in open_lib.text:
        return 'Not found in Open Library.'

    else:
        open_items_elem = openSoup.select('.searchResultItemCTA-lending')
        acc = 0
        for x in range(len(open_items_elem)):
            if 'Not in Library' in open_items_elem[x].getText():
                continue
            else:
                acc += 1
        if acc > 0:
            # print(title)
            # print(author)
            return 'Found in Open Library.', open_lib_url

        else:
            return 'Not found in Open Library.'


# basic searcher for the different books with spreadsheets
def spread_sheet_searcher(title, author, books, title_col, author_col, db_name):
    acc = 0
    for row in range(1, books.nrows):
        title_unformatted = books.cell_value(row, title_col)
        author_unformatted = books.cell_value(row, author_col)
        # string cleanup for each
        if title.lower() in title_unformatted.lower() and author.lower() in author_unformatted.lower():
            acc += 1
    if acc > 0:
        return "Found in " + db_name
    else:
        return "Not found in " + db_name


###searching michigan fulcrum project
def michigan_searcher(title, author, searchtext):
    regex = re.compile('[^a-zA-Z]')
    title_search = regex.sub('', title)
    title_search = title_search.lower()
    author_search = regex.sub('', author)
    author_search = author_search.lower()
    if title_search in searchtext and author_search in searchtext:
        return "Found in Michigan Press Open Access."
    else:
        return "Not found in Michigan Press Open Access."


#right now only searching ISSNs
def textbook_searcher(tb_issn_col,issn):
    if issn == "":
        return "Not found in Students Stores textbook spreadsheet."
    else:
        for row in range(1,textbooks.nrows):
            if issn == str(textbooks.cell_value(row,tb_issn_col)):
                vital_query = title.replace(" ", "%20")
                vital_url = "https://bookshelf.vitalsource.com/#/search?q=%s" % vital_query
                return "Found in Textbooks", "Could be available here: " + vital_url
            else:
                return "Not found in Students Stores textbook spreadsheet."

#Project Gutenberg searcher
#currently using database GUTINDEX.txt
def pg_searcher(title_input,author_input,file):
    regex = re.compile('[^a-zA-Z]')
    title = regex.sub('', title_input)
    author = regex.sub('', author_input)
    title = title.lower()
    author = author.lower()
    acc = 0
    for line in file:
        line = line.lower()
        line = regex.sub('',line)
        if title in line and author in line:
            acc +=1

    if acc > 0:
        query = title_input + ' ' + author_input
        query = query.replace(' ', '+')
        gutenberg_url = "http://www.gutenberg.org/ebooks/search/?query=%s" % query
        return "Found at Project Gutenberg.", gutenberg_url
    else:
        return "Not found at Project Gutenberg."


###ask whether we want
###to search certain dbs
jstor_cont = input("Search JSTOR? (press y): ")
if jstor_cont == "y":
    jstor_workbook = xlrd.open_workbook('jstor_books.xlsx')
    jstor_books = jstor_workbook.sheet_by_index(0)
    jstorheadings = jstor_books.row_values(0)
    jstortitle_col = jstorheadings.index('Title')
    jstorauthor_col = jstorheadings.index('Authors')

muse_cont = input("Search Project Muse? (press y): ")
if muse_cont == "y":
    muse_workbook = xlrd.open_workbook("project_muse_free_covid_book.xlsx")
    muse_books = muse_workbook.sheet_by_index(0)
    museheadings = muse_books.row_values(0)
    musetitle_col = museheadings.index('Title')
    museauthor_col = museheadings.index('Contributor')

ohio_cont = input("Search Ohio? (press y): ")
if ohio_cont == "y":
    ohio_workbook = xlrd.open_workbook("OhioStateUnivPress-OpenTitles-KnowledgeBank.xlsx")
    ohio_books = ohio_workbook.sheet_by_index(0)
    ohioheadings = ohio_books.row_values(0)
    ohiotitle_col = ohioheadings.index('Title')
    ohioauthor_col = ohioheadings.index('Contributors')

science_direct_cont = input("Search Science Direct? (press y): ")
if science_direct_cont == "y":
    sd_workbook = xlrd.open_workbook("sciencedirect.xlsx")
    sd_books = sd_workbook.sheet_by_index(0)
    sdheadings = sd_books.row_values(0)
    sdtitle_col = sdheadings.index("publication_title")
    sdauthor_col = sdheadings.index("first_author")

#michigan searcher
michigan_cont = input("Search Michigan? (press y): ")
if michigan_cont == "y":
    fulcrum_searchtxt = ""
    regex = re.compile('[^a-zA-Z]')
    for n in range(1,13):
        fulcrum = s.get("https://www.fulcrum.org/michigan?locale=en&page={n}&per_page=1000&view=list".format(n=n))
        fulcrumSoup = bs4.BeautifulSoup(fulcrum.text,'html.parser')
        fulcrumSouptags = fulcrumSoup.select("#documents")
        fulcrumugly = fulcrumSouptags[0].getText()
        fulcrumtext = regex.sub('', fulcrumugly)
        fulcrumtext = fulcrumtext.lower()
        fulcrum_searchtxt = "".join([fulcrum_searchtxt,fulcrumtext])

#vitalsource/textbook searcher
#could be used for general textbook searching
vitalsource_cont = input("Search vitalsource? (press y): ")
if vitalsource_cont == "y":
    issn_col = headings.index('ISSN')
    textbook_workbook = xlrd.open_workbook('Spring 2020 Book List.xlsx')
    textbooks = textbook_workbook.sheet_by_index(0)
    tb_issn_col = 2

#to search project gutenberg
gutenberg_cont = input("Search Project Gutenberg? (press y): ")
if gutenberg_cont == 'y':
    fin = open('GUTINDEX.txt', encoding="utf-8")

###helper functions###
#helper function for appending and printing
#what is returned from functions
def return_helper(result,list):
    if result.__class__.__name__  == 'tuple':
        for item in result:
            print(item)
            list.append(item)
    else:
        print(result)
        list.append(result)

#helper for converting list to textfile
def list_to_file(var_list,name):
    file_name = name + '.txt'
    #open file for writing
    outfile = open(file_name, 'w')
    #write the list to file
    for item in var_list:
        outfile.write(item + '\n')
    outfile.close()
    return (file_name)


#main loop
batch =[]
for row in range(1,books.nrows):
    ####this title search needs to be fixed maybe????????
    title_long = books.cell_value(row,title_col)
    title_split = title_long.split(':')
    title = title_split[0]
    author_num = books.cell_value(row,author_col)
    author = ""
    if len(author_num) > 0:
        for i in author_num:
            if i.isalpha() or i.isspace():
                author = "".join([author, i])
    print(title.upper())
    batch.append(title.upper())
    print(author)
    batch.append(author)
    open_library_result = open_library_access(title, author)
    return_helper(open_library_result, batch)
    red_shelf_result = red_shelf_access(title, author)
    return_helper(red_shelf_result, batch)
    hathi_result = hathi_access(title, author)
    return_helper(hathi_result, batch)
    proquest_result = proquest_access(title, author)
    return_helper(proquest_result, batch)
    cambridge_result = cambridge_access(title, author)
    return_helper(cambridge_result, batch)
    
    if jstor_cont == "y":
        return_helper(spread_sheet_searcher(title, author, jstor_books, jstortitle_col, jstorauthor_col, "JSTOR."),
                      batch)
    if muse_cont == "y":
        return_helper(spread_sheet_searcher(title, author, muse_books, musetitle_col, museauthor_col, "Project Muse."),
                      batch)
    if ohio_cont == 'y':
        return_helper(
            spread_sheet_searcher(title, author, ohio_books, ohiotitle_col, ohioauthor_col, "Ohio State Uni Press."),
            batch)
    if science_direct_cont == 'y':
        return_helper(
            spread_sheet_searcher(title, author, sd_books, sdtitle_col, sdauthor_col, "Science Direct Holdings."),
            batch)
    if michigan_cont == "y":
        return_helper(michigan_searcher(title, author, fulcrum_searchtxt), batch)
    if vitalsource_cont == "y":
        issn = str(books.cell_value(row,issn_col))
        return_helper(textbook_searcher(tb_issn_col,issn),batch)
    if gutenberg_cont == "y":
        return_helper(pg_searcher(title,author,fin),batch)
    print('----------------')
    batch.append('----------------')

timestamp = str(datetime.datetime.now())
timestamp = timestamp.replace(':','_')
timestamp = timestamp.replace('.','_')
list_to_file(batch,'batch imported at ' + timestamp)
