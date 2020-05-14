# Processes a batch of ILLiad requests and checks different online databases
# to see if there's an available copy. Export ILLiad requests into folder
# python file is located. Rename file 'requests.xlsx.'

import requests, bs4,  xlrd, re, datetime
from nltk.corpus import stopwords

s = requests.Session()

#import your illiad requests in this spreadsheet
requests_workbook = xlrd.open_workbook('requests.xlsx')
books = requests_workbook.sheet_by_index(0)

headings = books.row_values(0)
title_col = headings.index('Loan Title')
author_col = headings.index('Loan Author')

#UNC catalog function
#maybe a better update on proquest/cambridge/hathi
def unc_searcher(title,author):
    # helper string
    oncheck = "onlinefulltextavailable"
    #compile regular expression
    regex = re.compile('[^a-zA-Z]')
    #string editing to make query
    title = title.lower()
    title = title.replace(" ", "+")
    author = author.lower()
    author = author.replace(" ", "+")
    #query for unc url
    query = title + "+" + author
    unc_url = "https://catalog.lib.unc.edu/?utf8=%E2%9C%93&search_field=all_fields&q={query}&f%5Baccess_type_f%5D%5B%5D=Online".format(
        query=query)
    unc_get = s.get(unc_url)
    #create bsoup object
    uncSoup = bs4.BeautifulSoup(unc_get.text, 'html.parser')
    #this is one way to select the items on the UNC page
    #using schema.org/Thing seemed like the best way
    unc_items = uncSoup.find_all("div", {"itemtype": "http://schema.org/Thing"})

    title = regex.sub('', title)
    author = regex.sub('', author)
    #dictionary in case multiple links
    found_dic = {"Found at UNC: ": []}
    unc_acc = 0
    for item in unc_items:
        search_soup = item.getText()
        search_clean = regex.sub('', search_soup)
        search_clean = search_clean.lower()
        if title in search_clean and author in search_clean and oncheck in search_clean:
            #for each item find all links
            link_elems = item.find_all("a")
            link_urls = [x.get("href") for x in link_elems]
            for url in link_urls:
                if url.startswith('/'):
                    url = "https://catalog.lib.unc.edu" + url
                found_dic["Found at UNC: "].append(url)
            unc_acc +=1

    if unc_acc > 0:
        return found_dic
    else:
        return "Not found in UNC catalog."


#no longer using proquest searcher
#use unc catalog searcher
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

#no longer using
#use unc catalog searcher
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


def hathi_temp_access(title, author):
    title = title.lower()
    author = author.lower()
    title_unc = title.replace(" ", "+")
    author_no_end_space = author.strip(" ")
    author_unc = author_no_end_space.replace(" ", "+")
    query = author_unc + "+" + title_unc
    hathi_url = "https://catalog.lib.unc.edu/?utf8=%E2%9C%93&search_field=all_fields&q={query}&f%5Baccess_type_f%5D%5B%5D=Online".format(
        query=query)
    hathi_get = s.get(hathi_url)
    hathiSoup = bs4.BeautifulSoup(hathi_get.text, 'html.parser')
    hathi_items = hathiSoup.find_all("div", {"itemtype": "http://schema.org/Thing"})

    regex = re.compile('[^a-zA-Z]')
    title = regex.sub('', title)
    author = regex.sub('', author)
    #string for checking
    oncheck = "temporarilyavailable"
    hathi_dic = {"Temporary access: ": []}
    hathi_acc = 0
    for item in hathi_items:
        search_soup = item.getText()
        search_clean = regex.sub('', search_soup)
        search_clean = search_clean.lower()
        if title in search_clean and author in search_clean and oncheck in search_clean:
            bib_num = item.get("data-document-id")
            hathi_url = "https://catalog.lib.unc.edu/catalog/" + bib_num
            hathi_dic["Temporary access: "].append(hathi_url)
            hathi_acc +=1

    if hathi_acc > 0:
         return hathi_dic
    else:
         return "Not found in HathiTemp Access"

# search actual hathitrust site for always accessible
def hathi_full_time_access(title,author):
    title = title.lower()
    author = author.strip(" ")
    title_query = title.replace(" ","+")
    author_query = author.replace(" ","+")
    query = title_query + "+" + author_query
    ht_url = "https://catalog.hathitrust.org/Search/Home?lookfor={query}&ft=ft&setft=true".format(
        query=query)
    ht_page = s.get(ht_url)
    htSoup = bs4.BeautifulSoup(ht_page.text, 'html.parser')
    records = htSoup.select('.record')

    if len(records) > 0:
        hathi_acc = 0
        # get text for each record
        records_list = [x.getText() for x in records]
        # cleanup using regular expressions
        regex = re.compile('[^a-zA-Z]')
        record_clean_list = []
        for record in records_list:
            record_clean = regex.sub('', record)
            record_clean = record_clean.lower()
            record_clean_list.append(record_clean)
        # clean our search terms
        title_caps = regex.sub('', title)
        title_test = title_caps.lower()
        author_caps = regex.sub('', author)
        author_test = author_caps.lower()

        for record in record_clean_list:
            if title_test in record and author_test in record:
                hathi_acc += 1
        if hathi_acc > 0:
            return "Found in HathiTrust fulltime access.", ht_url
        else:
            return "Not found in HathiTrust fulltime access."

    else:
        return "Not found in HathiTrust fulltime access."


###function for red_shelf_access
###pretty much works
def red_shelf_access(title,author):
    title = title.lower()
    author = author.lower()
    query = title + " " + author
    red_query = query.replace(" ", "+")
    red_shelf_url = 'https://studentresponse.redshelf.com/search/?terms=%s' % red_query
    red_shelf = s.get(red_shelf_url)

    redSoup = bs4.BeautifulSoup(red_shelf.text, 'html.parser')
    red_items = redSoup.select(".price-content")
    if len(red_items) > 0:
        if "Borrow through" in red_items[0].getText():
            title_items = redSoup.select(".title-row")
            # print(title_items[0].getText())
            t1 = title_items[0].getText()
            # remove stopwords
            t1 = t1.lower()
            t1_list = t1.split()
            t1_no_stop = [word for word in t1_list if not word in stopwords.words()]
            t2 = "".join(t1_no_stop)
            #remove stopwords from query
            title_list = title.split()
            title_no_stop = [word for word in title_list if not word in stopwords.words()]
            title_nospace = "".join(title_no_stop)

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
    title = title.lower()
    author = author.lower()
    regex = re.compile('[^a-zA-Z]')
    title = regex.sub('',title)
    author = regex.sub('',author)
    acc = 0
    for row in range(1, books.nrows):
        book_title = books.cell_value(row, title_col)
        book_title = regex.sub('',book_title)
        book_title = book_title.lower()
        author_title = books.cell_value(row, author_col)
        author_title = regex.sub('',author_title)
        author_title = author_title.lower()
        # string cleanup for each
        if title in book_title and author in author_title:
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
#jstor_cont = input("Search JSTOR? (press y): ")
jstor_cont = "y"
if jstor_cont == "y":
    print("loading jstor books")
    jstor_workbook = xlrd.open_workbook('jstor_books.xlsx')
    jstor_books = jstor_workbook.sheet_by_index(0)
    jstorheadings = jstor_books.row_values(0)
    jstortitle_col = jstorheadings.index('Title')
    jstorauthor_col = jstorheadings.index('Authors')

#muse_cont = input("Search Project Muse? (press y): ")
muse_cont = "y"
if muse_cont == "y":
    print("loading muse books")
    muse_workbook = xlrd.open_workbook("project_muse_free_covid_book.xlsx")
    muse_books = muse_workbook.sheet_by_index(0)
    museheadings = muse_books.row_values(0)
    musetitle_col = museheadings.index('Title')
    museauthor_col = museheadings.index('Contributor')

#ohio_cont = input("Search Ohio? (press y): ")
ohio_cont = "y"
if ohio_cont == "y":
    print("loading ohio books")
    ohio_workbook = xlrd.open_workbook("OhioStateUnivPress-OpenTitles-KnowledgeBank.xlsx")
    ohio_books = ohio_workbook.sheet_by_index(0)
    ohioheadings = ohio_books.row_values(0)
    ohiotitle_col = ohioheadings.index('Title')
    ohioauthor_col = ohioheadings.index('Contributors')

#science_direct_cont = input("Search Science Direct? (press y): ")
science_direct_cont = "y"
if science_direct_cont == "y":
    print("loading science direct books")
    sd_workbook = xlrd.open_workbook("sciencedirect.xlsx")
    sd_books = sd_workbook.sheet_by_index(0)
    sdheadings = sd_books.row_values(0)
    sdtitle_col = sdheadings.index("publication_title")
    sdauthor_col = sdheadings.index("first_author")

#michigan searcher
#michigan_cont = input("Search Michigan? (press y): ")
michigan_cont = "y"
if michigan_cont == "y":
    print("loading michigan press books")
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
#vitalsource_cont = input("Search vitalsource? (press y): ")
vitalsource_cont = "y"
if vitalsource_cont == "y":
    print("loading textbook list")
    issn_col = headings.index('ISSN')
    textbook_workbook = xlrd.open_workbook('Spring 2020 Book List.xlsx')
    textbooks = textbook_workbook.sheet_by_index(0)
    tb_issn_col = 2

#to search project gutenberg
#gutenberg_cont = input("Search Project Gutenberg? (press y): ")
gutenberg_cont = "y"
if gutenberg_cont == 'y':
    print("loading gutenberg books")
    fin = open('GUTINDEX.txt', encoding="utf-8")

###helper functions###
#helper function for appending and printing
#what is returned from functions
def return_helper(result,list):
    if result.__class__.__name__  == 'tuple':
        for item in result:
            print(item)
            list.append(item)
    elif result.__class__.__name__ == 'dict':
        for key in result:
            print(key)
            list.append(key)
            for elem in result[key]:
                print(elem)
                list.append(elem)
    else:
        print(result)
        list.append(result)

#helper for converting list to textfile
def list_to_file(var_list,name):
    file_name = name + '.txt'
    #open file for writing
    outfile = open(file_name, 'w',encoding='utf-8')
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

    unc_result = unc_searcher(title,author)
    return_helper(unc_result,batch)
    open_library_result = open_library_access(title, author)
    return_helper(open_library_result, batch)
    red_shelf_result = red_shelf_access(title,author)
    return_helper(red_shelf_result, batch)
    hathi_temp_result = hathi_temp_access(title, author)
    return_helper(hathi_temp_result, batch)
    hathi_full_result = hathi_full_time_access(title,author)
    return_helper(hathi_full_result,batch)
    #proquest_result = proquest_access(title, author)
    #return_helper(proquest_result, batch)
    #cambridge_result = cambridge_access(title, author)
    #return_helper(cambridge_result, batch)
    
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
