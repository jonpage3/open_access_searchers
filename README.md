# open_access_searchers
Dependences: requests, bs4, xlrd, re
Use pip install to download dependencies. 

This a program that searches for items in a set of open access and temporary access databases. 
UNC-Chapel Hill focused.

Setup:
To perform searches of the following databases download these files into same folder as program:

jstor_books.xlsx (https://www.jstor.org/open/)
project_muse_free_covid_book.xlsx (https://about.muse.jhu.edu/resources/freeresourcescovid19/#freepublishers)
OhioStateUnivPress-OpenTitles-KnowledgeBank.xlsx (https://kb.osu.edu/handle/1811/131)
sciencedirect.xlsx (https://guides.lib.unc.edu/go.php?c=23609201)
Spring 2020 Book List.xlsx
GUTINDEX.txt (http://www.gutenberg.org/dirs/GUTINDEX.ALL)

Run program:
Run online_request_searcher:
When first starting program enter 'y' for the static databases you want to search. This will load those spreadsheets listed above
and will be used for remainder of time the program is running. 

Program will ask for title and author. Case-insensitive. Doesn't matter if you only have the title or only have author. 
Program will ask for ISBN if you're searching the textbook database. 

When program finishes search it will generate text file <title>_result.txt. 

Enter 'y' to run another search. 

Work in progress. Check here for updates.
