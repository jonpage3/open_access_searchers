import requests, bs4

s = requests.Session()
from nltk.corpus import stopwords

def red_shelf_access(title,author):
    title = title.lower()
    author = author.lower()
    query = title + " " + author
    red_query = query.replace(" ", "+")
    red_shelf_url = 'https://studentresponse.redshelf.com/search/?terms=%s' % red_query
    red_shelf = s.get(red_shelf_url)

    redSoup = bs4.BeautifulSoup(red_shelf.text, 'html.parser')
    red_items = redSoup.select(".price-content")
    print(red_items)
    if len(red_items) > 0:
        if "Borrow through" in red_items[0].getText():
            title_items = redSoup.select(".title-row")
            # print(title_items[0].getText())
            t1 = title_items[0].getText()
            t1 = t1.lower()
            t1_list = t1.split()
            t1_no_stop = [word for word in t1_list if not word in stopwords.words()]
            print(t1_no_stop)
            t2 = "".join(t1_no_stop)

            print(t2)

            #remove stopwords from title
            title_list = title.split()
            title_no_stop = [word for word in title_list if not word in stopwords.words()]
            title_nospace = "".join(title_no_stop)
            print(title_nospace)
            if title_nospace in t2:

                return "Found in Red Shelf.", red_shelf_url
            else:
                return "Not found in Red Shelf."
        else:
            return "Not found in Red Shelf"
    else:
        return "Not found in Red Shelf"

if __name__ == "__main__":
    cont = "y"
    while cont == "y":
        title = input("enter title: ")
        author = input("enter author: ")
        returns = red_shelf_access(title)
        print(returns)
        cont = input("search again?: ")



