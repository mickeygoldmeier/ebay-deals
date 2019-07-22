from bs4 import BeautifulSoup
import requests
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font


def printh2anddiv(title, div):
    items_list = [title.span.text]

    for item in div.find_all("div", class_="dne-itemtile-detail"):
        item_name = item.h3.text

        try:
            item_price = item.find("div", class_="dne-itemtile-price").span.text
        except:
            item_price = None
        try:
            item_old_price = item.find(
                "div", class_="dne-itemtile-original-price"
            ).span.span.text
        except:
            item_old_price = None

        items_list.append((item_name, item_price, item_old_price))
    return items_list


source = requests.get("https://www.ebay.com").text
soup = BeautifulSoup(source, "lxml")
d_list = soup.find("ul", class_="hl-popular-destinations-elements")

book = Workbook()
sheet = book.active

all_items = []
for item in d_list.find_all("li", class_="hl-popular-destinations-element"):
    if item.h3.text == "Deals":
        link = "https://www.ebay.com/{deal}".format(deal=item.a.text)
        get = requests.get(link).text
        soup = BeautifulSoup(get, "lxml")
        main = soup.find("main")
        sections = main.find("div", class_="sections-container")

        for div in sections.children:
            if div.name == "div":
                for h2 in div.find_all("h2"):
                    if h2.next_sibling == None:
                        if h2.parent.next_sibling.name == "div":
                            all_items.append(printh2anddiv(h2, h2.parent.next_sibling))

                    if h2.next_sibling != None:
                        if h2.next_sibling.name == "div":
                            all_items.append(printh2anddiv(h2, h2.next_sibling))
                        else:
                            all_items.append(
                                printh2anddiv(
                                    h2, h2.next_sibling.next_sibling.next_sibling
                                )
                            )

        for item in all_items:
            sheet = book.create_sheet("{name}".format(name=item[0]))
            sheet.append(("item_name", "new_price", "old_price"))
            sheet["A1"].font = Font(bold=True)
            sheet["B1"].font = Font(bold=True)
            sheet["C1"].font = Font(bold=True)
            for name in item[1:]:
                sheet.append(name)
            book.save("deals.xlsx")
