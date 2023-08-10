import requests
from bs4 import BeautifulSoup
import openpyxl


res = requests.get('https://news.ycombinator.com/')

soup = BeautifulSoup(res.text, 'html.parser')

titles = soup.select('.titleline > a')
subtext = soup.select('.subtext')


def sort_stories_by_votes(hnlist):
    return sorted(hnlist, key=lambda k: k['votes'], reverse=True)


def create_custom_hacker_news(links, subtext):

    arr = []
    # Create a new workbook
    wb = openpyxl.Workbook()

    # Add a sheet to the workbook
    sheet = wb.create_sheet("Sheet")

    sheet['A1'] = "Title"
    sheet['B1'] = "Link"
    sheet['C1'] = "Votes"
    for idx, link_tag in enumerate(links):
        title = links[idx].getText()
        href = link_tag.get('href', None)
        vote = subtext[idx].select('.score')

        if len(vote):
            points = int(vote[0].getText().replace(' points', ''))
            if points > 99:
                arr.append({'title': title, 'link': href, 'votes': points})

                # create a new row in the sheet

    for idx, item in enumerate(arr):
        sheet[f"A{idx+2}"] = item['title']
        sheet[f"B{idx+2}"] = item['link']
        sheet[f"C{idx+2}"] = item['votes']

    # Save the workbook
    wb.save("news.xlsx")


create_custom_hacker_news(titles, subtext)
