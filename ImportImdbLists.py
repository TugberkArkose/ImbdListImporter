from HTMLParser import HTMLParser
from xlsxwriter import Workbook
import urllib2

#
# Xlswriter write_url doesn't writes urls properly.
# Uses \ instead of / as a file directory(url) separator.
# Possible bug within the library !!!
#


# Create a subclass and override the handler methods
class MyHTMLParser(HTMLParser):

    # Open the excel sheet
    workbook = Workbook(local_directory)
    worksheet = workbook.add_worksheet()

    # Widen the second column to fit movie name
    worksheet.set_column('B:B', 35)

    # Title format
    title_format = workbook.add_format({'bold': True})
    title_format.set_bg_color("gray")

    # Url format
    url_format = workbook.add_format({
        'font_color': 'blue',
        'underline': 1
    })

    # Titles
    worksheet.write(0, 0, '#', title_format)
    worksheet.write(0, 1, 'Name', title_format)
    worksheet.write(0, 2, 'Year', title_format)
    worksheet.write(0, 3, 'Rating', title_format)
    #

    previousStartTag = ''
    attribute = []
    count = 1

    def handle_starttag(self, tag, attrs):
        self.previousStartTag = tag
        self.attribute = attrs

    def handle_data(self, data):

        # Name, link and the count
        if self.previousStartTag == 'a' and self.attribute and "/title/" in self.attribute[0][1] and "\n" not in data:
            self.worksheet.write(self.count, 0, self.count)
            url = "www.imdb.com" + self.attribute[0][1]
            url = url.replace("\\", "\/")
            self.worksheet.write_url(self.count, 1, url, self.url_format, data)

        # Year
        elif self.previousStartTag == 'span' and self.attribute and 'year_type' in self.attribute[0][
            1] and "\n" not in data:
            data = data.replace("(", "")
            data = data.replace(")", "")
            self.worksheet.write(self.count, 2, data)

        # Rating
        elif self.previousStartTag == 'span' and self.attribute and 'value' in self.attribute[0][
            1] and "\n" not in data:
            self.worksheet.write(self.count, 3, data + "/10")
            self.count += 1


# Config variables
url_to_list = 'http://www.imdb.com/list/ls077686615/'
local_directory = 'C:\\Users\\bscuser\\movies.xlsx'
#

# Get raw html from the imdb
response = urllib2.urlopen(url_to_list)
html = response.read()
#

# Create an instance from the Parser
parser = MyHTMLParser()
parser.feed(html)
parser.workbook.close()
#