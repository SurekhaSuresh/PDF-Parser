import io
import os
import datetime
from pdfminer.converter import HTMLConverter, XMLConverter, PDFPageAggregator
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import re
import pytesseract
from pdf2image import convert_from_path
import PyPDF2
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import HTMLConverter
from pdfminer.layout import LAParams, LTTextLineHorizontal, LTTextLineVertical, LTTextLine, LTTextBoxHorizontal
from pdfminer.pdfdocument import PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFSyntaxError



def pdftohtml(pdf_file):
    current_time = currentTime()
    output_file_path = f"processed_{current_time}.html"
    rsrcmgr = PDFResourceManager()
    codec = 'utf-8'
    laparams = LAParams()

    with io.BytesIO() as output_stream:
        converter = HTMLConverter(rsrcmgr, output_stream, codec=codec, laparams=laparams)
        with open(pdf_file, 'rb') as input_file:
            for page in PDFPage.get_pages(input_file):
                try:
                    interpreter = PDFPageInterpreter(rsrcmgr, converter)
                    interpreter.process_page(page)
                except PDFSyntaxError:
                    pass

            html_content = output_stream.getvalue().decode()
            with open(output_file_path, 'w') as output_file:
                output_file.write(html_content)

    return output_file_path


def getPDF():
    for i in range(3):
        input_file_path = input("Enter the PDF file path: ")
        if os.path.isfile(input_file_path):
            return input_file_path
        else:
            print("File not found.")
    print("Maximum number of attempts reached. Exiting program.")
    exit()

def isSearchable():
    for i in range(3):
        isSearchable = input("Is the given PDF searchable? (Enter y or n): ")
        if isSearchable == 'y':
            return True
        elif isSearchable == 'n':
            return False
        else:
            print("Invalid Entry!!")
    print("Maximum number of attempts reached. Exiting program.")
    exit()

def currentTime():
    return datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")

def fix_unordered_lists(html):
    # Match any <span> tag that contains bullet characters
    pattern = r"<span[^>]*>(?:•|&#8226;|&bull;|\u2022|-|–|&ndash;|&diamond;|\n)\s*([^<]*)</span>"

    # Replace matched spans with bullet tags
    html = re.sub(pattern, r"<bullet>\1</bullet>", html, flags=re.MULTILINE)

    list_items = []
    current_list_item = ''

    # Find all occurrences of the </bullet> tag
    bullet_closing_indices = [m.end() for m in re.finditer('</bullet>', html)]

    # Iterate through the bullet_closing_indices and get the text between bullet tags
    for i, closing_index in enumerate(bullet_closing_indices):
        opening_index = html.find('<bullet>', closing_index)
        if opening_index == -1:
            opening_index = len(html)
        text = html[closing_index:opening_index]
        text = text.strip()
        p = re.compile(r'<span[^>]*>|<\/span>')
        text = p.sub('', text)
        if not text or text.isspace():
            continue
        text = f"<li>{text}</li>"
        if i == 0:
            # If this is the first item, start a new list
            list_items.append([text])
        elif opening_index == len(html):
            # If this is the last item and there is no opening <bullet> tag, add it to the current list
            current_list_item += text
            list_items[-1].append(current_list_item)
        elif opening_index > bullet_closing_indices[i - 1] + 1:
            # If there is a gap between this bullet and the previous one, start a new list
            list_items.append([text])
            current_list_item = ''
        else:
            # Otherwise, add it to the current list
            current_list_item += text
            list_items[-1].append(current_list_item)

    if len(list_items) == 0:
        result = ""
    else:
        result = '<ul>\n' + '\n'.join([''.join(l) for l in list_items]) + '\n</ul>'
    firstBulletTagIndex = html.find('<bullet>', 0)
    openingText = html.split('<bullet>', 1)[0]
    result = openingText + result

    # Replace <span> tags with <p> tags
    result = re.sub(r'<span\s+[^>]*>(.*?)<\/span>', r'<p>\1</p>', result, flags=re.DOTALL)

    return result

def fix_ordered_lists(html):
    inputHTML = html
    # Match ordered list indices
    pattern = r"(?:<p>|<br\/>)(\(\d+\)|\d+\)|\d+\.)"

    # Replace matched pattern with bullet tags
    html = re.sub(pattern, r"<bullet>\1", html)

    pattern = r"<bullet>\s*(.*?)\s*(?=<bullet>|\Z)(?:</p>)?"
    matches = re.findall(pattern, html, re.DOTALL)

    # Wrap each match inside <li> tags
    list_items = []
    prev_num = None
    for match in matches:
        num_match = re.match(r"^\(?(\d+)\)?(\(\d+\)|\d+\)|\d+\.)*", match.strip())
        if num_match:
            curr_num = int(num_match.group(1))
            if prev_num is None or curr_num == prev_num + 1:
                if prev_num is None:
                    list_items.append("<ol>")
                list_items.append(f"<li>{match.strip()}</li>")
            else:
                list_items.append("</ol>")
                list_items.append("<ol>")
                list_items.append(f"<li>{match.strip()}</li>")
            prev_num = curr_num

    if list_items:
        list_items.append("</ol>")

    # Join the list items into a single string
    list_html = "\n".join(list_items)

    pattern = r'</p>(?!(\s*<p>|$))'

    list_html = re.sub(pattern, '', list_html)
    if not list_html:
        return inputHTML
    return list_html

def htmltocsv(html_file_name):
    with open(html_file_name) as f:
        html_text = f.read()

    soup = BeautifulSoup(html_text, 'html.parser')

    stack = []
    headings = []
    hCounter = 0
    shCounter = 0
    latest_type = -1 # Title - 0, Heading - 1, Subheading - 2, Content - 3
    titleParsed = False

    for div in soup.find_all(lambda div: div.name == 'div' and 'left' in div.get('style').lower()):
        # Fetching value of left attribute
        style = div.get('style').lower()
        left_value = None
        for style_prop in style.split(';'):
            if 'left' in style_prop:
                left_value = int(style_prop.split(':')[1].strip('px'))
                break

        if left_value is not None:
            span_tags = div.find_all('span')
            for tag in span_tags:
                # Fetching value of font-size attribute
                sty = tag.get('style').lower()
                font_size = None
                for style_prop in sty.split(';'):
                    if 'font-size' in style_prop:
                        font_size = int(style_prop.split(':')[1].strip('px'))
                        break

                content = tag.get_text().strip()
                raw = str(tag)

                if tag in soup.find_all(lambda tag: tag.name == 'span' and '-bold' in tag.get('style').lower() and font_size is not None and font_size >= 14) and not titleParsed:
                    # Title
                    headings.append({'data': content, 'type': 'title'})
                    latest_type = 0
                    titleParsed = True

                elif tag in soup.find_all(lambda tag: tag.name == 'span' and '-bold' in tag.get('style').lower()):

                    if latest_type == 0:
                        # Heading
                        headings.append({'data': content, 'type': 'heading'})
                        latest_type = 1
                        stack.append({'H'+str(hCounter) : left_value})
                        hCounter+= 1

                    elif latest_type == 1:
                        # Subheading
                        headings.append({'data': content, 'type': 'subheading'})
                        latest_type = 2
                        stack.append({'SH'+str(shCounter) : left_value})
                        shCounter+= 1

                    elif latest_type == 3:
                        isParsed = False
                        while len(stack) != 0:
                            top = stack[-1]
                            key = list(top.keys())[0]
                            val = top[key]

                            if left_value < val:
                                stack.pop()

                            elif left_value == val:
                                if key.startswith('H'):
                                    stack.append({'H'+str(hCounter) : left_value})
                                    hCounter+= 1
                                    headings.append({'data': content, 'type': 'heading'})
                                    latest_type = 1
                                    isParsed = True
                                elif key.startswith('SH'):
                                    stack.append({'SH'+str(shCounter) : left_value})
                                    shCounter+= 1
                                    headings.append({'data': content, 'type': 'subheading'})
                                    latest_type = 2
                                    isParsed = True
                                break

                            elif left_value > val:
                                if key.startswith('H'):
                                    stack.append({'SH'+str(shCounter) : left_value})
                                    shCounter+= 1
                                    headings.append({'data': content, 'type': 'subheading'})
                                    latest_type = 2
                                    isParsed = True
                                elif key.startswith('SH'):
                                    if latest_type == 3:
                                        latest_dict = headings[-1]
                                        latest_dict['data'] += raw
                                        headings[-1] = latest_dict
                                    else:
                                        headings.append({'data': raw, 'type': 'content'})
                                    latest_type = 3
                                    isParsed = True
                                break

                        if len(stack) == 0 and isParsed == False:
                            stack.append({'H'+str(hCounter) : left_value})
                            hCounter+= 1
                            headings.append({'data': content, 'type': 'heading'})
                            latest_type = 1

                elif tag not in soup.find_all(lambda tag: tag.name == 'span' and '-bold' in tag.get('style').lower()):
                    # Content
                    if latest_type == 3:
                        latest_dict = headings[-1]
                        latest_dict['data'] += raw
                        headings[-1] = latest_dict
                    else:
                        headings.append({'data': raw, 'type': 'content'})
                    latest_type = 3

    # Create a DataFrame
    df = pd.DataFrame(headings)

    # Add a new column with the title text
    df['Topic'] = df['data'].where(df['type'] == 'title', '').ffill()

    # Add a new column with the heading text
    df['Categ'] = df['data'].where(df['type'] == 'heading', '').ffill()

    # Add a new column with the subheading text
    df['Sub_cat'] = df['data'].where(df['type'] == 'subheading', '').ffill()

    # Add a new column with the content text
    df['Text'] = df['data'].where(df['type'] == 'content', '')


    # Drop 'data' column
    df.drop('data', axis=1, inplace=True)
    # Drop 'type' column
    df.drop('type', axis=1, inplace=True)

    df = df.replace(r'^\s*$',np.nan,regex=True)
    df['Topic'].fillna(method="ffill",inplace=True)
    df['Categ'].fillna(method="ffill",inplace=True)
    df['Sub_cat'] = df.groupby('Categ', sort=False)['Sub_cat'].apply(lambda x: x.ffill())
    current_time = currentTime()
    df['Crawl_datetime'] = current_time

    df = df[df['Text'].notna()]

    df = df[["Crawl_datetime","Topic","Categ","Sub_cat","Text"]]



    for i, row in df.iterrows():
        processed_text = fix_unordered_lists(str(row['Text']))
        df.loc[i, 'Processed_Text'] = processed_text

    df.drop('Text', axis=1, inplace=True)
    df.rename(columns={'Processed_Text': 'Text'}, inplace=True)
    df = df.replace(r'^\s*$',np.nan,regex=True)
    df = df[df['Text'].notna()]

    for i, row in df.iterrows():
        processed_text = fix_ordered_lists(str(row['Text']))
        df.loc[i, 'Processed_Text'] = processed_text

    df.drop('Text', axis=1, inplace=True)
    df.rename(columns={'Processed_Text': 'Text'}, inplace=True)

    csv_file_name = f"parser_{current_time}.csv"
    # Save the DataFrame to a CSV file
    df.to_csv(csv_file_name, index=False)
    print("The CSV file path is:",csv_file_name)

def getSearchable(pdf_file_path):
    # Convert each page of the PDF file to an image
    images = convert_from_path(pdf_file_path)

    # Create a new PDF file
    output_pdf = PyPDF2.PdfWriter()

    # Loop through each image and perform OCR using PyTesseract
    for i, image in enumerate(images):
        # Get a searchable PDF
        pdf = pytesseract.image_to_pdf_or_hocr(image, extension='pdf')
        # Add the page to the output PDF file
        output_page = PyPDF2.PdfReader(io.BytesIO(pdf)).pages[0]
        output_pdf.add_page(output_page)

    current_time = currentTime()
    pdf_file_name = f"searchable_{current_time}.pdf"
    with open(pdf_file_name, 'wb') as f:
        output_pdf.write(f)
    return pdf_file_name

def parser():
    pdf_file_path = getPDF()
    print("The PDF file path is:", pdf_file_path)
    if not isSearchable():
        pdf_file_path = getSearchable(pdf_file_path)
    processedHTML = pdftohtml(pdf_file_path)
    htmltocsv(processedHTML)


parser()
