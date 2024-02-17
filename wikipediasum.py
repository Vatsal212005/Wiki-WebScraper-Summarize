import os
import requests
from bs4 import BeautifulSoup
from tkinter import *
from tkinter import messagebox
import tkinter.font as font
from docx import Document
import re
import xml.etree.ElementTree as ET

def scrapeData():
    search_query = searchWord.get()
    url = f"https://en.wikipedia.org/wiki/{search_query}"

    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Extracting data from the Wikipedia page
        paragraphs = soup.find_all('p')
        text = '\n\n'.join(p.text for p in paragraphs)

        # Get the keyword from the user input
        keyword = keywordEntry.get()

        # Save the original scraped data
        saveScrapedData(search_query, text)

        # Extract details related to the keyword from the scraped data
        keyword_data = extractKeywordData(text, keyword)

        # Saving the keyword data to a Word file
        saveKeywordData(keyword, keyword_data)

        messagebox.showinfo("Wikipedia Scraper", f"Data related to '{keyword}' saved successfully")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def saveScrapedData(search_query, text):
    # Save the original scraped data to an XML file
    path = 'C:\\Users\\ironm\\Desktop\\datascrape'
    filename = f"{search_query.replace(' ', '_')}_Original_Data.docx"
    full_path = os.path.join(path, filename)
    root = ET.Element("data")
    doc = ET.SubElement(root, "paragraphs")
    doc.text = text
    tree = ET.ElementTree(root)
    tree.write(full_path)

def extractKeywordData(text, keyword):
    # Search for the keyword in the text
    keyword_data = re.findall(fr'\b{re.escape(keyword)}\b.*?(?=\.\s|$)', text, re.IGNORECASE)

    # Remove patterns like ".[137][135][138][139]"
    keyword_data = [re.sub(r'\.\[\d+\]\[\d+\]\[\d+\]\[\d+\]', '', entry) for entry in keyword_data]

    # Join the extracted data into a single string
    keyword_data_str = ' '.join(keyword_data)
    
    return keyword_data_str

def saveKeywordData(keyword, keyword_data):

    path = 'C:\\Users\\ironm\\Desktop\\datascrape'
    filename = f"{keyword.replace(' ', '_')}_Keyword_Data.docx"
    full_path = os.path.join(path, filename)
    document = Document()
    document.add_paragraph(keyword_data)
    document.save(full_path)

# GUI setup
wn = Tk()
wn.geometry("500x350")
wn.configure(bg='azure2')
wn.title("Vatsal's Wikipedia Scraper")

searchWord = StringVar()
keyword = StringVar()

headingFrame1 = Frame(wn, bg="gray91", bd=5)
headingFrame1.place(relx=0.05, rely=0.1, relwidth=0.9, relheight=0.16)

headingLabel = Label(headingFrame1, text=" Welcome to Vatsal's Wikipedia Scraper", fg='grey19', font=('Courier', 12, 'bold'))
headingLabel.place(relx=0, rely=0, relwidth=1, relheight=1)

Label(wn, text='Enter the topic to be searched on Wikipedia', bg='azure2', font=('Courier', 10)).place(x=20, y=150)
Entry(wn, textvariable=searchWord, width=35, font=('calibre', 10, 'normal')).place(x=20, y=170)

Label(wn, text='Enter the keyword to extract data:', bg='azure2', font=('Courier', 10)).place(x=20, y=200)
keywordEntry = Entry(wn, textvariable=keyword, width=35, font=('calibre', 10, 'normal'))
keywordEntry.place(x=20, y=220)

ScrapeBtn = Button(wn, text='Scrape', bg='honeydew2', fg='black', width=15, height=1, command=scrapeData)
ScrapeBtn['font'] = font.Font(size=12)
ScrapeBtn.place(x=15, y=270)

QuitBtn = Button(wn, text='Exit', bg='old lace', fg='black', width=15, height=1, command=wn.destroy)
QuitBtn['font'] = font.Font(size=12)
QuitBtn.place(x=345, y=270)

wn.mainloop()
