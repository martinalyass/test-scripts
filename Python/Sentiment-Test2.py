import string

from textblob import TextBlob
from textblob import Word

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

from openpyxl import load_workbook
from numpy import median

def write_sentiment(path,master_list):
    
    reader = pd.read_excel(path, header=None, sheet_name="Analysis")
    writer = pd.ExcelWriter(path,engine = 'openpyxl')

    book = load_workbook(path)
    writer.book = book

    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    reader.to_excel(writer,sheet_name="Analysis",index=False,header=False)

    for idx,lst in enumerate(master_list):
        df = pd.DataFrame([lst])
        df.to_excel(writer,sheet_name="Analysis",index=False,header=False,startrow=idx+1,startcol=1)

    writer.save()
    writer.close()

def sentiment_analysis(reviews, keywords, path):

    master_list = []
    for keyword in keywords: 
        filtered_reviews = filter(lambda sentence: keyword in sentence, reviews)
        polarity = 0
        subjectivity = 0
        polarity_list = []
        subjectivity_list = []
        for sentence in filtered_reviews:
            polarity += sentence.sentiment.polarity
            polarity_list.append(sentence.sentiment.polarity)
            subjectivity += sentence.sentiment.subjectivity
            subjectivity_list.append(sentence.sentiment.subjectivity)
        master_list.append([polarity/len(filtered_reviews),median(polarity_list),subjectivity/len(filtered_reviews),median(subjectivity_list),len(filtered_reviews)])
    
    return master_list

def extract_reviews(review_sheet):
    new_sentences = []
    for index,row in review_sheet.iterrows():
        review = filter(lambda x: x in string.printable, (row['text'])).lower()
        sentences = TextBlob(review).sentences
        for sentence in sentences:
            new_sentences.append(sentence)
    return new_sentences

def extract_keywords(keyword_sheet):
    return keyword_sheet["Keywords"].tolist()

if __name__ == "__main__":
    path = "sentiment_analysis.xlsx"
    xl = pd.ExcelFile(path)

    review_sheet = xl.parse('Reviews')
    keyword_sheet = xl.parse('Analysis')

    reviews = extract_reviews(review_sheet)
    keywords = extract_keywords(keyword_sheet)

    master_list = sentiment_analysis(reviews, keywords,path)
    write_sentiment(path,master_list)


