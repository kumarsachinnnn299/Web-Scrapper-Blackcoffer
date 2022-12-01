#-------------------------------> Importing important libraries <----------------------------------------

from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
import re
import os.path
import string
import requests
import csv
import pandas as pd
from bs4 import BeautifulSoup as bs
headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0',
}



#---------------------------------> Loading the input excel file <----------------------------------------
df = pd.read_excel('Input.xlsx')

# Defining directory for text files
text_file_path = 'Text Files'



#----------------------------------> Defining Output csv file <--------------------------------------------
csv_file = open('Output.csv', 'w', newline='')
csv_writer = csv.writer(csv_file)
csv_writer.writerow(['URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
                     'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE',
                     'COMPLEX WORD COUNT', 'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH', ' '])



# -------------------------------> Collecting required data for analysis function <------------------------------------------

stopping_words = []#Contains all stopping words
final_words = [] #Contains words after removing all stopping words 
Master_positive_words = []
Master_negative_words = []

# Loading stopping words
for filename in os.listdir('StopWords'):
    f = os.path.join('StopWords', filename)
    with open(f, 'r') as file:
        for line in file:
            for word in line.split():
                cleaned_word = re.sub(r'[^\w\s]', '', word)  #To remove unnecessary  punctuations from text
                stopping_words.append(cleaned_word)


stopping_words = [ele for ele in stopping_words if ele.strip()]  # to remove empty spaces from list of strings


# Loading Positive and negative words
# Master Postive words
with open('MasterDictionary/positive-words.txt', 'r') as file:
    for line in file:
        for word in line.split():
            Master_positive_words.append(word)

Master_positive_words = [ele for ele in Master_positive_words if ele.strip()]# To remove empty spaces from list

# Master negative words
with open('MasterDictionary/negative-words.txt', 'r') as file:
    for line in file:
        for word in line.split():
            Master_negative_words.append(word)# To remove empty spaces from list

Master_negative_words = [ele for ele in Master_negative_words if ele.strip()]




#--------------------------------> Main Analysis Function <------------------------------------------
def analyse(name, url, content):
    # Firstly tokenizing the content
    cleaned_text = re.sub(r'[^\w\s]', '', content)#To remove unnecessary punctuations from text
    tokenized_sentence = sent_tokenize(content)
    tokenized_words = word_tokenize(cleaned_text)

    # Removing the StopWords
    for word in tokenized_words:
        if word not in stopping_words:
            final_words.append(word)

    # Calculating Positive Score
    positive_score = 0
    for word in final_words:
        if word in Master_positive_words:
            positive_score += 1

    # Calculating Negative Score
    negative_score = 0
    for word in final_words:
        if word in Master_negative_words:
            negative_score += 1

    # Calculating Polarity Score
    polarity_score = (positive_score-negative_score) / \
        ((positive_score + negative_score) + 0.000001)

    # Calculating Subjectivity Score
    total_words_after_cleaning = len(final_words)
    subjectivity_score = (positive_score+negative_score) / \
        ((total_words_after_cleaning) + 0.000001)

    # Calculating Average Sentence Length
    avg_sentence_len = len(final_words)/len(tokenized_sentence)

    # Calculating percentage of complex words
    complex_words = 0
    for word in final_words:
        if word.endswith('es') or word.endswith('ed'):
            continue
        else:
            syllable = 0
            for char in word:
                if char == 'a' or char == 'e' or char == 'i' or char == 'o' or char == 'u':
                    syllable += 1
                if(syllable > 2):
                    complex_words += 1
                    break

    percent_complex_words = len(final_words)/complex_words

    # Calculating Fog Index
    fog_index = 0.4*(avg_sentence_len + percent_complex_words)

    # Calculating Average Number of Words Per Sentence
    avg_words_per_sentence = len(final_words)/len(tokenized_sentence)

    # Calculating Syllable per word
    syllable_count = 0
    for word in final_words:
        if word.endswith('es') or word.endswith('ed'):
            continue
        else:
            for char in word:
                if char == 'a' or char == 'e' or char == 'i' or char == 'o' or char == 'u':
                    syllable_count += 1
    syllable_per_word = syllable_count/len(final_words)

    # Calculating Personal Pronouns
    pronounRegex = re.compile(r'\b(I|we|my|ours|(?-i:us))\b', re.I)
    pronouns_present = pronounRegex.findall(cleaned_text)
    total_pronouns = len(pronouns_present)

    # Calculating Average Word Length
    total_char = 0
    for word in final_words:
        total_char += len(word)

    avg_word_length = total_char/len(final_words)
    csv_writer.writerow([name, url, positive_score, negative_score, polarity_score, subjectivity_score,
                         avg_sentence_len, percent_complex_words, fog_index, avg_words_per_sentence, complex_words, len(
                             final_words),
                         syllable_per_word, total_pronouns, avg_word_length])
    print(f'file: {filename} Completed.... ')





# --------------------------------------->Scrapping Part <-----------------------------------

# Scrapping Links
for index in range(0, df.shape[0], 1):
    filename = int(df.iloc[index, 0])
    source = requests.get(df.iloc[index, 1], headers=headers).text
    data = bs(source, "html.parser")
    try:
        # Getting Title of Article
        title = data.find("h1", class_="entry-title").text

        # Getting Content of article
        body = data.find('div', class_='td-post-content').text
    except Exception as x:
        csv_writer.writerow([filename, df.iloc[index, 1], '            NA', '            NA', '            NA',
                             '            NA', '            NA', '            NA', '            NA', '            NA',
                             '            NA', '            NA', '            NA', '            NA', '            NA'])
        print(f'file: {filename} Completed.... ')
        continue
    try:
        # Removing footer from the content of article
        footer = data.find('div', class_="td-post-content").pre.text
        body = body.replace(footer, '')
    except Exception as x:
        footer = ''

    # For Creating text files

    # To create directory for text files if not present
    if not os.path.isdir(text_file_path):
        os.makedirs(text_file_path)

    completename = os.path.join(text_file_path, f'{filename}.txt')
    fh = open(completename, "w", encoding="utf-8")
    fh.write('Title: '+title+'\n')
    fh.write(body)
    fh.close()

    # Calling Analysis function

    analyse(filename, df.iloc[index, 1], body)

csv_file.close()
