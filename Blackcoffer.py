import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from textblob import TextBlob
import os

# Load the input Excel file
input_file = r'C:\Users\shail\Downloads\Input.xlsx'
df = pd.read_excel(input_file)

# Load the output structure
output_file = r'C:\Users\shail\Downloads\Output Data Structure.xlsx'
output_df = pd.read_excel(output_file)

# Define functions for text analysis
def positive_score(text):
    return sum(1 for word in text.split() if word.lower() in positive_words)

def negative_score(text):
    return sum(1 for word in text.split() if word.lower() in negative_words)

def polarity_score(positive, negative):
    return (positive - negative) / ((positive + negative) + 0.000001)

def subjectivity_score(text):
    blob = TextBlob(text)
    return blob.sentiment.subjectivity

def avg_sentence_length(text):
    sentences = re.split(r'[.!?]', text)
    words = text.split()
    return len(words) / max(1, len(sentences))

def percentage_of_complex_words(text):
    words = text.split()
    complex_words = sum(1 for word in words if syllable_count(word) > 2)
    return (complex_words / len(words)) * 100

def fog_index(avg_sentence_len, perc_complex_words):
    return 0.4 * (avg_sentence_len + perc_complex_words)

def complex_word_count(text):
    return sum(1 for word in text.split() if syllable_count(word) > 2)

def word_count(text):
    return len(text.split())

def syllable_per_word(text):
    words = text.split()
    return sum(syllable_count(word) for word in words) / len(words)

def personal_pronouns(text):
    pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.I)
    return len(pronouns)

def avg_word_length(text):
    words = text.split()
    return sum(len(word) for word in words) / len(words)

def syllable_count(word):
    word = word.lower()
    vowels = "aeiouy"
    count = 0
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count += 1
    return count

# Example positive and negative words lists
positive_words = {"good", "great", "excellent", "positive", "fortunate", "correct", "superior"}
negative_words = {"bad", "terrible", "poor", "negative", "unfortunate", "wrong", "inferior"}

# Create a directory to save the text files
output_dir = 'articles'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Initialize results list
results = []

# Extract articles, analyze text, and save results
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Assuming the article title is within <title> tags and the article text is within <article> tags
        title = soup.find('title').get_text()
        article_text = soup.find('article').get_text()
        
        # Save the article text to a file
        with open(os.path.join(output_dir, f'{url_id}.txt'), 'w', encoding='utf-8') as file:
            file.write(title + '\n' + article_text)
        
        text = title + '\n' + article_text
        
        pos_score = positive_score(text)
        neg_score = negative_score(text)
        pol_score = polarity_score(pos_score, neg_score)
        subj_score = subjectivity_score(text)
        avg_sent_len = avg_sentence_length(text)
        perc_complex_words = percentage_of_complex_words(text)
        fog_idx = fog_index(avg_sent_len, perc_complex_words)
        avg_words_per_sent = avg_sent_len
        comp_word_count = complex_word_count(text)
        word_cnt = word_count(text)
        syll_per_word = syllable_per_word(text)
        pers_pronouns = personal_pronouns(text)
        avg_word_len = avg_word_length(text)
        
        results.append({
            'URL_ID': url_id,
            'POSITIVE SCORE': pos_score,
            'NEGATIVE SCORE': neg_score,
            'POLARITY SCORE': pol_score,
            'SUBJECTIVITY SCORE': subj_score,
            'AVG SENTENCE LENGTH': avg_sent_len,
            'PERCENTAGE OF COMPLEX WORDS': perc_complex_words,
            'FOG INDEX': fog_idx,
            'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sent,
            'COMPLEX WORD COUNT': comp_word_count,
            'WORD COUNT': word_cnt,
            'SYLLABLE PER WORD': syll_per_word,
            'PERSONAL PRONOUNS': pers_pronouns,
            'AVG WORD LENGTH': avg_word_len
        })
        
        print(f'Successfully processed article {url_id}')
    except Exception as e:
        print(f'Failed to process article {url_id}: {e}')

# Convert results to DataFrame and merge with output_df
results_df = pd.DataFrame(results)
output_df = output_df.merge(results_df, on='URL_ID', how='left')

# Save the output to a new Excel file
output_df.to_excel('Output.xlsx', index=False)
