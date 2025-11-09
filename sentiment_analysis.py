import pandas as pd

import requests

from bs4 import BeautifulSoup

import nltk

from nltk.tokenize import sent_tokenize, word_tokenize

from nltk.corpus import stopwords

import re

import os



df = pd.read_excel("Input.xlsx")

try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')

try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

try:
    nltk.data.find('tokenizers/punkt_tab')
except LookupError:
    nltk.download('punkt_tab')

class SentimentAnalyzer:
    def __init__(self, stopwords_folder='StopWords', master_dict_folder='MasterDictionary'):
        self.stopwords_folder = stopwords_folder
        self.master_dict_folder = master_dict_folder
        self.stop_words = self.load_stop_words()
        self.positive_words, self.negative_words = self.load_master_dictionary()
    
    def load_stop_words(self):
        stop_words = set(stopwords.words('english'))
        
        if os.path.exists(self.stopwords_folder):
            for filename in os.listdir(self.stopwords_folder):
                filepath = os.path.join(self.stopwords_folder, filename)
                try:
                    with open(filepath, 'r', encoding='latin-1') as f:
                        words = f.read().lower().split()
                        stop_words.update(words)
                except Exception as e:
                    print(f"Error loading {filename}: {e}")
        
        return stop_words
    
    def load_master_dictionary(self):
        positive_words = set()
        negative_words = set()
        
        if os.path.exists(self.master_dict_folder):
            # Load positive words
            pos_file = os.path.join(self.master_dict_folder, 'positive-words.txt')
            if os.path.exists(pos_file):
                with open(pos_file, 'r', encoding='latin-1') as f:
                    for line in f:
                        word = line.strip().lower()
                        if word and word not in self.stop_words:
                            positive_words.add(word)
            
            # Load negative words
            neg_file = os.path.join(self.master_dict_folder, 'negative-words.txt')
            if os.path.exists(neg_file):
                with open(neg_file, 'r', encoding='latin-1') as f:
                    for line in f:
                        word = line.strip().lower()
                        if word and word not in self.stop_words:
                            negative_words.add(word)
        
        return positive_words, negative_words
    
    def scrape_article(self, url):
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Remove script and style elements
            for script in soup(['script', 'style', 'nav', 'footer', 'header']):
                script.decompose()
            
            # Try to find article content
            article_text = ''
            
            # Common article selectors
            article = soup.find('article')
            if article:
                article_text = article.get_text()
            else:
                # Try other common content containers
                for tag in ['div', 'main', 'section']:
                    content = soup.find(tag, class_=re.compile(r'content|article|post|entry', re.I))
                    if content:
                        article_text = content.get_text()
                        break
            
            # Fallback to body if nothing found
            if not article_text:
                article_text = soup.get_text()
            
            # Clean text
            lines = (line.strip() for line in article_text.splitlines())
            chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
            article_text = ' '.join(chunk for chunk in chunks if chunk)
            
            return article_text
        
        except Exception as e:
            print(f"Error scraping {url}: {e}")
            return ""
    
    def clean_text(self, text):
        words = word_tokenize(text.lower())
        cleaned_words = [word for word in words if word.isalpha() and word not in self.stop_words]
        return cleaned_words
    
    def count_syllables(self, word):
        word = word.lower()
        vowels = 'aeiou'
        syllable_count = 0
        previous_was_vowel = False
        
        for char in word:
            is_vowel = char in vowels
            if is_vowel and not previous_was_vowel:
                syllable_count += 1
            previous_was_vowel = is_vowel
        
        # Handle exceptions for words ending with 'es' or 'ed'
        if word.endswith('es') or word.endswith('ed'):
            syllable_count -= 1
        
        # Every word has at least one syllable
        if syllable_count == 0:
            syllable_count = 1
        
        return syllable_count
    
    def count_personal_pronouns(self, text):
        pattern = r'\b(I|we|my|ours|us)\b' # use of regex to find pronouns
        matches = re.findall(pattern, text, re.IGNORECASE)
        
        # Filter out 'US' (country) by checking if it's uppercase
        pronouns = [m for m in matches if not (m == 'US' or m == 'us' and re.search(r'\bUS\b', text))]
        
        return len(pronouns)
    
    def analyze_text(self, text):
        if not text:
            return self.get_empty_metrics()
        
        # Tokenize sentences and words
        sentences = sent_tokenize(text)
        words = word_tokenize(text.lower())
        
        # Clean words (remove punctuation and stop words)
        cleaned_words = self.clean_text(text)
        
        # Calculate scores
        positive_score = sum(1 for word in cleaned_words if word in self.positive_words)
        negative_score = sum(1 for word in cleaned_words if word in self.negative_words)
        
        # Polarity Score
        polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
        
        # Subjectivity Score
        total_words_cleaned = len(cleaned_words)
        subjectivity_score = (positive_score + negative_score) / (total_words_cleaned + 0.000001)
        
        # Readability metrics
        total_words = len([w for w in words if w.isalpha()])
        total_sentences = len(sentences)
        
        avg_sentence_length = total_words / (total_sentences + 0.000001)
        
        # Complex words (more than 2 syllables)
        complex_words = [w for w in words if w.isalpha() and self.count_syllables(w) > 2]
        complex_word_count = len(complex_words)
        percentage_complex_words = complex_word_count / (total_words + 0.000001)
        
        # Fog Index
        fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
        
        # Syllable count per word
        total_syllables = sum(self.count_syllables(w) for w in words if w.isalpha())
        syllable_per_word = total_syllables / (total_words + 0.000001)
        
        # Personal pronouns
        personal_pronouns = self.count_personal_pronouns(text)
        
        # Average word length
        total_characters = sum(len(w) for w in words if w.isalpha())
        avg_word_length = total_characters / (total_words + 0.000001)
        
        return {
            'POSITIVE SCORE': positive_score,
            'NEGATIVE SCORE': negative_score,
            'POLARITY SCORE': round(polarity_score, 4),
            'SUBJECTIVITY SCORE': round(subjectivity_score, 4),
            'AVG SENTENCE LENGTH': round(avg_sentence_length, 2),
            'PERCENTAGE OF COMPLEX WORDS': round(percentage_complex_words, 4),
            'FOG INDEX': round(fog_index, 2),
            'AVG NUMBER OF WORDS PER SENTENCE': round(avg_sentence_length, 2),
            'COMPLEX WORD COUNT': complex_word_count,
            'WORD COUNT': total_words_cleaned,
            'SYLLABLE PER WORD': round(syllable_per_word, 2),
            'PERSONAL PRONOUNS': personal_pronouns,
            'AVG WORD LENGTH': round(avg_word_length, 2)
        }
    
    def get_empty_metrics(self):
        return {
            'POSITIVE SCORE': 0,
            'NEGATIVE SCORE': 0,
            'POLARITY SCORE': 0,
            'SUBJECTIVITY SCORE': 0,
            'AVG SENTENCE LENGTH': 0,
            'PERCENTAGE OF COMPLEX WORDS': 0,
            'FOG INDEX': 0,
            'AVG NUMBER OF WORDS PER SENTENCE': 0,
            'COMPLEX WORD COUNT': 0,
            'WORD COUNT': 0,
            'SYLLABLE PER WORD': 0,
            'PERSONAL PRONOUNS': 0,
            'AVG WORD LENGTH': 0
        }

# Main execution
def process_urls(df, url_column='URL'):
    analyzer = SentimentAnalyzer()
    results = []
    
    for idx, row in df.iterrows():
        url = row[url_column]
        print(f"Processing {idx + 1}/{len(df)}: {url}")
        
        # Scrape article
        text = analyzer.scrape_article(url)
        
        # Analyze text
        metrics = analyzer.analyze_text(text)
        
        # Add URL to results
        result = {'URL_ID': row['URL_ID'],'URL': url}
        result.update(metrics)
        results.append(result)
    
    # Create output dataframe
    output_df = pd.DataFrame(results)
    
    # Export to Excel
    output_df.to_excel('output.xlsx', index=False)
    print("\nAnalysis complete! Results saved to output.xlsx")
    
    return output_df

output = process_urls(df, url_column='URL')
