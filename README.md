# 🧠 Automated Web Article Sentiment and Readability Analyzer

## 🌟 Project Overview

This Python script is a comprehensive tool designed to **automate the process of extracting text content** from a list of provided URLs and subjecting that content to a deep linguistic analysis.  
It calculates various **sentiment scores** as well as multiple **quantitative readability and complexity metrics**, storing the final results in a structured **Excel output file**.

---

## ✨ Features

### 🔹 Automated Web Scraping
- Uses `requests` and `BeautifulSoup` to fetch web pages.
- Intelligently extracts the **core article or blog content**, minimizing noise from headers, footers, and advertisements.

### 🔹 Custom Dictionary Integration
- Loads and utilizes **custom positive and negative word lists** from the `MasterDictionary` for sentiment scoring.

### 🔹 Advanced Stop Word Filtering
- Combines **NLTK’s default stop words** with **custom stop words** loaded from the `StopWords` directory.
- Ensures **highly accurate text cleaning** for analysis.

### 🔹 Comprehensive Metric Calculation
Generates **14 distinct metrics**, including:

| Category | Metrics |
|-----------|----------|
| **Sentiment Scores** | Positive Score, Negative Score, Polarity Score, Subjectivity Score |
| **Readability Indices** | Average Sentence Length, Percentage of Complex Words, FOG Index |
| **Text Structure** | Word Count (cleaned), Syllable Per Word, Personal Pronouns, Average Word Length |

### 🔹 Excel I/O
- Reads input URLs from **`Input.xlsx`**.
- Exports all analysis results to **`output.xlsx`**.

---

## 🛠️ Installation and Dependencies

### 1. Install Required Packages

This project requires **Python 3.x** and the following libraries:

```bash
pip install pandas requests beautifulsoup4 nltk openpyxl
```


## 2. NLTK Data

The script automatically checks for and downloads the necessary NLTK packages (punkt and stopwords) on its first run, so no manual action is typically required.

### 📁 Setup and Input Data

The script relies on a specific folder structure to locate the input data and linguistic dictionaries. Please ensure the following files and folders are present in the same directory as your main Python script.

```bash
.
├── sentiment_analysis.py    # Your Python script
├── Input.xlsx               # Required: Input file with URL_ID and URL columns
├── StopWords/               # Required: Directory containing custom stop word lists (e.g., .txt files)
└── MasterDictionary/        # Required: Directory containing sentiment word lists
    ├── positive-words.txt
    └── negative-words.txt
```


## Input File Format (Input.xlsx)

The input Excel file must contain, at minimum, these two columns:

| Column Name | Description                                          |
| ----------- | ---------------------------------------------------- |
| `URL_ID`    | A unique identifier for the article.                 |
| `URL`       | The complete web address to be scraped and analyzed. |


### 🚀 Usage

Prepare: Ensure all dependencies are installed and the file structure is correct.

Run: Execute the Python script from your terminal:

```bash
python sentiment_analysis.py
```


# Output

Upon successful execution, the script will generate a new Excel file named output.xlsx in the same directory. This file will contain the original URL_ID and URL, followed by the 14 computed analysis metrics for each article.
