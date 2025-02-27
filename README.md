# ITS Name Email Extractor

## Overview
ITS CV Name & Email Extractor is a Python-based tool designed to extract names and email addresses from resumes (CVs) in PDF, DOCX and DOC formats. It uses:
- Natural Language Processing (NLP) with spaCy and additional techniques to detect names.
- Regex-based extraction for identifying email addresses.
- SequenceMatcher from Python’s difflib module to compare detected names with email usernames, improving accuracy.
- A Tkinter GUI for easy folder selection and batch processing of multiple CVs.


## Features
- Extracts names and email addresses from PDF, DOCX and DOC resumes.
- Uses spaCy's en_core_web_lg NLP model to detect names.
- Identifies the most probable candidate name by analyzing text structure and similarity to email usernames.
- Processes entire folders containing multiple resumes at once.
- Displays extracted data in a user-friendly Tkinter GUI.
- Provides an option to export results to Excel and JSON for easy review.

## How Name Extraction Works
The tool processes resumes in a selected folder. Here's how it determines the best name candidate:

1. **Extracts names using multiple methods**:
   - The first two words of the resume.
   - Names detected using spaCy.
   - A name derived by comparing email usernames with extracted names using `SequenceMatcher`.

2. **Uses a majority voting approach to select the most probable name**:
   - If a name appears multiple times, it is selected as the most probable name.
   - If no majority is found, the name most similar to the email username is prioritized, provided the similarity ratio exceeds 50%.
   - If there's still no clear choice, the tool will fall back to the spaCy-extracted names, selecting the one with the highest similarity to the email username.
   - As a final fallback, the first two words are used if all else fails.

## How Email Extraction Works
The tool extracts email addresses from resumes using regex-based extraction, identifying patterns like username@domain.c

To find a name similar to the email username:
   - Only the three lines surrounding the email-containing line are considered for matching.
   - Words from these lines are compared to the email username using SequenceMatcher to determine the best match.
   - Then the tool attempts to combine adjacent words (e.g., first last) and compares the combined string with the email username.
It checks both the natural and reversed order of adjacent words (e.g., last first) to increase the chances of finding a best match.


## Setup Instructions

### 1. Install Dependencies
Run the following command to install required packages:

```sh
pip install -r requirements.txt
```

### 2. Download spaCy Language Model
This project requires the **en_core_web_lg** model. Download it using:

```sh
python -m spacy download en_core_web_lg
```

### 4. Setting Up Tesseract OCR (For Image-to-Text Extraction)

Ensure Tesseract OCR is installed on your system:

- Download and install Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki
- Add Tesseract to the system PATH environment variable.

### 5. Installing Poppler (For PDF Processing)

Poppler is required for handling PDF rendering. Install it as follows:

- Download the latest Poppler for Windows from https://github.com/oschwartz10612/poppler-windows/releases.
- Extract the files and add the bin directory to the system PATH.


### 3. Run the Application
Execute the script with:

```sh
python app.py
```

## Usage

1. Use the GUI to browse and select a folder containing resumes.
2. Click "Save to Excel" to export the results to EXCEL.
3. Click "Save to JSON" to export the results to JSON.

## Notes
- This project uses **Python 3.10**.
