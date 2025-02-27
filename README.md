# ITS Name Email Extractor

## Overview
ITS CV Name & Email Extractor is a Python-based tool designed to extract names and email addresses from resumes (CVs) in PDF and DOCX formats. It uses:
- Natural Language Processing (NLP) with spaCy and additional techniques to detect names.
- Regex-based extraction for identifying email addresses.
- SequenceMatcher from Python’s difflib module to compare detected names with email usernames, improving accuracy.
- A Tkinter GUI for easy folder selection and batch processing of multiple CVs.


## Features
- Extracts names and email addresses from PDF and DOCX resumes.
- Uses spaCy's en_core_web_lg NLP model to detect names.
- Identifies the most probable candidate name by analyzing text structure and similarity to email usernames.
- Processes entire folders containing multiple resumes at once.
- Displays extracted data in a user-friendly Tkinter GUI.
- Provides an option to export results to Excel for easy review.

## How Name Extraction Works
The tool processes resumes in a selected folder. Here's how it determines the best name candidate:

1. **Extracts names using multiple methods**:
   - The first two words of the resume.
   - Names detected using spaCy.
   - A name derived by comparing email usernames with extracted names using `SequenceMatcher`.

2. **Uses a majority voting approach to select the most probable name**:
   - If a name appears multiple times, it is selected as the most probable.
   - If no clear majority exists, the name most similar to the email username is prioritized.
   - If still undecided, it falls back to the first two words or spaCy-extracted names.


## How Email Extraction Works
The tool extracts email addresses from resumes using **regex-based extraction**. It scans the text of each resume for patterns that match typical email formats (e.g., `username@domain.com`). 



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

### 3. Run the Application
Execute the script with:

```sh
python app.py
```

## Installation & Usage

1. Clone or download this repository.
2. Run the script:
```sh
 python app.py
 ```
3. Use the GUI to select a folder containing resumes.
4. Click "Save to Excel" to export the results.

## Notes
- This project uses **Python 3.10**.
