# Term Extractor

Term Extractor is a desktop application built with Python and Tkinter that allows you to extract candidate terms and their contexts from various text-based documents. It supports multiple file formats including `.txt`, `.docx`, `.xliff`, and `.html`. The tool uses spaCy's NLP capabilities to identify multi-word terms and provides a user-friendly interface to review, select, and export extracted terms to CSV.

<img width="894" height="675" alt="image" src="https://github.com/user-attachments/assets/afb5cd3a-7ad6-4a55-b2ed-328b92fe4983" />


---

## Features

* Supports multiple input file formats: TXT, DOCX, XLIFF, HTML
* Extracts candidate terms using linguistic patterns (adjective+noun, noun+noun, noun of noun, etc.)
* Displays term frequency and context
* Interactive GUI to browse files, adjust minimum frequency, and review terms
* Select/deselect terms for export
* Export selected terms and their contexts to CSV
* Progress bar and multi-threading to keep UI responsive during extraction

---

## Requirements

* Python 3.7+
* spaCy (`en_core_web_sm` model)
* pandas
* python-docx
* beautifulsoup4

---

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/yourusername/term_extractor.git
   cd term_extractor
   ```

2. Create a virtual environment and activate it (optional but recommended):

   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/macOS
   venv\Scripts\activate     # Windows
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Download the spaCy English model:

   ```bash
   python -m spacy download en_core_web_sm
   ```

---

## Usage

Run the application:

```bash
python term_extractor.py
```

1. Click **Browse** to select a supported input file.
2. Set the minimum frequency for terms to be extracted (default is 1).
3. Click **Extract terms** to process the file.
4. Review extracted terms, toggle selection using the checkbox column.
5. Click on a term to view its context with the term highlighted.
6. Click **Export Selected to CSV** to save selected terms and contexts.

---

## Supported File Formats

* Plain text (`.txt`)
* Microsoft Word documents (`.docx`)
* XLIFF translation files (`.xliff`)
* HTML files (`.html`)

---

## Acknowledgments

* [spaCy](https://spacy.io/) for natural language processing
* [python-docx](https://python-docx.readthedocs.io/en/latest/) for Word document reading
* [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/) for HTML parsing
