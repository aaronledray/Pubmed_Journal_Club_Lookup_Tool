# Pubmed Journal Club Lookup Tool

- Papers are fetched from a pubmed query using journals of interest and keywords of interest.
- The user is prompted for date range for the pubmed queries.
- A powerpoint file containing a slide for each paper found is created.

---

## Preparing your environemnt (installing requirements):

- If you're unfamiliar with managing python packages, then I recommend:
  1) Install Anaconda - this is so you can use Conda as your package manager (https://www.anaconda.com/download)
    - If you're on a Windows machine, you will now need to launch the Anaconda Command Prompt .exe. This is the window where you will input command-line instructions in the next point.
    - If you're on a Mac or Linux machine, launch a terminal window for input:
  2) input in terminal: ```conda install -c conda-forge biopython```
    - note, -c conda-forge means use the channel conda-forge, biopython is the package to install.
  3) input in terminal: ```conda install -c conda-forge python-pptx```

---

## Running Instructions:

- modify config.txt to reflect:
  - your email address
  - single blank line!
  - journals of interest
  - single blank line!
  - keywords of interest

- Excamples for journals and keywords are found in examples_of_journal_names.txt and examples_of_keywords.txt
 
- in a terminal, navigate to the directory containing the script file Journal_Lookup_Tool.py
  - input in terminal: ```python Journal_Lookup_Tool.py```

- Default settings: fetch papers from 1 week period.
