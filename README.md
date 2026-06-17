# CHPL Tracker - Litigation Information Extraction Tool
Completed by: Kayla Logar

## Overview

This project, completed for the O'Neill Institute's Center for Health Policy and the Law (CHPL), parses through pdfs of complaints and dockets downloaded from Bloomberg Law. As an output, the code generates 3 files, in csv, Excel, and JSON formats, containing both metadata about the cases and information about the goals, issues, and impact of each case. To compile this information, the Python code parses through the PDFs, pulls some information from Excel files, applies validation rules according to the CHPL team's conventions, and uses AI to generate several fields. Please see below for a list of the fields contained in the output, along with the source(s) used to generate each field:

| Field | Source(s) |
|---|---|
| File Name | Bloomberg Docket |
| Case Name | Bloomberg Docket |
| Docket Number | Bloomberg Docket |
| Date Filed | Bloomberg Docket |
| Court | Bloomberg Docket; 28 USC Ch5 District Courts.pdf |
| Division | Bloomberg Docket; 28 USC Ch5 District Courts.pdf |
| Judge | Bloomberg Docket |
| President Who Appointed Judge | Judge Field; Federal Judicial Center Export.csv |
| Plaintiffs | Bloomberg Docket |
| Intervenor Plaintiffs | Bloomberg Docket |
| Lead Counsel for Plaintiff(s) | Bloomberg Docket |
| Defendants | Bloomberg Docket |
| Intervenor Defendants | Bloomberg Docket |
| Lead Counsel for Defendant(s) | Bloomberg Docket |
| Goals | Complaint; AI; GoalsMapping.csv; GoalsExamples.xlsx |
| Issues | Complaint; IssuesMapping.csv; AI (fallback); LegalIssuesExamples.xlsx (fallback); Bloomberg Docket (fallback if complaint unreadable) |
| Potential Impact | Complaint; AI; AnalysisExamples.xlsx |
| Why This Matters | Complaint; AI; AnalysisExamples.xlsx |


The CHPL team is currently working to integrate this output with their healthcare litigation tracking application. They are in the midst of working with a software team to update the website, and I have been communicating with both teams about the potential to use this tool as a preliminary step, with quality review protections in place. Because I am employed as a research assistant with the CHPL team, and will continue to work with them this summer, I will continue to work on this project with them to ensure the smooth handoff of a polished product with thorough documentation so they may continue to use the product into the future. Plans for finishing the project include additional quality testing in even more scenarios (I have already done some but aspire to be as thorough as possible) and ensuring the output is consistent with all the team's data validation conventions and goals. As an aside, though I work with the team as a research assistant, please note that I did not log any hours or receive any pay for work completed on this project during the school semester because the project was primarily for this class.

## User Documentation

### Running the Code Locally

#### Prerequisites

- **Python 3.11 or later** — To run the code successfully, a user should confirm that the version of Python they are working with is Python 3.11 or later. To check the version currently installed, they can run the following in a terminal: `python --version` (or `python3 --version` on Mac/Linux) 
- **An OpenRouter API key** — To successfully connect to OpenRouter for the AI-generated fields, a user who is running the code locally will also need an OpenRouter API key. Anyone may acquire an OpenRouter API key by. creating an account at [openrouter.ai](https://openrouter.ai) and generating a key under *Keys*. Note that, as the model is currently configured, the LLM is not free, so the user will be charged when they run the code that connects to the LLM. The user must have some credit in their OpenRouter balance to successfully run the code without encountering an error.
- **Bloomberg Law access** — The code is written under the presumption that the user is able to access Bloomberg law to download docket and complaint PDFs. The files currently in the "Complaints" and "Dockets" folders were used for testing and may be used to test and review the code.

#### Setup

1. **Create a codespace**: One effective way for the user to access and use the code is to begin by opening the GitHub repository and creating a Codespace.

2. **Create and activate a virtual environment**: By implementing the following code, depending on the applicable operating system, a user can isolate any dependencies from their global Python system.

   ```bash
   # Mac / Linux
   python3 -m venv .venv
   source .venv/bin/activate

   # Windows
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. **Install dependencies:** The user should next run the following line of code to ensure all dependencies needed for the code to run smoothly are installed.

   ```bash
   # Mac / Linux
   pip3 install -r requirements.txt

   # Python
   pip install -r requirements.txt
   ```
    This will then install the following requirements listed in the requirements.txt file:
    - streamlit version 1.57.0
    - pandas version 3.0.2
    - openpyxl version 3.1.5
    - PyMuPDF version 1.27.2.3
    - pymupdf4llm version 1.27.2.3
    - openai version 2.36.0
    - pypdf version 6.10.2
    - titlecase version 2.4.1
    - certifi version 2026.4.22

4. **Configure your API key**: To ensure effective connection with the OpenRouter API, the user should create a file at `.streamlit/secrets.toml` in the project folder with the following content:

   ```toml
   export OPENROUTER_API_KEY = "sk-or-your-key-here"
   ```

5. **Download the latest federal judge data**: To optimize the accuracy of the check for the "President Who Appointed Judge" field for each case, the user should be begin by running the following code to make sure the Federal Judicial Center export is up-to-date:

   ```bash
   # Mac / Linux
   python3 update_judge_data.py

   # Windows
   python update_judge_data.py
   ```

6. **Launch the web app:** Now that the preliminary set-up steps are taken care of, the user may launch the web application to process the documents. Note that, from here, the user may follow the further instructions found below in the "Running the Code via Web Application" section.

   ```bash
   # Mac / Linux
   python3 -m streamlit run tracker_web_app.py

   # Windows
   python -m streamlit run tracker_web_app.py
   ```

   Streamlit will print a local URL (typically `http://localhost:8501`) — open it in any browser.

7. **Run the code on the backend:** If the user would prefer to run the code on the backend rather than by navigating the web application, they may do so by running the command below.

   ```bash
   # Mac / Linux
   python3 summarize_cases.py

   # Windows
   python summarize_cases.py
   ```

   The resulting output files will then appear in the folder where the code was run, with the below names (where MM.DD.YY denotes the month, date, and year on which the code was run). Note that when the user runs the script multiple times in one day, the file names will have a number at the end, e.g., (1), (2), to denote which instance of the script generated each file.
   - Tracker Data Summary MM.DD.YY.csv
   - Tracker Data Summary MM.DD.YY.xlsx
   - Tracker Data Summary MM.DD.YY.json
   

### Running the code via web application

A user may navigate to the following url to access the web application directly: 

This is the simplest path to test the application because it does not require configuring or preparing a Python virtual environment. However, one **major caveat** of using the web application is that files uploaded during a session will only be present for the duration of the current session. Thus, when a user accesses the application using this link, the Output Files page will only contain files generated during the current session, not during previous sessions, and updates to other files used as inputs to the model similarly will not permanently change these files on the backend. Thus if a user needs to make such long-term changes, they should first do so locally and then git commit and push to the github repo.

The user may use any Complaints and Dockets in the corresponding GitHub folders to test the application. See the "Guide to web application" section below for further instructions regarding the web application.

### Guide to web application

#### Process Cases

The "Process Cases" page of the web application is the main page for the central analysis to the case. Users will first upload any pdfs of case dockets downloaded from Bloomberg Law and then upload any pdfs of complaints from the cases. The script matches dockets and complaints based on filenames, so users should name docket pdfs according to the naming convention "[Case Name], Docket.pdf" and complaint PDFs according to the naming convention "[Case Name], Complaint.pdf" to assist the code in coordinating these. Users looking to test the application may use the files in the Test-Input-Files folder in GitHub and should upload the files in pairs so that each case processed has both a docket file and a complaint file.

Uploaded files will appear in the "Staged dockets" and "Staged complaints" dropdowns under the spaces for these files to be uploaded. When the application is launched locally, complaints and dockets stored in the "TrialCourtComplaints" and "TrialCourtDockets" folders will automatically be in the "Staged dockets" and "Staged complaints" dropdowns, prepared to process. Note that if they are removed here, they will also be removed from the corresponding folders on the backend.

Underneath the uploaded files, there is an indicator of whether the OpenRouter API Key is configured, which the user may reference to ensure that the API key is properly set up to process the files.

Next, the user will have the opportunity to name the output file - this name will default to "Tracker Data Summary MM.DD.YY" to be consistent with the backend naming convention.

After the user has confirmed the pdf uploads and made any desired changes to the output file name, they should click the Process button to begin the process of parsing through the PDFs to generate the output. A message will appear telling the user that the process is running, and it will take some time, particularly if processing many cases simultaneously. When the process is complete, a message will pop up saying "Processing complete! 3 file(s) saved to Output Files."

#### Output Files

On the Output Files page, users can download, rename, and delete files generated on the Process Cases page. 

#### Judicial Data

On the Judicial Data page, the user can click the "UPDATE JUDICIAL DATA" button to run the update_judge_data.py file, which pulls updated data from the Federal Judicial Center. This will then update on the backend. Additionally, the user can preview and download the data from the "Federal Judicial Center Export.csv" file that stores the data from the Federal Judicial Center.

#### District Courts

The District Courts page permits users to download or preview 18 U.S.C. Chapter 5, which lists the federal district courts and divisions by state. Users may also upload an updated version of this code, which will then replace the existing pdf on the backend.

#### Other Model Inputs

The Other Model Inputs page permits users to download, preview, and replace additional files that act as inputs to the data extraction process, including:
- **Goals Mapping:** A file mapping phrases that may appear in the Relief Requested portions of complaints to associated goals of litigants
- **Goals Examples:** A file with examples of cases, requests made in those cases, and the goals those requests correspond to
- **Issues Mapping:** A file mapping phrases that may appear in the portions of complaints expressing the legal claims and counts to associated legal issues
- **Issues Examples:** A file with examples of cases and the legal issues that arise in those cases
- **Analysis Examples:** A file with examples of cases and the "Potential Impact" and "Why This Matters" passages for those cases, as drafted by the CHPL team
