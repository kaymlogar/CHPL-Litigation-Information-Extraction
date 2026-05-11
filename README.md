# CHPL Tracker - Litigation Information Extraction Tool

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


The CHPL team is currently working to integrate this output with their healthcare litigation tracking application. They are in the midst of working with a software team to update the website, and I have been communicating with both teams about the potential to use this tool as a preliminary step, with quality review protections in place. Because I am employed as a research assistant with the CHPL team, and will continue to work with them this summer, I will continue to work on this project with them to ensure the smooth handoff of a polished product with thorough documentation so they may continue to use the product into the future. Plans for finishing the project include additional quality testing in even more scenarios (I have already done some but aspire to be as thorough as possible) and ensuring the output is consistent with all the team's data validation conventions and goals.


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
   OPENROUTER_API_KEY = "sk-or-your-key-here"
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

The "Process Cases" page of the web application is the main page for the central analysis to the case. Users will first upload any pdfs of case dockets downloaded from Bloomberg Law and then upload any pdfs of complaints from the cases. The script matches dockets and complaints based on filenames, so users should name docket pdfs according to the naming convention "[Case Name], Docket.pdf" and complaint PDFs according to the naming convention "[Case Name], Complaint.pdf" to assist the code in coordinating these. 

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

## Reflection

As technology evolves and particularly as AI improves and lawyers learn to leverage it in different ways, programming seems to be a very useful skill for lawyers. It is necessary to address concerns that are always prevalent in the practice of law, such as confidentiality and quality assurance, when lawyers attempt to incorporate programming into their practice, but this kind of work has the potential to make lawyers' work significantly more efficient. 

This project is a great example of how lawyers who have some knowledge of programming may complete work more efficiently, when permitted to use the necessary tools. Fortunately, the information involved is not confidential and therefore does not pose a confidentiality risk. The task that the code completes is something that used to take a member of the team at least an hour to complete. The technology has some ways to go to ensure that the level of quality consistently meets the team's expectations, particularly for the longer AI-generated written fields, but even a team member processing a case using this application and then confirming that all of the information is correct and revising the written fields will spend significantly less time doing so than they would have spent identifying and recording the information in all the fields manually. This demonstrates that incorporating computer programming into legal practice has great potential to make tasks more efficient.

It will be interesting to see how this course evolves with technology, as we have already seen as AI has become a more and more useful tool. One potential opportunity to enrich the study of computer programming and the law could be to incorporate a component of alumni connection. I would be interested in seeing how alumni, and especially alumni who have taken this course, have used these skills in practice. Particularly, I would be interested in whether alumni working for various types of employers (e.g., big law, government, non-profits) have found their employers encouraging or discouraging of their use of technology and how, if at all, their employers have permitted them to incorporate work like this into their practice.

## Project Process, Effort, and Use of AI

Overall, I began this project by writing  the code that was parsing the pdfs and extracting the necessary information (such as Case Name, Docket No., etc.) myself and using AI for support occasionally when debugging some issues that popped up. As I became more comfortable with AI, I started using it as a starting point to write new sections of the code. AI-assisted coding was very helpful with both debugging and creating a starting place for coding. One drawback I noticed was that sometimes, when debugging more complex issues related to parsing PDFs, the AI chat generated solutions that missed the root of the problem and tended to be overly tailored to the specific scenario at hand rather than broadly applicable. 

For example, it took me a long time to figure out how to best parse through the PDFs to obtain information regarding the lead counsel of each party given the two-column block structure. At first, I tried to do so myself, but after having difficulty with this for half a day, I asked the AI chat in cursor to help solve the problem. As a starting point, the AI agent helped me decide to split the data into two columns. However, the parsing from here still posed several problems, and after spending several hours attempting to work with the AI agent to develop a solution, I decided to switch back to manually developing a solution. Notably, this was when I was working predominantly with Cursor. 

Later in the course of the project, I switched to Claude Code as a coding assistant and found it significantly better at debugging issues. In addition to providing better solutions, I also found that Claude Code provided superior explanations so that I felt as though I really understood its approach where in Cursor, I may have spent longer trying to review the code to really see what the AI agent did and why. Because of this improved experience, after I switched to predominantly using Claude, I utilized the AI agent more heavily in my coding process and found it very helpful for quality testing, debugging, and cleaning up the code to make it more efficient. It was also very useful to create a baseline UI which I could then adjust as needed.

In addition to using AI as a coding assistant and directly in the code to generate several of the fields, I used AI to scan the CHPL team's litigation tracker website and generate a representative subset of cases that I could use to test. Before this, I had a few example cases from the team, but I asked Claude to find cases with a variety of combinations of values in all the relevant fields (e.g., different courts, different government parties) so that I could complete testing of the code's functionality, and this was very helpful. It was also very helpful to be able to ask Claude to ensure that I had recorded all the necessary dependencies for the code. 

As I worked on the project, there were several times that I expended effort that is not visible in the final result. First, I began the project parsing the metadata fields based on the Complaint PDFs, but after a certain point, I realized that they did not contain all the information I would need (for example, the Complaint PDFs do not contain the name of the judge appointed to the case). Because of this, and because the Complaint PDFs were often inconsistently formatted across various jurisdictions, I switched to parsing the Bloomberg Dockets as a primary source. I then went back and rewrote the code I had completed up to this point to pull information from the docket PDFs instead of the complaint PDFs, so the several hours I spent parsing the complaints did not end up having a major impact in the final product. 

Additionally, I spent a long time trying to figure out how to identify the Court Division, particularly because sometimes all that is present in the docket is the city where the court is, which does not always perfectly coincide with the name of the division. I spent about a day trying to resolve this - at first attempting to build a helper function that would parse through the PDF with the Court Division code. However, at the end of this time, I decided to simply directly pull the text from the PDF and parse through it while processing the code rather than generating a csv using a helper function, so the helper function I attempted to build is not present in the final product. Finally, as reflected in the discussion of AI use above, it took me probably about two days in total to figure out how to handle two-column PDF formatting and separate the column text into blocks for the Lead Counsel fields.
