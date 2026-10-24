96-Well Plate Map Reformatter
============================

PURPOSE
-------
This school project turns an Excel 96-well plate map into a formatted workbook
with a plate layout and sample lookup. The main interface is a single-page
Streamlit web app that runs locally or can be hosted for an instructor.

FEATURES
--------
- Upload .xlsx or legacy .xls workbooks.
- Enter a plate name and create an .xlsx output named StudyNumber_PlateName.xlsx.
- See missing-data prompts, warnings, a short processing indicator, and a
  downloadable result.
- Correct missing information per well. The app previews original and corrected
  entries side by side before making the workbook.
- Keep identical sample entries together only when they are horizontally
  adjacent. Labels are recalculated in the existing sequence after corrections.
- Warn when different source entries share a sample number and test period; the
  entries remain separate for review.
- Before creating a workbook, flag separate lookup rows with the same sample
  number, test period, and dilution, ignoring capitalization. Review every group
  and choose whether to continue. If a group is declined, workbook creation is
  canceled.
- Warn about irregular plate headers, row labels, and extra or missing grid
  dimensions.
- Include a Review Notes sheet when manual corrections or processing warnings
  exist. Clean workbooks omit that sheet.
- Preserve the displayed source test-period wording in the lookup.
- Format plate and lookup sheets, including alternating lookup-row shading and
  panes frozen at A3.

WARNING EXAMPLES
----------------
- If two wells have the same sample number and test period but different full
  source entries, the warning names both wells and entries. For example, two
  wells may share a sample number and Day 7 but have different matrix/split or
  dilution text. They stay in separate lookup rows for review.
- A message such as "missing column number 9" means the worksheet's numbered
  plate headers do not include all of 1–12. The app warns that the generated
  grid may not match the source layout.
- A message about non-empty cells below the 8-row plate area means there may be
  extra plate rows or other content in that area. The output includes rows A–H;
  check the source before using it.

INSTALL AND RUN LOCALLY (WINDOWS)
---------------------------------
Install Python 3.10 or newer from https://www.python.org/downloads/ . Open
Command Prompt in this project folder and run:

    py -m venv .venv
    .venv\Scripts\activate
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    python -m streamlit run app.py

Streamlit prints a local address in the terminal and normally opens it in your
browser. Keep the terminal window open while using the app; press Ctrl+C there
to stop it. After installing the requirements once, you can also double-click
Run_Reformatter.bat. The original desktop workflow is still available with:

    python plate_reformatter.py

INPUTS AND OUTPUTS
------------------
- .xlsx files are read from their active worksheet. Legacy .xls files are read
  from their first worksheet using xlrd.
- The input is expected to contain one 96-well plate. The app checks the plate
  layout and warns when the detected grid is irregular.
- The original source file is not changed. The Streamlit app provides the
  generated .xlsx as a download.
- The output contains the formatted plate map, sample lookup, and a Review Notes
  sheet only when corrections or warnings need review.
- The app does not guess missing metadata or silently reassign samples. If
  required information is missing, correct it or cancel before output creation.
- Different complete source entries with the same sample number and test period
  are flagged. Review their source entries before using the workbook.
- Separate lookup rows with the same sample number, test period, and dilution
  (ignoring capitalization) require a decision for each group before workbook
  creation. Approved groups are listed in Review Notes with their labels, wells,
  and approval status.

PER-FILE AUDIT USING A SUPPLIED WORKBOOK
----------------------------------------
The audit helper processes one supplied workbook at a time. It copies the source
to a temporary directory, writes generated output separately to test_outputs,
and reports labels, wells, metadata, warnings, and formatting checks. It does
not compare against an expected workbook; no expected outputs have been
provided. It does not create sample workbooks or synthetic data.

Example using supplied Workbook 2:

    python audit_one_workbook.py "2nd ID test for school.xls" --plate-name T01

For another supplied file, replace the quoted filename. Inspect the generated
workbook and console report before moving to another file.

THREE-MINUTE DEMONSTRATION
-------------------------
Use the supplied file "2nd ID test for school.xls" and plate name T01. A
minute-by-minute script is in DEMO_PLAN.md. The demonstration shows legacy .xls
upload, workbook checking, output creation, the retained QC1-1 label, and review
of the generated map and lookup.

INSTRUCTOR-FACING CHECKLIST
---------------------------
Before the demonstration:
1. Decide whether the student will show the app locally or provide an already
   published link. The app has not been published, so a hosted link does not yet
   exist.
2. If showing the app locally, confirm Python and the project requirements are
   installed, then start the app with `Run_Reformatter.bat` or the local command
   above. Keep the terminal open while using it.
3. Have the supplied `2nd ID test for school.xls` workbook ready. Do not upload
   confidential workbooks to a hosted app unless that use has been approved.

During the demonstration:
1. Upload `2nd ID test for school.xls` and enter plate name `T01`.
2. Select **Check workbook** and review the detected study number and any
   missing information or warnings.
3. If the app reports a correction or duplicate-review step, review the listed
   source information before continuing. Cancel if the information cannot be
   confirmed.
4. Create and download the workbook. Confirm the downloaded filename, then
   inspect the plate map, lookup, and any Review Notes sheet against the source.
5. Explain that an output is a review aid and must be checked against its source.

HOSTING CHOICES
---------------
Choose based on how the instructor should access the app and whether the files
may be handled by a hosted service.

- **Run locally:** Streamlit runs on your computer and opens in a browser. Show
  it on that computer or share the screen. The uploaded workbook stays on the
  local processing path.
- **Streamlit Community Cloud:** The app runs remotely and provides a URL to
  share with the instructor. Uploaded workbooks are sent to the hosted app for
  processing. Use this option only with files approved for hosted processing.
  Private access settings may limit who can open the app, but processing still
  takes place on the hosted service.

For Community Cloud, see the official
[deployment guide](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)
and [sharing guide](https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app).

To publish after choosing this route:
1. Create a GitHub repository containing the application source, requirements,
   and documentation. The supplied .gitignore excludes Excel workbooks and
   generated test outputs.
2. Sign in to Streamlit Community Cloud with GitHub and choose **Create app**.
3. Select the repository, branch, and `app.py` entrypoint, then deploy.
4. Choose the app's access setting. If private access is available for your
   account, invite the instructor as a viewer and share the app URL.
5. Use only workbooks approved for hosted processing. The app has not been
   published, and no project files have been uploaded to an external service.

Current Streamlit sharing and account limits can change; check its official
documentation before publishing. In particular, check the current rules for
private apps and viewer access. Do not commit supplied Excel files or generated
workbooks to GitHub. Neither hosting option has been published or set up for
instructor access.

DEPLOYMENT TROUBLESHOOTING
-------------------------
If Streamlit Cloud shows `ImportError` mentioning `_tkinter`, check that the
latest `plate_reformatter.py` has been pushed to the GitHub repository. The
desktop file-picker imports Tkinter only when the desktop entry point runs;
the Streamlit app does not need Tkinter. Do not add `tkinter` to
`requirements.txt`: it is part of some desktop Python installations, not a
package needed by the hosted web app. After pushing the updated file, wait for
Streamlit Cloud to redeploy and check **Manage app** logs if startup still fails.

KNOWN LIMITATIONS AND REVIEW
----------------------------
- The app reads one worksheet (active .xlsx sheet or first .xls sheet) and
  expects a recognizable 96-well layout. Warnings ask for human review; they do
  not repair an unusual worksheet layout.
- Missing or ambiguous metadata requires user input. A full corrected source
  entry uses this order: Study Number, Custom ID, Sample Number, Group Number,
  Matrix/Split, Test Period, Dilution Fold. Matrix/Split may contain spaces.
- The app does not compare the total number of source sample wells with output
  assignments. Review the plate map and lookup against the source workbook.
- No expected output workbooks were supplied, so the audit helper cannot verify
  exact expected results for every input.
- The output is a review aid, not a laboratory decision system. Check each
  generated workbook against its original before use.

FINAL PROJECT EVALUATION
------------------------
Problem and intended users
The project addresses the time and transcription risk involved in manually
reformatting 96-well plate maps and transferring sample details into a lookup.
It is intended for students and laboratory staff who need a consistent,
reviewable Excel output from a supported plate workbook.

Application and capabilities
The final application is a single-page Streamlit app. It accepts .xlsx and
legacy .xls files, asks for a plate name, creates a formatted .xlsx plate map
and lookup, and offers the result for download. It keeps horizontally adjacent
identical sample entries together, recalculates labels after corrections,
prompts users to complete missing metadata, and asks for review before creating
a workbook when duplicate lookup groups are found. Processing warnings and
manual corrections are recorded in Review Notes when applicable.

Build environment and tools
The local check environment used Windows with Python 3.14.7, Streamlit 1.66.1,
openpyxl 3.1.5, and xlrd 2.0.2. These versions describe that environment; the
requirements file remains flexible and does not pin them. Python runs the
processing logic; Streamlit provides the web interface; openpyxl writes .xlsx
output; xlrd reads legacy .xls input. The project also includes a Windows
launcher, an audit helper, and a three-minute demonstration plan.

Testing performed and actual results
- The supplied Workbook 2 (.xls) was inspected after the parser update. The app
  identified study number 34567891 and reported no missing metadata or layout
  warnings. Its multiword Matrix/Split value was recognized as metadata.
- The supplied Workbook 7 was corrected using the user-provided entries for
  wells F4 and G10, then the workbook was regenerated. No expected output file
  was supplied for an exact cell-by-cell comparison.
- The source files supplied by the user were used for workbook processing; no
  synthetic workbooks were created. Because expected output workbooks were not
  provided, exact output accuracy for every well and formatting detail remains
  subject to review against each original source.

Known limitations and failure modes
The app expects one recognizable 96-well plate on the active .xlsx sheet or
first .xls sheet. Unusual layouts may produce warnings and need human review.
Missing or ambiguous metadata requires a user correction. The app does not
compare source sample-well totals with output assignments, and no expected
workbooks are available for automated exact-result comparison. Users should
check every output against its source before using it.

Improvements made and future work
The project added a Streamlit upload-and-download interface, legacy .xls
support, clearer missing-data corrections, correction previews, warning and
Review Notes handling, preservation of source-displayed test periods, and
approval checks for repeated lookup details. Further work could add a broader
automated audit against instructor-approved expected results, test more unusual
worksheet layouts, and publish the app only after choosing an approved hosting
and privacy approach. The app has not been published.
