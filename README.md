# CwaMarking

A script and Excel macro to help speed up marking for CWA courses.

## Before Marking
- Place the **Feedback.xlsm** **script.py** and **script.bat** files in the directory of your students marked work
- Open the **Feedback.xlsm** file
- Setup the **Marksheet** sheet as the template for your feedback
- Setup the **Namesheet** so column A contains all your student details (Name and ID)
  - You can get the student details via the student portal
    - Select the course
    - Filter by active students
    - Export Summary Grids
    - Learner Summary
- Press the **Submit** button to generate individual Marksheets for students (It is recommended not to change the sheet names) 
- Mark the learners work

## After Marking
- Open **Command Prompt** or **Windows PowerShell**
- `cd` to the directory of **script.bat**
- run **script.bat** by typing `script.bat`
  - The batch script will create a virtual environment and install the dependancies needed, **this requires a compatible version of Python 3** and then it will run the **script.py** file
- Follow the instructions in the program to either:
  - Rename Folders (Remove _assignsubmssion_file)
  - Export all Worksheets as .pdfs
