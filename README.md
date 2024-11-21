# CwaMarking

A script to help speed up marking for CWA courses.
Built for compatibility with **Feedback.xlsm**

## Before Marking
- Open the **Feedback.xlsm** file
- Setup the **Marksheet** sheet as the template for your feedback
- Setup the **Namesheet** so column A contains all your student details (Name and ID)
  - You can get the student details via the student portal
    - Select the course
    - Filter by active students
    - Export Summary Grids
    - Learner Summary
- Press the **Submit** button to generate individual Marksheets for students
- Mark the learners work

## After Marking
- Run script.bat
  - You may need to unblock the file in the file's properties.
  - Otherwise you may need to run it via CMD: 
    - Open **Command Prompt** or **Windows PowerShell**
    - `cd` to the directory of **script.bat**
    - run **script.bat** by typing `script.bat`
- This may take a while as the batch script will create a virtual environment and install the dependancies needed, **this requires a compatible version of Python 3** and then it will run the **script.py** file
- Follow the instructions in the program:
  - Select a file to convert
  - Select a directory to save to
  - Select the sheets to convert
  - Click convert
  - Close program when done
