# README

## Overview
This tool can automatically organize students' respondents from Moodle's Questionnaire into an Excel sheet.

**Please note:** This method is no longer functional due to Moodle's implementation of captcha verification.

## Usage Instructions

1. **Setup:**
   - Ensure you have Python installed on your machine.
   - Install the required packages by running:
     ```bash
     pip install selenium beautifulsoup4 openpyxl
     ```
   - Place this script in the desired directory where you want to save the output.

2. **Configuration:**
   - Open the script and locate line 131.
   - Enter the course URL on this line. **Example:** `website = "https://moodle.youruniversity.edu/course/view.php?id=123"`
   - On line 132, input your username and password for Moodle. 

3. **Execution:**
   - Run the script.
   - After the prompt "Please enter the week number you are looking for:", enter the title (usually week number) that you want to search for. 
   - Upon successful execution, an Excel file named `OUTPUT.xlsx` will be generated in the same directory as the script.
   - You can change the output file location by modifying line 90 in the script. For example:
     ```python
     save_to_excel(i, data, filename='Desktop\\OUTPUT.xlsx')
     ```

## Output
- The resulting file will contain the organized data from the respondents.

## Contact
For questions or support, please reach out to h54091122@gs.ncku.edu.tw.
