# Word_Automation

![Logo](https://miro.medium.com/v2/resize:fit:1400/1*Q5R55WBwqSKr87CuV4QiuQ.jpeg)

## This project automates the creation of student grade reports in Word format. It extracts data from an Excel file with grades for three subjects and fills a personalized letter template for each student, ensuring accuracy and efficiency.

## Features

- **Installing resources and libraries:** Installs the docxtpl and python-docx libraries required to work with Word document templates.
- **Upload Word template:** Upload the Word document template that will be used to generate the reports and define constant information that will be included in each report, such as name, phone number, email and date.
- **Read data from Excel file:** Read student data and grades from an Excel file for later extraction into the text document.
- **Generate and save individual reports:** For each student, extract their grades and generate a personalized report in Word format from a python loop that identifies each of the labels.
- **Save a test document:** Generate and save a test document with the constant information.

## Technologies used

- **Python:** Programming language used to write the automation script, providing the foundation for the project and enabling data manipulation and document generation.
- **Pandas:** Python library for data manipulation and analysis, used to read and handle data from the Excel file containing student grades.
- **Datetime:** Python module for working with dates and times, allowing the current date to be obtained and formatted for inclusion in the reports.
- **Docxtpl:** Python library for creating and manipulating Word documents using templates, facilitating the generation of personalized Word documents from a template and specific data.
- **Python-docx:** Python library for creating, modifying, and extracting information from Word documents, complementing docxtpl in manipulating Word documents, although docxtpl is the main tool used in this project.

## **Documentation**
! https://www.youtube.com/watch?v=8XXpheZ8uRs&t=533s
! https://www.youtube.com/watch?v=T4up6ePInXg
! https://medium.com/@ejeraldo/python-para-automatizar-excel-y-word-cf3c37c44c56
