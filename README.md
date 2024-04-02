# Python_Project_Forecast_Flow_Report

This project was developed as part of a data engineering course, aiming optimize process within my workplace.

The report is designed to take a dataset of loans and generate future cash flows for them. A key feature of this logic allows for adjustments in the early repayment rate, facilitating faster or slower loan portfolio repayment according to our analytical needs.

The code is capable of processing a CSV file containing basic loan information and outputs the future loan flow in an Excel format, including a summary sheet that aggregates the entire flow.

The most challenging aspect of this project was aligning the business logic with the financial model. Upon testing against the business requirements, the financial model embedded within the report was validated for accuracy.

For generating around 1000 fictitious loan records for testing, I utilized the Mock Data website (https://mockaroo.com/).



# Prerequisites
- Python 3.11
- pip
- pandas
- numpy_financial
- calendar
- datetime
- xlsxwriter

# For an overview and Documentation
- Enter the notebook called  Forecast_Flow_Report

# Installing & Running 
- Clone the repository 
- Navigate to the project directory
- Install the required packages using pip: "pip install -r requirements.txt"
- exac "Run.py" with the params or enter "NoteBookToRun" notebook
- Please note, after you run the program there is an option to download Excel to your computer

This is what the final result looks like in Excel:

<img width="446" alt="image" src="https://github.com/anatzaf/Python_Project_Forecast_Flow_Report/assets/157733416/14ee6742-d209-4b4b-bf9c-cf12bb45948a">
