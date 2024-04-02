# Python_Project_Forecast_Flow_Report

This project was created as part of the data engineering course, with the aim of optimizing processes at my workplace.

The report knows how to take a data set of loans, and generate a future flow for them.

As part of the logic in the report, it is possible to play with the early repayment rate so that the loan portfolio is repaid faster or slower - depending on what we want to examine.

The code knows how to receive a CSV file of basic loan data, and return an output of the loan flow forward in Excel, including a sheet that summarizes the entire flow

The hardest part of the project was matching the business logic and the financial model. In the tests against the business side, the financial model in the report was found to be correct.

I used the Mock Data website to generate about 1000 fictitious loan records
https://mockaroo.com/



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
