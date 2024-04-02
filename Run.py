

import pandas as pd
import numpy_financial as npf
from datetime import datetime
import calendar
import xlsxwriter
import argparse
import sys
from google.colab import drive,files


if 'ipykernel' in sys.modules:
    #if in notebook
    rate_of_early_repayments = 0.025
    fixed_fee = 24
    current_month = datetime(2024, 2, 29)

else:
    parser = argparse.ArgumentParser(description="Financial parameters.")

    parser.add_argument('--rate_of_early_repayments', type=float, default=0.025,
                        help='Rate of early repayments')
    parser.add_argument('--fixed_fee', type=int, default=24,
                        help='Fixed fee')
    parser.add_argument('--current_month', type=lambda s: datetime.strptime(s, '%Y-%m-%d'),
                        default=datetime(2024, 2, 29),
                        help='Current month (format YYYY-MM-DD)')
    args = parser.parse_args()

    rate_of_early_repayments = args.rate_of_early_repayments
    fixed_fee = args.fixed_fee
    current_month = args.current_month



def calculate_monthly_payment(am_EstimatedBalance, annual_interest_rate, nb_TotalPayments):
    monthly_interest_rate = annual_interest_rate / 12
    return npf.pmt(monthly_interest_rate, nb_TotalPayments, -am_EstimatedBalance)

def calculate_principal_payment(am_EstimatedBalance, annual_interest_rate, month, nb_TotalPayments):
    monthly_interest_rate = annual_interest_rate / 12
    remaining_payments = nb_TotalPayments - month + 1
    monthly_payment = calculate_monthly_payment(am_EstimatedBalance, annual_interest_rate, remaining_payments)
    interest_payment = am_EstimatedBalance * monthly_interest_rate
    return monthly_payment - interest_payment

def calculate_early_repayment(previous_balance, principal_payment):
    return (previous_balance - principal_payment) * rate_of_early_repayments

def calculate_EarlyRepaymentFee(payment_number, previous_balance, principal_component, nb_TotalPayments, annual_interest_rate, rate_of_early_repayments, fixed_fee):
    monthly_interest_rate = annual_interest_rate / 12
    adjusted_balance = previous_balance - principal_component
    EarlyRepaymentFee = 0
    for future_month in range(1, 7):
        adjusted_nper = nb_TotalPayments - payment_number
        if future_month <= adjusted_nper:
            interest_payment = npf.ipmt(monthly_interest_rate, future_month, adjusted_nper, -adjusted_balance)
            EarlyRepaymentFee += interest_payment
    EarlyRepaymentFee = (EarlyRepaymentFee * rate_of_early_repayments) + (fixed_fee / nb_TotalPayments)
    return EarlyRepaymentFee

def calculate_loan_repayments(loan_id,am_EstimatedBalance, annual_interest_rate, nb_TotalPayments,rate_of_early_repayments, fixed_fee):
    payment_records = []
    previous_balance = am_EstimatedBalance
    for month in range(1, nb_TotalPayments + 1):
        principal_payment = calculate_principal_payment(previous_balance, annual_interest_rate, month, nb_TotalPayments)
        interest_payment = previous_balance * (annual_interest_rate / 12)
        early_repayment = calculate_early_repayment(previous_balance, principal_payment)
        balance = previous_balance - (principal_payment + early_repayment)
        if month == nb_TotalPayments:
            EarlyRepaymentFee = 0
        else:
            EarlyRepaymentFee = calculate_EarlyRepaymentFee(month, previous_balance, principal_payment, nb_TotalPayments, annual_interest_rate, rate_of_early_repayments, fixed_fee)

        payment_record = {
            'Month': month,
            'Balance': balance,
            'Principal Component': principal_payment,
            'Interest Component': interest_payment,
            'Expected Sum of Early Repayments': early_repayment,
            'Early Repayment Fee': EarlyRepaymentFee,
        }
        payment_records.append(payment_record)
        previous_balance = balance
    return payment_records



def csv_drive_path_generatoer(url):
 path = 'https://drive.google.com/uc?export=download&id='+url.split('/')[-2]
 return path


def export_data_to_excel(link,excel_file_path, rate_of_early_repayments, fixed_fee, current_month):

    path = csv_drive_path_generatoer(url = link)
    loan_data_df = pd.read_csv(path)
    writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')
    workbook = writer.book

    components = ['Principal Component', 'Interest Component', 'Early Repayment Fee']
    max_payments = loan_data_df['nb_TotalPayments'].max()


    number_format = workbook.add_format({'num_format': '#,##0.00'})
    header_format = workbook.add_format({'bg_color': '#ADD8E6', 'bold': True})
    title_format = workbook.add_format({'bold': True, 'color': 'blue'})
    loan_id_format = workbook.add_format()  # Default format for loan ID


    payment_numbers = [f'Month {i+1}' for i in range(max_payments)]
    month_end_dates = [(current_month + pd.offsets.MonthEnd(n)).strftime('%Y-%m-%d') for n in range(max_payments)]


    summary_data = pd.DataFrame(index=month_end_dates, columns=components).fillna(0)

    for component in components:
        worksheet = workbook.add_worksheet(component)
        worksheet.right_to_left()

        # Apply header format for the first two rows
        worksheet.write_row(0, 1, payment_numbers, header_format)
        worksheet.write_row(1, 1, month_end_dates, header_format)
        worksheet.write('A1', 'Loan ID', header_format)
        writer.sheets[component] = worksheet

    for index, row in loan_data_df.iterrows():
        loan_id = row['nk_Deal']
        am_EstimatedBalance = row['am_EstimatedBalance']
        annual_interest_rate = row['annual_interest_rate']
        nb_TotalPayments = row['nb_TotalPayments']
        schedule = calculate_loan_repayments(loan_id, am_EstimatedBalance, annual_interest_rate, nb_TotalPayments, rate_of_early_repayments, fixed_fee)


        for component in components:
            worksheet = writer.sheets[component]
            worksheet.write(index + 2, 0, loan_id, loan_id_format)  # Write Loan ID in each sheet
            for month_idx, data in enumerate(schedule, start=1):
                value = data[component]
                worksheet.write(index + 2, month_idx, value, number_format)
                # Update summary data
                summary_data.loc[month_end_dates[month_idx-1], component] += value


    summary_sheet = workbook.add_worksheet('Summary')
    summary_sheet.right_to_left()
    summary_sheet.write_row('B1', components, title_format)
    for i, date in enumerate(summary_data.index):
        summary_sheet.write(i + 2, 0, date, title_format)
        for j, component in enumerate(components):
            summary_sheet.write(i + 2, j + 1, summary_data.at[date, component], number_format)

    writer.close()



export_data_to_excel(
    link = "https://drive.google.com/file/d/1MKRuDwZTLByN2MO6aT1pciKnMVE4e6rg/view?usp=drive_link",
    excel_file_path = 'split_loan_data.xlsx',
    rate_of_early_repayments=rate_of_early_repayments,
    fixed_fee=fixed_fee,
    current_month=current_month
)


files.download('split_loan_data.xlsx')

