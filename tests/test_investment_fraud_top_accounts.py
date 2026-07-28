from datetime import date
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook

from src.investment_fraud_top_accounts import (
    build_investment_fraud_report,
    generate_investment_fraud_excel,
    is_investment_fraud,
)


def test_is_investment_fraud_accepts_common_variations():
    assert is_investment_fraud("INVESTMENT FRAUD - trading app")
    assert is_investment_fraud("Online investmant scheme fraud")
    assert is_investment_fraud("Fraud through an investment platform")
    assert not is_investment_fraud("Job fraud")
    assert not is_investment_fraud("Investment advice only")


def test_build_investment_fraud_report_matches_aggregates_and_ranks_accounts():
    additional_df = pd.DataFrame(
        [
            {
                "Acknowledgement No.": "3110001",
                "Crime Aditional Information": "Investment fraud through trading app",
            },
            {
                "Acknowledgement No.": "3110002",
                "Crime Aditional Information": "INVESTMANT SCHEME FRAUD",
            },
            {
                "Acknowledgement No.": "3110003",
                "Crime Aditional Information": "Job fraud",
            },
        ]
    )
    layerwise_df = pd.DataFrame(
        [
            {
                "Acknowledgement No.": "3110001",
                "Account No.": "001234567890",
                "IFSC Code": "TEST0001",
                "Address": "First address",
                "District": "Ahmedabad",
                "State": "Gujarat",
                "Transaction Amount": "₹1,000.00",
                "Disputed Amount": "750",
                "Bank/FIs": "Example Bank",
                "Layers": "Layer 1",
            },
            {
                "Acknowledgement No.": "3110002",
                "Account No.": "001234567890",
                "IFSC Code": "",
                "Address": "",
                "District": "",
                "State": "",
                "Transaction Amount": 500,
                "Disputed Amount": 500,
                "Bank/FIs": "",
                "Layers": 1,
            },
            {
                "Acknowledgement No.": "3110002",
                "Account No.": "99887766",
                "IFSC Code": "SECOND001",
                "Address": "Second address",
                "District": "Surat",
                "State": "Gujarat",
                "Transaction Amount": 2_000,
                "Disputed Amount": 900,
                "Bank/FIs": "Second Bank Ltd",
                "Layers": "1",
            },
            {
                "Acknowledgement No.": "3110002",
                "Account No.": "WALLET-1",
                "IFSC Code": "",
                "Address": "",
                "District": "",
                "State": "",
                "Transaction Amount": 5_000,
                "Disputed Amount": 5_000,
                "Bank/FIs": "Example Wallet",
                "Layers": 1,
            },
            {
                "Acknowledgement No.": "3110001",
                "Account No.": "LAYER-2",
                "IFSC Code": "TEST0002",
                "Address": "",
                "District": "",
                "State": "",
                "Transaction Amount": 9_000,
                "Disputed Amount": 9_000,
                "Bank/FIs": "Example Bank",
                "Layers": 2,
            },
            {
                "Acknowledgement No.": "3110003",
                "Account No.": "OTHER-FRAUD",
                "IFSC Code": "TEST0003",
                "Address": "",
                "District": "",
                "State": "",
                "Transaction Amount": 8_000,
                "Disputed Amount": 8_000,
                "Bank/FIs": "Example Bank",
                "Layers": 1,
            },
        ]
    )

    report_df, summary = build_investment_fraud_report(additional_df, layerwise_df)

    assert report_df["Fraudster Bank Account Number"].tolist() == [
        "001234567890",
        "99887766",
    ]
    first_account = report_df.iloc[0]
    assert first_account["All Acknowledgement Numbers"] == "3110001;3110002"
    assert first_account["ACK Count"] == 2
    assert first_account["Total Transactions"] == 2
    assert first_account["Total Amount"] == 1_500
    assert first_account["Total Disputed Amount"] == 1_250
    assert first_account["IFSC Code"] == "TEST0001"
    assert summary == {
        "qualifying_acknowledgements": 2,
        "matched_layer_one_transactions": 4,
        "matched_accounts": 2,
        "excluded_non_bank_accounts": 1,
        "output_accounts": 2,
    }


def test_generate_investment_fraud_excel_preserves_identifiers_and_layout():
    report_df = pd.DataFrame(
        [
            {
                "Sr.No.": 1,
                "Fraudster Bank Account Number": "001234567890",
                "All Acknowledgement Numbers": "3110001;3110002",
                "ACK Count": 2,
                "Bank Name": "Example Bank",
                "IFSC Code": "TEST0001",
                "Address": "First address",
                "District": "Ahmedabad",
                "State": "Gujarat",
                "Total Transactions": 2,
                "Total Amount": 1_500.0,
                "Total Disputed Amount": 1_250.0,
            }
        ]
    )

    excel_bytes, filename = generate_investment_fraud_excel(
        report_df, report_date=date(2026, 7, 27)
    )
    workbook = load_workbook(BytesIO(excel_bytes), data_only=False)
    worksheet = workbook["Top 20 Suspect Accounts"]

    assert filename == (
        "27-07-2026 Investment Fraud Top 20 Suspect Accounts from Layer 1.xlsx"
    )
    assert worksheet["A1"].value == (
        "27-07-2026 Investment Fraud Top 20 Suspect Accounts from Layer 1"
    )
    assert worksheet["B4"].value == "001234567890"
    assert worksheet["B4"].number_format == "@"
    assert worksheet["B4"].quotePrefix is True
    assert worksheet["C4"].number_format == "@"
    assert worksheet["L4"].number_format == '"₹"#,##0.00'
    assert worksheet.freeze_panes == "A4"
    assert worksheet.auto_filter.ref == "A3:L4"
