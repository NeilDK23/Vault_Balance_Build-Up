# Fireblocks Vault Asset Build-Up

A Python-based workbook generator for reconciling Fireblocks vault transaction exports into a structured Excel asset build-up file.

This project combines Fireblocks source and destination transaction exports, filters and reshapes the data, builds per-asset balance tabs, and produces a reconciliation sheet against a Fireblocks vault account report.

## How to use

1. Navigate to https://console.fireblocks.io/v2/ -> Transaction History -> Filters. 
  Filter Source -> Select specific Vault and download CSV report
  Filter Destination -> Select same Vault and download CSV report

  Paste the CSV files int he "Vault CSV" folder

2. Naviagate to https://console.fireblocks.io/v2/accounts/vault and download the Vault  Account Report in CSV format

Paste the CSV file in the Recon Vault Report folder.

3. Run the python script, checking the recon in the output file in the Output folder for accuracy of the build-up.

## Features

- Combines Fireblocks source and destination CSV exports into a single workbook
- Filters transaction data to relevant assets:
  - `TRX`
  - `ETH`
  - `USDT_ERC20`
  - `TRX_USDT_S2UZ`
- Creates a consolidated transaction sheet
- Creates dedicated asset sheets for:
  - `TRX`
  - `ETH`
  - `USDT`
- Calculates running balances using Excel formulas
- Includes gas fee logic for `TRX` and `ETH`
- Builds a `Recon` sheet to compare asset build-up balances to the Fireblocks vault account report
- Applies workbook formatting, filters, colours, and fixed-width columns for reporting use

## Project Structure

```text
Fireblocks Vault Build-Up/
├── fireblocks_vault_to_excel.py
├── .gitignore
├── Vault CSV/
├── Recon Vault Report/
├── Output/
└── Backups/
