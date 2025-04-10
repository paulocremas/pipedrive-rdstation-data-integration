# Google Sheets Data Integration Automation

This project automates the integration and synchronization of CRM data from two different Google Sheets, updated daily via Google Apps Script.

## Overview

The goal is to consolidate transaction data from **Pipedrive** and **RD Station**—made available through an API consumer configured by the client using **Kondado**—into a single final spreadsheet.

## Structure

### Spreadsheet 1
- **Sheet: Transactions**  
  Contains detailed information about transactions, each identified by a `lead_id`.
- **Sheet: Email-Transaction Mapping**  
  Maps `email` addresses to their respective `lead_id`.

### Spreadsheet 2
- Contains detailed information for all users (`email`, `name`, etc.).

##  How It Works

The script runs **automatically every day** and follows these steps:

1. **Extract Email-Transaction Pairs**  
   Reads new data from the **Email-Transaction Mapping** sheet (Spreadsheet 1) and builds a list of valid `(email, lead_id)` pairs.

2. **Filter Valid Users**  
   Filters out entries where the `email` does not exist in the user database (Spreadsheet 2).

3. **Merge Data**  
   For each valid pair:
   - Gets transaction details from the **Transactions** sheet (Spreadsheet 1).
   - Gets user details from Spreadsheet 2.  
   Merges both and appends to the **Final Spreadsheet**. User info may be repeated if linked to multiple transactions.

4. **Update Transaction Status**  
   Finds transactions marked as "completed" but not yet labeled as `closed` in the Final Spreadsheet. Updates their status accordingly.

## Automation

This workflow is powered by **Google Apps Script**, running on a time-based trigger for daily execution.

## Data Sources

- CRM 1: **Pipedrive**
- CRM 2: **RD Station**
- API Consumer: **KONDADO**

## Final Output

A consolidated Google Sheet combining:
- Transaction details
- User information
- Status tracking
