# Privacy Policy

**IsBilled – QuickBooks Invoice Lookup Tool**  
**Floweigh LLC**  
**Effective Date:** April 14, 2026

---

## Overview

IsBilled is an internal tool developed by Floweigh LLC for the purpose of looking up invoice status within the company's QuickBooks Online account. This policy describes what data is accessed, how it is used, and how it is protected.

## Data Accessed

IsBilled connects to Floweigh LLC's QuickBooks Online account via the Intuit API and retrieves the following data:

- Invoice records, including invoice number, date, total amount, and balance
- Customer name associated with each invoice
- Custom fields on invoices: Order Number, PO Number, and Quote Number

## How Data Is Used

Data retrieved from QuickBooks Online is used solely to display invoice lookup results to authorized Floweigh LLC employees within the application. No data is transmitted to any third party.

## Data Storage

IsBilled does **not** store, cache, or log any QuickBooks invoice or customer data. The application log (`qbo_lookup.log`) records only operational events such as search activity counts, API response status codes, and error messages. No invoice content, customer names, or financial figures are written to the log.

## Credentials

API credentials (Client ID, Client Secret, and OAuth tokens) are stored locally in a `.env` configuration file on Floweigh LLC's internal network drive. These credentials are never transmitted to any party other than Intuit's OAuth and API endpoints.

## Third-Party Services

IsBilled communicates exclusively with the following Intuit endpoints:

- `https://oauth.platform.intuit.com` — OAuth 2.0 token exchange
- `https://quickbooks.api.intuit.com` — QuickBooks Online REST API

No other third-party services are contacted.

## Access Control

Access to IsBilled is limited to authorized Floweigh LLC employees with access to the company's internal network drive. There is no public-facing interface.

## Changes to This Policy

This policy may be updated at any time. Continued use of the software constitutes acceptance of any revised policy.

---

*For questions, contact Floweigh LLC internally.*
