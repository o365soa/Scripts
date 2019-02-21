# DMARC Reporting

The DMARC Reporting script is an example script for parsing DMARC reports, and storing parsed results in to an SQL table.

Parsed results can then be used by other programs, or displayed graphically such as with the included Power BI example.

## Requirements

1. SQL Table
2. PowerShell v4.0
3. DMARC Reports sent to an Office 365 mailbox (can be shared)
4. An application set up in Azure AD for the script to use, see below.

## Changelog

29/3/2017 - Updated to use MSAL libraries

## Provision AAD Application for DMARC Reporting Script
