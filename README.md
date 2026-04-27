# Azure Email Extractor

## Overview
A Python utility that connects to Microsoft Outlook using Azure app authentication and Microsoft Graph API to retrieve emails from a specific sender, extract relevant email data, and prepare it for export into a structured Excel format.

## Context
This program was developed as a component of a larger task management system, where email data can later be used as input for tracking tasks, announcements, or updates.

## My Role
- Built the Python email extraction program
- Integrated Azure authentication using MSAL
- Retrieved emails using Microsoft Graph API
- Parsed email HTML content using BeautifulSoup
- Extracted sender, subject, timestamp, and image data
- Structured the output using pandas for Excel export

## Features
- Microsoft login using Azure app authentication
- Token caching and interactive login flow
- Device code fallback authentication
- Email retrieval from a selected sender
- HTML parsing to extract embedded images
- Data structuring for Excel export

## Tech Stack
Python, Microsoft Graph API, Azure App Registration, MSAL, pandas, BeautifulSoup, Requests

## Security Note
Sensitive credentials such as `CLIENT_ID` are stored in environment variables using a `.env` file and should not be committed to the repository.

## Future Improvements
- Enable configurable sender filters
- Export directly to Excel
- Add support for pagination beyond the first 50 emails
- Integrate extracted emails into a task management dashboard
