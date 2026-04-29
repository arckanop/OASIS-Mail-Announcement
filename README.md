# Organtogo - OASIS Mail Announcement

This is the source code for the OASIS Email Announcement system for Organtogo Thailand.

The system sends personalized HTML email announcements to selected teams for the **Organ Ambassador Student Innovation Scheme: OASIS**. It pulls team data from a Google Sheet, generates a styled email, sends it through Gmail, and records the sending status back into the sheet.

## Features

- Sends personalized OASIS announcement emails
- Supports teams with 3–5 members
- Dynamically generates member rows from Google Sheets data
- Includes both HTML email and plain-text fallback
- Records email sending status in the sheet
- Uses email-safe table layouts for better Gmail compatibility
- Escapes user-provided text to prevent broken HTML

## Tech Stack

- Google Apps Script
- Google Sheets
- GmailApp
- MailApp
- HTML email with inline CSS

## Google Sheet Format

The script expects the data to be in a sheet named `Finalists`.

## License

AGPL-3.0

## Project

Created for Organtogo Thailand.