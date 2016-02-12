# google-sheets-form-copier
Google Apps Script to copy calculations (as they are inserted, not entered)

Google Sheets, in conjunction with Google Forms are great tools for recordng and analysing data.

However, I recently realised that (as one might expect), when a new form entry is added, the entry is inserted. So if you have calculations in a separate sheet, it is difficult to copy these down in a simple way.

Thus, I decided to create a script that did the following:

1: Check the responses (Source) sheet for new entries
2: Verify this number against the number of calculations in the calculations (Targe) sheet.
3: Copy a certain (user defined) row of data / formulae down the sheet for the applicable number of new responses.

These settings are managed in a custom sidebar, and the script is run by a "Refresh" option in a menu.

This certainly has helped with managing form responses when I've used them, so I hope it helps others.
