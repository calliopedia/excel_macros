This VBA macro automates the weekly update of an Excel issue-tracking workbook.

Main Functions
Selects an already-open source workbook.
Filters relevant issues using configurable criteria.
Copies or refreshes the latest weekly worksheet.
Updates existing issues while preserving manual information.
Adds new issues at the top of the tracker.
Removes duplicates and prioritizes selected records.
Applies configurable status colors.
Updates tracking dates and aging information.
Displays a summary after completion.
Configuration
Project filters and status colors are maintained in the Parameters sheet, allowing users to adjust the rules without changing the VBA code.

Weekly columns are identified by their headers, so their position may change without affecting the macro.

Usage
Open the source workbook in the same Excel session.
Click UPDATE WEEKLY FILE on the Instructions sheet.
Select the source workbook.
Review the completion message.
The source workbook remains unchanged, and existing manual tracking data is preserved.
