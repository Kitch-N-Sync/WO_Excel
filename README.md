# WO_Excel
Excel Doc used to categorize Amazon maintenance techs' daily work orders and send a passdown email through Outlook.

EAM setup
Status | 1-WO Owner | Work Order | Alias | Description | PM Compliance Min | PM Compliance Max | Sched. Start Date | Sched. End Date | Trouble Ticket No. | Trade

Run your PM Compliance Min search and open that as an Excel file.
Run your Sched. Start Date search and open that as a separate Excel File.

Merge the two files, remove duplicates, and change the WO format to *number*.
Select only the cells with data, NOT to include the header line.
Copy these to the WO file.
Erase all data from the *Status* column.

Hidden sheet BTS should have lists that you can add to freely, no need to edit names for columns.

If you want to reorder how things are listed in the main spreadsheet, you will need to edit the VBA code accordingly.
