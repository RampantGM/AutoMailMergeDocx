# AutoMailMergeDocx
Connect to an Oracle database to fill out MS Word Mailmerge Letter Templates. Can do a single account or upload a spreadsheet with account numbers.
Allows manual field population by identifying mergefields with a particular prefix.

I created this project in order to take control of a good portion of the letter templates being used in our office.
The reason for this is to manage the letter templates themselves to prevent accidental overwrites.
I also wanted to make it easier for our users to fill in the information in the templates.

There are currently multiple procedures for filling out the templates. Some do screen scrapes of our application on each account, saving
them to excel spreadsheets and importing them one by one into the template. This required a MS Word Macro to read the cells that the
screen scrape put the information into. Others are using search results from our application and exporting that information into excel
spreadsheets or access databases and actually using mailmerge the way it was intended.


