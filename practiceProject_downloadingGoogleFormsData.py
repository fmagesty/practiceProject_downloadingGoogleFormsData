# Program that collects a list of email addresses on a spreadsheet.
# This specific .xlsx file was created from a google forms that would ask the name and e-mails of users.

import ezsheets

# Open the spreadsheet
ss = ezsheets.Spreadsheet('1UfzA0O5kr0S4Z-WkeCs0e5AQ_teJYO0emhJ4fk4Puyw')
print('Your file is located at %s' % (ss.url))
sheet = ss[0]
# Find the e-mails.
emails = sheet.getColumn('C')
# Extract the e-mails.
print('Extracting e-mails...')
for i in emails:
    if i in emails != None and len(str(i)) > 1:
        print(i)
print('Done.')