# Secret Santa Google Sheets Automailer
# Author: Aidan Chang

This is a script I created in December 2021 for YoungLife 2021 Winter Secret Santa coordination. It was my first year in YoungLife and the first year we did a Secret Santa, so I wanted to apply my JavaScript knowledge to automate the randomization and emailing process. Selection of gift givers is not 100% random because some people signed up late. Personal information has been redacted for safety reasons.

# IMPLEMENTATION
I created a Google Form to handle user input and gift preferences, then stored it automatically a Google Sheet which is an equivalent to Microsoft Excel. I then created the driver for the auto mailer in Google Apps Script, a Google derivative of JavaScript that is found by going to "Extensions->Apps Script" from your Google Sheet. The script can be run on any Sheet of the same format.
