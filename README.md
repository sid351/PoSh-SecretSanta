# PoSh-SecretSanta
A PowerShell Script to load 2 functions used to help organise Secret Santa.

**Get-SantaPairs** 

Takes a list of strings and creates pairs (where one item cannot be paired with itself)

Best used if the strings are in the "Firstname Lastname <email@address.com>" format, as the email the "Santa" gets looks like this:

"Your Secret Santa Recipient is $Recipient.

The budget is $Budget.

Please don't reply to this message!" 

**Send-SantaList** 

Sends an email to the "Santa" telling them about thier "Recipient"
