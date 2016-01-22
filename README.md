# PoSh-SecretSanta
A PowerShell Script to generate Secret Santa Pairings.

Currently achieved through using the messy "Send-SecretSanta" function.

Will be replaced by smaller functions dedicated to single actions, such as:

Get-SantaPairs - which takes a list of strings and creates pairs (where one item cannot be paired with itself)

Send-SantaList - Sends an email to the "Santa" telling them about thier "Recipient"

Things to do:
  1) Add a way to add someone's name next to their email address ...maybe using a CSV instead
