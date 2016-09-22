# PoSh-SecretSanta
A PowerShell Script to make sorting out Secret Santa a lot easier for teams that aren't always in the same place.

This all came about after "drawing out of a hat" took way too long because people where here, there and everywhere and then 3 times in a row the last person ended up drawing themselves!

It takes in a list of names and email addresses, and then mixes up the whole list and creates random pairings.  Then it sends an email to the "Santa" telling them the name of the person they need to buy for, with a nice reminder of what the budget is.

# There are 3 ways to use this:

1) Completely Graphical
	
	Just run .\Send-SecretSanta.ps1 and the form will guide you through what you need to do.
	
2) Partially Graphical
	
	You can pre-load some elements using the parameters.
	
3) Completely Commandline

	You can completely by-pass the GUI bits if you've got a CSV file ready to go.  
	Just make sure all the parameters are set first and use the "noGui" switch.