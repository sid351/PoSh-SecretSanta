Function Send-SecretSanta
{
<#

.Notes
    Accepts email addresses in and sends an email to each secret santa telling them the email address they need to buy for.

    TO DO:
        Sort out the Help bit
        Properly parameterise the whole function
        Reduce the number of arrays used (can easily scrap one of them, could probably get rid of 2 of the 3)
        Replace "testing" with WhatIf functionality
        Add Write-verbose functionality
        Add error handling (Try/Catch)
        Make a "generate list" function so full names can co-exist with email addresses


.Example
    Send-SecretSanta -testing -cmd "Name.Surname@domain.co.uk","Name@domain.co.uk","Name@domain.co.uk" -SmtpServer "email.domain.co.uk" -EmailFrom "Secret.Santa@domain.co.uk"

    Will output the randomised list back to the console, to act as a sanity test before sending out emails.

    NOTE: SmtpServer & EmailFrom are not actually used when using the Testing flag but a value is required
    These will be removed in future using parameter sets
    
#>

Param(
[array]$cmd,
$file,
[Parameter(Mandatory=$true)][string]$SmtpServer,
[Parameter(Mandatory=$true)][string]$EmailFrom,
$Credentials,
[string]$Port,
[switch]$testing
)

    $emailParam =@{
         From = $EmailFrom
         Subject = "Your Secret Santa Recipient is..." 
         SmtpServer = $SmtpServer
         #UseSSL = $true
         #Port = $Port
         #Credential = $credentials
    }

    $arrEntries = New-Object System.Collections.ArrayList
    $arrSantas = New-Object System.Collections.ArrayList
    $arrPairs = New-Object System.Collections.ArrayList

    If($cmd -eq $null)
    {
        If($file -eq $null)
        {
            #CMD and FILE are both null
            write-host "Please provide an input." -foregroundcolor "red"
            #exit
        }
        else
        {
            #CMD is null, use FILE
            $arrEntries = get-content $file
        }
    }
    else
    {
        #Default to CMD
        $arrEntries = $cmd
    }
    If($cmd -ne $null -and $cmd[0].contains("\"))
    {
        #CMD probably contains a file path
        write-host "To use a file as input please use the -file switch" -foregroundcolor "red"
        #exit
    }

    #Randomise the list
    $arrSantas = get-random -input $arrEntries -count $arrEntries.Count

    for($i=0; $i -lt $arrSantas.Count; $i++)
    {
        $strSanta = $arrSantas[$i]
        if($i -eq $arrSantas.Count-1)
        {
            #Set last person in the list to give to first person
            $strChoice = $arrSantas[0]
        }
        else
        {
            #Give to the next person in the list
            $strChoice = $arrSantas[$i + 1]
        }
        set $arrPairs.add("$strSanta*$strChoice")
    }

    Foreach ($pair in $arrPairs) 
    {
        #break apart the pairing into giftGiver and giftReciever
        $giftGiver = $pair.split("*")[0]
        $giftReceiver = $pair.split("*")[1]
        #$giftReceiver = $giftReceiver.split("@")[0]
        #$giftReceiver = $giftReceiver.split(".")[0]
    
        #write-host "Send email to: $giftGiver telling them to buy for $giftReciever"
    
        if(!$testing)
        {
            #Send the email
            send-mailmessage @emailParam -To $giftGiver -Body "Your Secret Santa Recipient is $giftReceiver.<br/> <br/>  Please don't reply to this message!" -BodyAsHtml
        }
        else
        {
            Write-Output "Giver: $giftGiver; Recipient: $giftReceiver"
        }
    } 
}