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
    


    Let's make this a little bit more ...powershelly:

        1) Take in an array of Strings
        2) Randomise the array of Strings
        3) Loop through the strings and build pairs as an object
            Santa | Recipient

        Break it down in to each little bit
            1) Get list of names
            2) Make Pairings
            3) Send out notification
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

Function Get-SantaPairs
{
[cmdletbinding()]
Param(
    [parameter(ValueFromPipeline=$false,Mandatory=$true)][String[]]$SantaList
    )

    #Randomise the list
    $SantaList = get-random -input $SantaList -count $SantaList.Count

    $pairs = for($i=0; $i -lt $SantaList.Count; $i++)
    {
        if($i -eq $SantaList.Count-1)
        {
            #Set last person in the list to give to first person
            $strChoice = $SantaList[0]
        }
        else
        {
            #Give to the next person in the list
            $strChoice = $SantaList[$i + 1]
        }
        
        $pair = New-Object -TypeName PSCustomObject -Property @{
            Santa = $SantaList[$i]
            Recipient = $strChoice
            }

        Write-Output $pair
    }
    
    Write-Output $pairs
}
<#
.PARAMETER SantaList
    The whole array of Santa's at once.

    Santa's names can be provided in the "Firstname Surname <email@address.co.uk>" format.

    That makes it easier to understand who "123@abc.def" is when you get your Santa email.

.EXAMPLE
    Get-SantaPairs (Get-Content C:\Temp\Names.txt)

    Recipient             Santa                
    ---------             -----                
    Four <4@email.domai>  Eight <8@email.domai>
    Five <5@email.domai>  Four <4@email.domai> 
    Three <3@email.domai> Five <5@email.domai> 
    Two <2@email.domai>   Three <3@email.domai>
    Seven <7@email.domai> Two <2@email.domai>  
    Six <6@email.domai>   Seven <7@email.domai>
    One <1@email.domai>   Six <6@email.domai>  
    Nine <9@email.domai>  One <1@email.domai>  
    Ten <10@email.domai>  Nine <9@email.domai> 
    Eight <8@email.domai> Ten <10@email.domai> 

#>

function Send-SantaList
{
[cmdletbinding()]
Param(
    [parameter(Mandatory=$true)][string]$Santa,
        #Email Address for the "Santa"
    [parameter(Mandatory=$true)][string]$Recipient,
        #Email Address for the "Recipient"
    [parameter(Mandatory=$true)][string]$SmtpServer,
    [string]$Budget = "£5",
    [string]$emailFrom = "Secret Santa <noReply@domain.com>",
    [switch]$UseSSL,
    [int]$Port = 25,
    $Bcc,
    $Cc,
    $Attachments,
    $DeliveryNotificationOption,
    $Encoding,
    $Priority,
    $Credentials
    )

    $emailParam = @{
        To = $Santa
        Body = "Your Secret Santa Recipient is $Recipient.<br/> <br/>The budget is $Budget.<br/> <br/>Please don't reply to this message!" 
        BodyAsHtml = $true
        From = $emailFrom
        Subject = "Your Secret Santa Recipient is..." 
        SmtpServer = $SmtpServer
        UseSSL = $UseSSL
        Port = $Port
        }

    If($credentials -ne $null -and $credentials -ne "") { $emailParam.Credential = $credentials }
    If($Bcc -ne $null -and $Bcc -ne "") { $emailParam.Bcc = $Bcc }
    If($Cc -ne $null -and $Cc -ne "") { $emailParam.Cc = $Cc }
    If($Attachments -ne $null -and $Attachments -ne "") { $emailParam.Attachments = $Attachments }
    If($DeliveryNotificationOption -ne $null -and $DeliveryNotificationOption -ne "") { $emailParam.DeliveryNotificationOption = $DeliveryNotificationOption }
    If($Encoding -ne $null -and $Encoding -ne "") { $emailParam.Encoding = $Encoding }
    If($Priority -ne $null -and $Priority -ne "") { $emailParam.Priority = $Priority }
    
    send-mailmessage @emailParam 

}
<#
.DESCRIPTION
    Assumes all the "names" (Santa & Recipient) input are valid email addresses

.EXAMPLE
    Get-SantaPairs (Get-Content C:\Temp\Names.txt) | % {Send-SantaList -Santa $_.Santa -Recipient $_.Recipient}
    
#>



<#
    Take in CSV with the following fields:
        Name
        Email Address
        Last years recipient?

#>