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

Function Send-SantaList
{
[cmdletbinding()]
Param(
    [parameter(Mandatory=$true)][string]$Santa,
        #Email Address for the "Santa"
    [parameter(Mandatory=$true)][string]$Recipient,
        #Email Address for the "Recipient"
    [parameter(Mandatory=$true)][string]$SmtpServer,
        #The SMTP Server address (IP or hostname) to use to send the email
    [string]$Budget = "£5",
        #The Budget for Secret Santa
    [string]$emailFrom = "Secret Santa <noReply@domain.com>",
        #The email address to send the notifications FROM
    [switch]$UseSSL,
    [int]$Port = 25,
        #The SMTP port to use
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
    Get-SantaPairs (Get-Content C:\Temp\Names.txt) | % {Send-SantaList -Santa $_.Santa -Recipient $_.Recipient -SmtpServer "mail.host.com"}
    
#>