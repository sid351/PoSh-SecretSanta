<#
.NAME 
Send-SecretSanta

.SYNOPSIS
A tool to generate Secret Santa pairings and email out the results to the individual Santas.

.DESCRIPTION
This can be run as-is which will launch a Graphical Tool to help you construct a list of "Santas" that will then be shuffled and have pairings picked.

Each "Santa" will then get an email sent to them telling them who they are buying for with a reminder of the Budget.
    
As such, you will need access to a SMTP server in order to send the emails out.

Through the GUI it's possible to add and remove people from the list.
    
It's possible to import a CSV list, but that CSV must include a "Name" and "Email" column for the import to work.  The Santas imported from the CSV will be added to the same list as those input manually (and can also be removed).

As published, there is no logging of the pairings so everything remains a surprise for as long as the Santa's keep quiet!

There is one small limitation of this approach and it's that it's not possible for a pairing to get paired with each other.

For example, the randomised list is: Adam, Bob, Charlie, Donna, Erica & Fran.

The pairings are selected like this:

    Adam <------------------------------------
       -> Bob                                |
            -> Charlie                       |
                     -> Donna                |
                            -> Erica         |
                                   -> Fran ---

So it is not possible for Donna to be Erica's Santa AND Erica to be Donna's Santa (unless there are only 2 people in the list, or people swap a recipients).

.NOTES
This multipart blog was incredibly helpful for learning how to craft a GUI for PowerShell scripts:
Part 1 = https://foxdeploy.com/2015/04/10/part-i-creating-powershell-guis-in-minutes-using-visual-studio-a-new-hope/
Part 2 = https://foxdeploy.com/2015/04/16/part-ii-deploying-powershell-guis-in-minutes-using-visual-studio/

.PARAMETER csvFile
Pre-loads the CSV input file before the GUI loads.

The CSV must contain a "Name" column and an "Email" column.  It may contain others but they will not be used.

.EXAMPLE Send-SecretSanta -csvFile C:\Names.csv -fromEmailAddress "SecretSanta@consto.com" -budget "£15" -smtpPort 25 -useSSL -smtpServer "smtp.consto.com" -smtpCredential (Get-Credential) -noGui

This will run without needing to launch the GUI at all.  It is not possible to interact with the Santa list in this mode.

.EXAMPLE Send-SecretSanta

Launches the GUI to allow you to build up the Santa list, enter all the email details and send out the pairings.
#>

[cmdletbinding(DefaultParameterSetName = 'Gui')]
Param(
    [Parameter(ParameterSetName = 'Gui', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NoGui', Mandatory = $true)]
    [ValidateScript({<# CsvFile must be a valid CSV file that already exists #> ($_.ToLower().EndsWith(".csv")) -and (Test-Path -Path $_ -PathType Leaf -IsValid)})] 
        [string]$CsvFile,
    
    [Parameter(ParameterSetName = 'Gui', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NoGui', Mandatory = $true)]
    [ValidateScript({<# FromEmailAddress must be a valid email address, but does not need to be a live mailbox #> ($_ -as [System.Net.Mail.MailAddress])})]
        [string]$FromEmailAddress = "Secret Santa <noreply@domain.com>",
            #Pre-load the "From" email address.  It must be in the form of a valid email address, but does not have to be a live mailbox depending on your SMTP Relay settings.
    
    [Parameter(ParameterSetName = 'Gui', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NoGui', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
        [string]$Budget = "£10",
            #Pre-load the Budget
    
    [Parameter(ParameterSetName = 'Gui', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NoGui', Mandatory = $true)]
        [string]$SmtpServer,
            #Pre-load the SMTP Server address
    
    [Parameter(ParameterSetName = 'Gui', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NoGui', Mandatory = $true)]
    [ValidateRange(1, 65535)]
        [int]$SmtpPort = 25,
            #Pre-load the SMTP port
    
    [Parameter(ParameterSetName = 'Gui', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NoGui', Mandatory = $false)]
        [switch]$UseSSL,
            #Pre-load the UseSSL checkbox
    
    [Parameter(ParameterSetName = 'Gui', Mandatory = $false)]
    [Parameter(ParameterSetName = 'NoGui', Mandatory = $true)]
        [System.Management.Automation.PSCredential]$SmtpCredential,
            #Pre-load the SMTP Username and Password
    
    [Parameter(ParameterSetName = 'NoGui', Mandatory = $true)]
        [switch]$NoGui
            #Take all the fun out of it and don't bother with the GUI :P.  All other parameters must be provided in order to use this method.
)

Function Get-FormVariables
{
[cmdletbinding()]
Param()

    if ($global:ReadmeDisplay -ne $true)
    {
        Write-Verbose "If you need to reference this display again, run Get-FormVariables"
        $global:ReadmeDisplay=$true
    }
    
    Write-Verbose "Found the following interactable elements from our form"

    If($VerbosePreference -eq "Continue")
    {
        Get-Variable WPF*
    }
}

$inputXML = @"
<Window x:Name="Secret_Santa_Sender" x:Class="Secret_Santa_Sender.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Secret_Santa_Sender"
        mc:Ignorable="d"
        Title="Secret Santa Sender" Height="650" Width="408.667">
    <Grid Margin="2">
        <Button x:Name="btnAddToList" Content="Add to list" HorizontalAlignment="Left" Margin="308,182,0,0" VerticalAlignment="Top" Width="74"/>
        <TextBox x:Name="txtInName" HorizontalAlignment="Left" Height="26" Margin="102,122,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="280" VerticalContentAlignment="Center"/>
        <TextBox x:Name="txtInEmail" HorizontalAlignment="Left" Height="26" Margin="102,152,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="280" VerticalContentAlignment="Center"/>
        <Label x:Name="lblHeader" Content="Add extra Santas to the list:" HorizontalAlignment="Left" Margin="10,92,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblName" Content="Name:" HorizontalAlignment="Left" Margin="10,122,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblEmail" Content="Email Address:" HorizontalAlignment="Left" Margin="10,152,0,0" VerticalAlignment="Top"/>
        <Separator HorizontalAlignment="Left" Height="20" Margin="10,76,0,0" VerticalAlignment="Top" Width="372"/>
        <ListView x:Name="listSantas" HorizontalAlignment="Left" Height="168" Margin="10,251,0,0" VerticalAlignment="Top" Width="372">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding Name}" Width="150"/>
                    <GridViewColumn Header="Email Address" DisplayMemberBinding ="{Binding Email}" Width="222"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Label x:Name="lblList" Content="Santa List:" HorizontalAlignment="Left" Margin="10,220,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblFromEmail" Content="Emails &quot;From&quot; Address:" HorizontalAlignment="Left" Margin="11,426,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblBudget" Content="Budget:" HorizontalAlignment="Left" Margin="11,456,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblServer" Content="Smtp Server:" HorizontalAlignment="Left" Margin="11,485,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblUsername" Content="Smtp Username:" HorizontalAlignment="Left" Margin="11,513,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblPassword" Content="Smtp Password:" HorizontalAlignment="Left" Margin="11,541,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblPort" Content="Smtp Port:" HorizontalAlignment="Left" Margin="148,456,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtInFrom" HorizontalAlignment="Left" Height="26" Margin="148,426,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="235" VerticalContentAlignment="Center"/>
        <TextBox x:Name="txtInBudget" HorizontalAlignment="Left" Height="26" Margin="67,456,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="74" VerticalContentAlignment="Center"/>
        <TextBox x:Name="txtInSmtpServer" HorizontalAlignment="Left" Height="26" Margin="148,485,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="235" VerticalContentAlignment="Center"/>
        <TextBox x:Name="txtInSmtpUsername" HorizontalAlignment="Left" Height="26" Margin="148,513,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="235" VerticalContentAlignment="Center"/>
        <TextBox x:Name="txtInSmtpPort" HorizontalAlignment="Left" Height="26" Margin="218,456,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="99" VerticalContentAlignment="Center"/>
        <CheckBox x:Name="boxUseSSL" Content="Use SSL" HorizontalAlignment="Left" Margin="322,456,0,0" VerticalAlignment="Top" Height="26" VerticalContentAlignment="Center"/>
        <PasswordBox x:Name="passInSmtpPassword" HorizontalAlignment="Left" Margin="148,541,0,0" VerticalAlignment="Top" Width="235" Height="26" VerticalContentAlignment="Center"/>
        <Button x:Name="btnSendNow" Content="Send Now!" HorizontalAlignment="Left" Margin="11,572,0,0" VerticalAlignment="Top" Width="372" RenderTransformOrigin="0.5,0.5" Height="34">
            <Button.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="-0.123"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Button.RenderTransform>
        </Button>
        <Button x:Name="btnRemoveSelected" Content="Remove Selected" HorizontalAlignment="Left" Margin="270,220,0,0" VerticalAlignment="Top" Width="111" Height="26"/>
        <Label x:Name="lblImportCSV" Content="Import a list of Santa's from a CSV file:" HorizontalAlignment="Left" Margin="10,2,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtCsvPath" HorizontalAlignment="Left" Height="23" Margin="10,33,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="314" VerticalContentAlignment="Center"/>
        <Button x:Name="btnImportCsv" Content="Import CSV File" HorizontalAlignment="Left" Margin="270,61,0,0" VerticalAlignment="Top" Width="112"/>
        <Separator HorizontalAlignment="Left" Height="20" Margin="10,412,0,0" VerticalAlignment="Top" Width="372"/>
        <Separator HorizontalAlignment="Left" Height="20" Margin="10,202,0,0" VerticalAlignment="Top" Width="372"/>
        <Button x:Name="btnBrowse" Content="Browse" HorizontalAlignment="Left" Margin="329,33,0,0" VerticalAlignment="Top" Width="53" Height="23"/>
    </Grid>
</Window>
"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try
{
    $Form=[Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xaml))
}
catch
{
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
}

#Load Windows Forms for Browse button for searching for files
Add-Type -AssemblyName System.windows.forms | Out-Null
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | ForEach-Object -Process { Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) }
 
Get-FormVariables
 
#===========================================================================
# Actually make the objects work
#===========================================================================

$WPFbtnBrowse.Add_Click({

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.ShowDialog()
    $WPFtxtCsvPath.Text = $OpenFileDialog.FileName

})

$WPFbtnImportCsv.Add_Click({
    
    #Check that the file exists
    If(Test-Path -Path $WPFtxtCsvPath.Text -PathType Leaf -IsValid)
    {
        Foreach($santa in (Import-Csv -Path $WPFtxtCsvPath.Text))
        {   
            $obj = New-Object -TypeName PSCustomObject -Property @{
                Name = $santa.Name
                Email = $santa.Email
                }

            $WPFlistSantas.AddChild($obj)
        }
    }
    
})

$WPFbtnAddToList.Add_Click({

    If(-not ([string]::IsNullOrEmpty($WPFtxtInName.Text) -and [string]::IsNullOrEmpty($WPFtxtInEmail.Text)))
    {
        $obj = New-Object -TypeName PSCustomObject -Property @{
            Name = $WPFtxtInName.Text
            Email = $WPFtxtInEmail.Text
            }

        $WPFlistSantas.AddChild($obj)
    
        $WPFtxtInName.Text = ""
        $WPFtxtInEmail.Text = ""
    }

})

$WPFbtnRemoveSelected.Add_Click({

    $WPFlistSantas.Items.Remove($WPFlistSantas.SelectedItem)

})

$WPFbtnSendNow.Add_Click({

    #Randomise the input list
    
    [array]$SantaList = ($WPFlistSantas.Items | Get-Random -Count $WPFlistSantas.Items.Count)
    
    For($i=0; $i -lt $SantaList.Count; $i++)
    {
        If($i -eq $SantaList.Count-1)
        {
            #The last person in the random list "picks" the first person in the list
            $Recipient = $SantaList[0]
        }
        Else
        {
            #This person "picks" the next person in the random list
            $Recipient = $SantaList[$i + 1]
        }

        $output = New-Object -TypeName PSCustomObject -Property @{
            SantaName = $SantaList[$i].Name
            SantaEmail = $SantaList[$i].Email
            RecipientName = $Recipient.Name
            RecipientEmail = $Recipient.Email
            }

        $paramEmail = @{
            Subject = "Your secret santa recipient is..."
            Body = "Hi $($output.SantaName)! You're buying a Secret Santa present for $($output.RecipientName).  Please remember the budget is only $($WPFtxtInBudget.Text)."    
            To = $output.SantaEmail
            From = $WPFtxtInFrom.Text
            SmtpServer = $WPFtxtInSmtpServer.Text
            Port = $WPFtxtInSmtpPort.Text
            BodyAsHtml = $true
            UseSSL = $WPFboxUseSSL.IsChecked
            }

        If(-not ([string]::IsNullOrEmpty($WPFtxtInSmtpUsername.Text) -and [string]::IsNullOrEmpty($WPFpassInSmtpPassword.Password)))
        {
            $paramEmail.Credential = (New-Object System.Management.Automation.PSCredential ($WPFtxtInSmtpUsername.Text, $(ConvertTo-SecureString -String $WPFpassInSmtpPassword.Password -AsPlainText -Force)))
        }

        Send-MailMessage @paramEmail

    }

    $Form.Close()

})

#Pre-populate if Params have been defined
If(-not [string]::IsNullOrEmpty($csvFile))
{ 
    $WPFtxtCsvPath.Text = $csvFile 
    $WPFbtnImportCsv.RaiseEvent([System.Windows.RoutedEventArgs]::New([System.Windows.Controls.Button]::ClickEvent))
}

If(-not [string]::IsNullOrEmpty($budget)){ $WPFtxtInBudget.Text = $budget }

If(-not [string]::IsNullOrEmpty($fromEmailAddress)){ $WPFtxtInFrom.Text = $fromEmailAddress }

If(-not [string]::IsNullOrEmpty($smtpServer)){ $WPFtxtInSmtpServer.Text = $smtpServer }
If(-not [string]::IsNullOrEmpty($smtpPort)){ $WPFtxtInSmtpPort.Text = $smtpPort }
If(-not [string]::IsNullOrEmpty($smtpCredential))
{ 
    $WPFtxtInSmtpUsername.Text = $smtpCredential.UserName
    $WPFpassInSmtpPassword.Password = $smtpCredential.GetNetworkCredential().Password
}
$WPFboxUseSSL.IsChecked = $useSSL

#===========================================================================
# Shows the form
#===========================================================================

If(-not $noGui)
{
    $Form.ShowDialog() | out-null
}
Else
{
    #Force the items to be called via code alone
    $WPFbtnSendNow.RaiseEvent([System.Windows.RoutedEventArgs]::New([System.Windows.Controls.Button]::ClickEvent))
}

#===========================================================================
# Post submit processing
#===========================================================================