#########
# au2mator PS Services
# New Service
# Active Directory - User Onboarding
# v 1.0 Initial Release
# Init Release: 22.08.2020
# Last Update: 22.08.2020
# Code Template V 1.2
# URL: 
# Github: 
# au2mator 4.x or above
#
##################



#region InputParamaters
##Question in au2mator
param (
    [parameter(Mandatory = $false)]
    [String]$c_Firstname,

    [parameter(Mandatory = $false)]
    [String]$c_LastName,

    [parameter(Mandatory = $false)]
    [String]$c_Description,

    [parameter(Mandatory = $false)]
    [String]$c_Office,

    [parameter(Mandatory = $false)]
    [String]$c_Telephone,


    [parameter(Mandatory = $false)]
    [String]$c_Mobile,

    [parameter(Mandatory = $false)]
    [String]$c_Fax,

    [parameter(Mandatory = $false)]
    [String]$c_JobTitle,

    [parameter(Mandatory = $false)]
    [String]$c_Department,

    [parameter(Mandatory = $false)]
    [String]$c_Location,

    [parameter(Mandatory = $false)]
    [String]$c_Manager,

    [parameter(Mandatory = $false)]
    [String]$c_Groups,


    # Office 365
    [parameter(Mandatory = $false)]
    [String]$c_MailboxLanguage,

    [parameter(Mandatory = $false)]
    [String]$c_Office365License,

    [parameter(Mandatory = $false)]
    [String]$c_MailDomain,




    ## au2mator Initialize Data
    [parameter(Mandatory = $true)]
    [String]$InitiatedBy,

    [parameter(Mandatory = $true)]
    [String]$RequestId,

    [parameter(Mandatory = $true)]
    [String]$Service,

    [parameter(Mandatory = $true)]
    [String]$TargetUserId
)
#endregion  InputParamaters




#region Variables
##Script Handling
$DoImportPSSession = $false
$ErrorCount = 0
#$ErrorActionPreference = "SilentlyContinue"

## Environment
[string]$DCServer = 'svdc01'
[string]$LogPath = "C:\_SCOworkingDir\TFS\PS-Services\AD - User Onboarding"
[string]$LogfileName = "User Onboarding"

[string]$CredentialStorePath = "C:\_SCOworkingDir\TFS\PS-Services\CredentialStore" #see for details: https://au2mator.com/documentation/powershell-credentials/


#Used PowerShell Modules
$Modules = @("ActiveDirectory")


## au2mator Settings
[string]$PortalURL = "http://demo01.au2mator.local"
[string]$au2matorDBServer = "demo01"
[string]$au2matorDBName = "au2matorNew"

## Control Mail
$SendMailToInitiatedByUser = $true #Send a Mail after Service is completed
$SendMailToTargetUser = $true #Send Mail to Target User after Service is completed

## SMTP Settings
$SMTPServer = "smtp.office365.com"
$SMPTAuthentication = $true #When True, User and Password needed
$EnableSSLforSMTP = $true
$SMTPSender = "michael.seidl@au2mator.com"
$SMTPPort="587"

$SMTPCredential_method = "Stored" #Stored, Manual
#Use stored Credentials
$SMTPcredential_File = "SMTPCreds.xml"
#Use Manual Credentials
$SMTPUser = ""
$SMTPPassword = ""

if ($SMTPCredential_method -eq "Stored") {
    $SMTPcredential = Import-CliXml -Path (Get-ChildItem -Path $CredentialStorePath -Filter $SMTPcredential_File).FullName
}

if ($SMTPCredential_method -eq "Manual") {
    $f_secpasswd = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
    $SMTPcredential = New-Object System.Management.Automation.PSCredential ($SMTPUser, $f_secpasswd)
}


#CustomVariables
$O365credential_File = "Office365Creds.xml"
$O365credential = Import-CliXml -Path (Get-ChildItem -Path $CredentialStorePath -Filter $O365credential_File).FullName
$Exchangecredential_File = "ExchangeCreds.xml"
$Exchangecredential = Import-CliXml -Path (Get-ChildItem -Path $CredentialStorePath -Filter $Exchangecredential_File).FullName



$ADPassword = "SuperS3cur3PW2020!!"
$AzureADSyncServer="rwsaad01.work.td.local"


#endregion Variables

#region Functions

function ConnectToDB {
    # define parameters
    param(
        [string]
        $servername,
        [string]
        $database
    )
    # create connection and save it as global variable
    $global:Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "server='$servername';database='$database';trusted_connection=false; integrated security='true'"
    $Connection.Open()
    Write-Verbose 'Connection established'
}
# function that executes sql commands against an existing Connection object; In pur case
# the connection object is saved by the ConnectToDB function as a global variable
function ExecuteSqlQuery {
    # define parameters
    param(

        [string]
        $sqlquery

    )

    Begin {
        If (!$Connection) {
            Throw "No connection to the database detected. Run command ConnectToDB first."
        }
        elseif ($Connection.State -eq 'Closed') {
            Write-Verbose 'Connection to the database is closed. Re-opening connection...'
            try {
                # if connection was closed (by an error in the previous script) then try reopen it for this query
                $Connection.Open()
            }
            catch {
                Write-Verbose "Error re-opening connection. Removing connection variable."
                Remove-Variable -Scope Global -Name Connection
                throw "Unable to re-open connection to the database. Please reconnect using the ConnectToDB commandlet. Error is $($_.exception)."
            }
        }
    }

    Process {
        #$Command = New-Object System.Data.SQLClient.SQLCommand
        $command = $Connection.CreateCommand()
        $command.CommandText = $sqlquery

        Write-Verbose "Running SQL query '$sqlquery'"
        try {
            $result = $command.ExecuteReader()
        }
        catch {
            $Connection.Close()
        }
        $Datatable = New-Object "System.Data.Datatable"
        $Datatable.Load($result)
        return $Datatable
    }
    End {
        Write-Verbose "Finished running SQL query."
    }
}

function Write-au2matorLog {
    [CmdletBinding()]
    param
    (
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$Type,
        [string]$Text
    )

    # Set logging path
    if (!(Test-Path -Path $logPath)) {
        try {
            $null = New-Item -Path $logPath -ItemType Directory
            Write-Verbose ("Path: ""{0}"" was created." -f $logPath)
        }
        catch {
            Write-Verbose ("Path: ""{0}"" couldn't be created." -f $logPath)
        }
    }
    else {
        Write-Verbose ("Path: ""{0}"" already exists." -f $logPath)
    }
    [string]$logFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogfileName
    $logEntry = '{0}: <{1}> <{2}> <{3}> {4}' -f $(Get-Date -Format dd.MM.yyyy-HH:mm:ss), $Type, $RequestId, $Service, $Text
    Add-Content -Path $logFile -Value $logEntry
}

function Get-UserInput ($RequestID) {
    [hashtable]$return = @{ }

    ConnectToDB -servername $au2matorDBServer -database $au2matorDBName

    $Result = ExecuteSqlQuery -sqlquery "SELECT        RPM.Text AS Question, RP.Value
    FROM            dbo.Requests AS R INNER JOIN
                             dbo.RunbookParameterMappings AS RPM ON R.ServiceId = RPM.ServiceId INNER JOIN
                             dbo.RequestParameters AS RP ON RPM.ParameterName = RP.[Key] AND R.RequestId = RP.RequestId
    where RP.RequestId = '$RequestID' order by [Order]"

    $html = "<table><tr><td><b>Question</b></td><td><b>Answer</b></td></tr>"
    $html = "<table>"
    foreach ($row in $Result) {
        $row
        $html += "<tr><td><b>" + $row.Question + "</b></td><td>" + $row.Value + "</td></tr>"
    }
    $html += "</table>"

    $f_RequestInfo = ExecuteSqlQuery -sqlquery "select InitiatedBy, TargetUserId,[ApprovedBy], [ApprovedTime], Comment from Requests where RequestId =  '$RequestID'"

    $Connection.Close()
    Remove-Variable -Scope Global -Name Connection

    $f_SamInitiatedBy = $f_RequestInfo.InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties Mail


    $f_SamTarget = $f_RequestInfo.TargetUserId.Split("\")[1]
    $f_UserTarget = Get-ADUser -Identity $f_SamTarget -Properties Mail

    $return.InitiatedBy = $f_RequestInfo.InitiatedBy
    $return.MailInitiatedBy = $f_UserInitiatedBy.mail
    $return.MailTarget = $f_UserTarget.mail
    $return.TargetUserId = $f_RequestInfo.TargetUserId
    $return.ApprovedBy = $f_RequestInfo.ApprovedBy
    $return.ApprovedTime = $f_RequestInfo.ApprovedTime
    $return.Comment = $f_RequestInfo.Comment
    $return.HTML = $HTML

    return $return
}

Function Get-MailContent ($RequestID, $RequestTitle, $EndDate, $TargetUserId, $InitiatedBy, $Status, $PortalURL, $RequestedBy, $AdditionalHTML, $InputHTML) {

    $f_RequestID = $RequestID
    $f_InitiatedBy = $InitiatedBy

    $f_RequestTitle = $RequestTitle
    $f_EndDate = $EndDate
    $f_RequestStatus = $Status
    $f_RequestLink = "$PortalURL/requeststatus?id=$RequestID"
    $f_RequestedBy = $RequestedBy
    $f_HTMLINFO = $AdditionalHTML
    $f_InputHTML = $InputHTML

    $f_SamInitiatedBy = $f_InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties DisplayName
    $f_DisplaynameInitiatedBy = $f_UserInitiatedBy.DisplayName


    $HTML = @'
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 1.5pt; background: #F7F8F3; mso-yfti-tbllook: 1184;" border="0" width="100%" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="padding: .75pt .75pt .75pt .75pt;" valign="top">&nbsp;</td>
    <td style="width: 450.0pt; padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top" width="600">
    <div style="box-sizing: border-box;">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: white; border: solid #E9E9E9 1.0pt; mso-border-alt: solid #E9E9E9 .75pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="1" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="border: none; background: #6ddc36; padding: 15.0pt 0cm 15.0pt 15.0pt;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><img src="https://au2mator.com/wp-content/uploads/2018/02/HPLogoau2mator-1.png" alt="" width="198" height="43" /></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="border: none; padding: 15.0pt 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 55.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="55%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="width: 18.75pt; border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="25">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm; font-color: #0000;"><strong>End Date</strong>: ##EndDate</td>
    </tr>
    <tr style="mso-yfti-irow: 1;">
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Status</strong>: ##Status</td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes;">
    <td style="border: solid #E3E3E3 1.0pt; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border: solid #E3E3E3 1.0pt; border-left: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Requested By</strong>: ##RequestedBy</td>
    </tr>
    </tbody>
    </table>
    </td>
    <td style="width: 5.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="5%">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 9.0pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    <td style="width: 40.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="40%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #FAFAFA; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="width: 100.0%; border: solid #E3E3E3 1.0pt; mso-border-alt: solid #E3E3E3 .75pt; padding: 7.5pt 0cm 1.5pt 3.75pt;" width="100%">
    <p style="text-align: center;" align="center"><span style="font-size: 10.5pt; color: #959595;">Request ID</span></p>
    <p class="MsoNormal" style="text-align: center;" align="center">&nbsp;</p>
    <p style="text-align: center;" align="center"><u><span style="font-size: 12.0pt; color: black;"><a href="##RequestLink"><span style="color: black;">##REQUESTID</span></a></span></u></p>
    <p class="MsoNormal" style="text-align: center;" align="center"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><strong><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">Dear ##UserDisplayname,</span></strong></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">the Request <strong>"##RequestTitle"</strong> has been finished.<br /> <br /> Please see the description for detailed information.<br /><b>##HTMLINFO</b><br /></span></p>
    <div>&nbsp;</div>
    <div>See the Details of the Request</div>
    <div>##InputHTML</div>
    <div>&nbsp;</div>
    <div>&nbsp;</div>
    Kind regards,<br /> au2mator Self Service Team
    <p>&nbsp;</p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';"><a style="border-radius: 3px; -webkit-border-radius: 3px; -moz-border-radius: 3px; display: inline-block;" href="##RequestLink"><strong><span style="color: white; border: solid #50D691 6.0pt; padding: 0cm; background: #50D691; text-decoration: none; text-underline: none;">View your Request</span></strong></a></span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 3; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #333333; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 50.0%; border: none; border-right: solid lightgrey 1.0pt; mso-border-right-alt: solid lightgrey .75pt; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    <td style="width: 50.0%; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </div>
    </td>
    <td style="padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    <p class="MsoNormal"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
'@

    $html = $html.replace('##REQUESTID', $f_RequestID).replace('##UserDisplayname', $f_DisplaynameInitiatedBy).replace('##RequestTitle', $f_RequestTitle).replace('##EndDate', $f_EndDate).replace('##Status', $f_RequestStatus).replace('##RequestedBy', $f_RequestedBy).replace('##HTMLINFO', $f_HTMLINFO).replace('##InputHTML', $f_InputHTML).replace('##RequestLink', $f_RequestLink)

    return $html
}


function Send-ServiceMail ($HTMLBody, $ServiceName, $Recipient, $RequestID, $RequestStatus) {
    $f_Subject = "au2mator - $ServiceName Request [$RequestID] - $RequestStatus"

    if ($SMPTAuthentication) {
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $SMTPcredential -UseSs -Port $SMTPPort
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $SMTPcredential -Port $SMTPPort
        }
    }
    else {
        if ($EnableSSLforSMTP) {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -UseSsl -Port $SMTPPort
        }
        else {
            Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Port $SMTPPort
        }
    }
}
#endregion Functions


#region Script


Write-au2matorLog -Type INFO -Text "Start Script"
if ($DoImportPSSession) {

    Write-au2matorLog -Type INFO -Text "Import-Pssession"
    $PSSession = New-PSSession -ComputerName $DCServer
    Import-PSSession -Session $PSSession -DisableNameChecking -AllowClobber
}
else {

}

Write-au2matorLog -Type INFO -Text "Import AD PS Module"
Import-Module ActiveDirectory


Write-au2matorLog -Type INFO -Text "Define all Values"

$ADFirstName = $c_Firstname
$ADLastName = $c_LastName
$ADName = "$c_LastName $c_Firstname"
$ADDisplayName = "$c_LastName $c_Firstname"
$ADDescription = "$c_Description"
$ADOffice = $c_Office
$ADTelephone = $c_Telephone

#MailFormat
$ADMail = "$c_Firstname.$c_LastName" + "@" + "$c_MailDomain"

$ADUPN = $ADMail





#Clean Username
$prestring = @'
ß;ö;ü;ä;à;á;é;è;ó;ò;ú;ù;í;ì;
'@;
$poststring = @'
ss;oe;ue;ae;a;a;e;e;o;o;u;u;i;i;
'@;
$delimeter = ';';
[array]$prestrings = $prestring -split $delimeter;
[array]$poststrings = $poststring -split $delimeter;
$ADUsername = $c_LastName.ToLower()
for ($i = 0; $i -le $prestrings.count; $i++) {
    $ADUsername = $ADUsername -replace ($prestrings[$i], $poststrings[$i]);
}
if ($ADUsername.Length -gt 10) {
    $ADUsername = $ADUsername.Substring(0, 10)
}

$ADHomePhone = $c_HomePhone
$ADPager = $c_Pager
$ADMobile = $c_Mobile
$ADFax = $c_Fax
$ADIPPhone = $c_IPPhone

$ADJobTitle = $c_JobTitle


$ADDepartment = $c_Department

$ADStreet=(Get-ADOrganizationalUnit -Identity $c_Location).StreetAddress
$ADZip=(Get-ADOrganizationalUnit -Identity $c_Location).Postalcode
$ADCity=(Get-ADOrganizationalUnit -Identity $c_Location).Name
$ADCountry=(Get-ADOrganizationalUnit -Identity $c_Location).country




$ADCompany = "au2mator.com"
$ADManager = $c_Manager
$ADGroups = $c_Groups.Split(";")
$ADDestOU = $DestOU

Write-au2matorLog -Type INFO -Text "Try to create the  User"

try {
    $NewUser = New-ADUser -UserPrincipalName $ADUPN -Path $c_Location -Country $ADCountry -GivenName $ADFirstName -Surname $ADLastName -SamAccountName $ADUsername -Name $ADName -DisplayName $ADDisplayName -Description $ADDescription -AccountPassword (ConvertTo-SecureString $ADPassword -AsPlainText -force) -Enabled $true -PassThru 
    Write-au2matorLog -Type INFO -Text "New User created"

    try {
        $NewUser | Set-ADUser -Office $ADOffice
        $NewUser | Set-ADUser -StreetAddress $ADStreet 
        $NewUser | Set-ADUser -PostalCode $ADZip 
        $NewUser | Set-ADUser -City $ADCity 
        
        
        Write-au2matorLog -Type INFO -Text "User updated with Adresse Details"


        try {
            $NewUser | Set-ADUser -fax $ADFax 
            $NewUser | Set-ADUser -MobilePhone $ADMobile 
            $NewUser | Set-ADUser -OfficePhone $ADTelephone 
            
            Write-au2matorLog -Type INFO -Text "User updated with Phone Details"

            try {
                $NewUser | Set-ADUser -Title $ADJobTitle 
                $NewUser | Set-ADUser -Department $ADDepartment 
                $NewUser | Set-ADUser -Company $ADCompany 
                $NewUser | Set-ADUser -Manager $ADManager 
                $NewUser | Set-ADUser -EmailAddress $ADMail
                Write-au2matorLog -Type INFO -Text "User updated with Advanced Details"

                try {
                    foreach ($Adgroup in $ADGroups) {
                        Write-au2matorLog -Type INFO -Text "Try to add User to Group: $Adgroup"
                        Add-ADGroupMember -Identity $Adgroup -Members $NewUser
                    }



                    if ($c_Office365License -eq "Office 365 E3") {
                        Write-au2matorLog -Type INFO -Text "Try to add User to Group: LIC-O365_E3"
                        Add-ADGroupMember -Identity "LIC-O365_E3" -Members $NewUser
                    }

                    if ($c_Office365License -eq "Office 365 F3") {
                        Write-au2matorLog -Type INFO -Text "Try to add User to Group: LIC-O365_F3"
                        Add-ADGroupMember -Identity "LIC-O365_F3" -Members $NewUser
                    }


                
                    try {
                        Write-au2matorLog -Type INFO -Text "run Azure Sync"
                        $ADsession=New-PSSession -ComputerName $AzureADSyncServer -Credential $O365credential -Authentication kerberos 
                        
                        Import-PSSession -Session $ADsession -DisableNameChecking -CommandName Start-ADSyncSyncCycle
                        
                        Start-ADSyncSyncCycle -PolicyType Delta

                        Write-au2matorLog -Type INFO -Text "Wait 60 Seconds"
                        Start-Sleep -Seconds 60
                        
                        try {
                            Write-au2matorLog -Type INFO -Text "run Office 365 Commands"

                            #Import-Module AzureAd
                            #Connect-AzureAD -Credential $O365credential
                            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365credential -Authentication Basic -AllowRedirection
                            Import-PSSession $Session -DisableNameChecking

                            Write-au2matorLog -Type INFO -Text "Wait for the Mailbox"

                            $i = 0
                            do {
                                if (Get-Mailbox -Identity $ADUPN) {
                                    Write-au2matorLog -Type INFO -Text "Mailbox found and synced"
                                    
                                    Set-MailboxRegionalConfiguration -Identity $ADUPN -Language $c_MailboxLanguage  -LocalizeDefaultFolderName:$true -Confirm:$false -TimeFormat $null -DateFormat $Null
                                    
                                    
                                    $i = 31
                                }
                                else {
                                    Start-Sleep -Seconds 30
                                    Write-au2matorLog -Type INFO -Text "Sleep 30 seconds"
                                    $i++
                                }
                            } until ($i -gt 30)

                        }
                        catch {
                            Write-au2matorLog -Type ERROR -Text "failed to set Office 365 Properties: $ADUPN"
                            $ErrorCount = 1
                            $AdditionalHTML = "<br>failed to set Office 365 Properties: $ADUPN<br>"
                        }


                    }
                    catch {
                        Write-au2matorLog -Type ERROR -Text "failed to run Azure AD Sync: $ADUPN"
                        $ErrorCount = 1
                        $AdditionalHTML = "<br>failed to run Azure AD Sync: $ADUPN<br>"
                    }


                }
                catch {
                    Write-au2matorLog -Type ERROR -Text "failed to add User to Group: $Adgroup"
                    $ErrorCount = 1
                    $AdditionalHTML = "<br>failed to add User to Group: $Adgroup<br>"
                }

            }
            catch {
                Write-au2matorLog -Type ERROR -Text "Failed to update Advanced Details"
                $ErrorCount = 1
                $AdditionalHTML = "<br>Failed to update Advanced Details<br>"
            }

        }
        catch {
            Write-au2matorLog -Type INFO -Text "Failed to update Phone Details"
            $ErrorCount = 1
            $AdditionalHTML = "<br>Failed to update Phone Details<br>"
        }

    }
    catch {
        Write-au2matorLog -Type ERROR -Text "Failed to update User Adresse Details"
        $ErrorCount = 1
        $AdditionalHTML = "<br>Failed to update User Adresse Details<br>"
    }


}
catch {
    Write-au2matorLog -Type ERROR -Text "Failed to create User"
    $ErrorCount = 1
    $AdditionalHTML = "<br>Failed to create User<br>"
}



if ($ErrorCount -eq 0) {
    $au2matorReturn = "User " + (Get-ADUser -identity $NewUser).DisplayName + "wurde erstellt"
    $AdditionalHTML = "<br>
    User " + (Get-ADUser -identity $NewUser).DisplayName + "wurde erstellt <br>
    Username: " + (Get-ADUser -identity $NewUser).UserPrincipalName + "
    <br>
    "
    $Status = "COMPLETED"
}
else {
    $au2matorReturn = "Fehler bei der Erstellung des Users, Error: $Error"
    $Status = "ERROR"
}
#endregion Script


#region Return
## return to au2mator Services



Write-au2matorLog -Type INFO -Text "Service finished"



if ($SendMailToInitiatedByUser) {
    Write-au2matorLog -Type INFO -Text "Send Mail to Initiated By User"

    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($UserInput.MailInitiatedBy)" -RequestStatus $Status -ServiceName $Service
}


if ($SendMailToTargetUser) {
    Write-au2matorLog -Type INFO -Text "Send Mail to Target User"

    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $TargetUserId -InitiatedBy $InitiatedBy -Status $Status -PortalURL $PortalURL -RequestedBy $InitiatedBy -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient "$($UserInput.MailTarget)" -RequestStatus $Status -ServiceName $Service
}


return $au2matorReturn
#endregion Return


