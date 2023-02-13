# Connect to Exchange Online
$me = 'username@example.com' # UPN
Connect-ExchangeOnline -UserPrincipalName $me

# Variables
$days = '90' # Number of days w/o logon to return
$date = (Get-Date).AddDays(-$days) # Don't modify
$outPath = 'C:\SCRIPTS\OUTPUT' # Folder/directory only.
$outFile = 'AdUsersInactive90Days' # Do not include .csv
$properties = @('DistinguishedName','GivenName','Surname','Name','SamAccountName','UserPrincipalName','Description','LastLogonDate','Created','PasswordLastSet','AccountExpirationDate')
$filename = "$outPath\$outFile`_$(Get-Date -Format yyyyMMdd-HHmmss).csv"

# Email Variables
$SMTPServer = 'smtp.example.com'
$from = 'NoReply@exmaple.com'
$to = @('admin@example.com')
$subject = 'Inactive User Accounts'
$body = 'Review the list of user account showning no logins in 90+ days.'
$attachment = $filename

$users = Get-ADUser -Filter {Enabled -eq $true} -Properties $properties | Where-Object {$_.LastLogonDate -lt $date -and $_.Created -lt $date} | Select-Object -Property $properties

If($users){
    
    # Add a property to the users object for MailboxLastUserActionTime
    $users | Add-Member -NotePropertyName MailboxLastUserActionTime -NotePropertyValue ''

    foreach($user in $users){
        # Check O365 for last user action on the mailbox and add that value to the user object
        $MailboxStatistics = ''
        $LastActionTime = ''
        $MailboxStatistics = Get-MailboxStatistics -Identity $user.UserPrincipalName
        $LastActionTime = $MailboxStatistics.LastUserActionTime

        $user.MailboxLastUserActionTime = $LastActionTime

    }
    
    # Export CSV
    $users | Export-Csv -Path "$filename" -NoTypeInformation
    # Send Email
    Send-MailMessage -From $from -To $to -Subject $subject -Body "$body" -Attachments $filename -SmtpServer $SMTPServer
}
