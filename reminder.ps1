

$startDate = [datetime]::Today
#Is the test-mode on or off?
$evalMode = $False
#To who the test-mails should go?
$testAdress = "sometestingadress@a.de"
#Define HelpDesk Adress
$helpDeskAdress = "helpdesk@somecompany.de"
#Define OWA Adress
$owaAdress = "https://remote.adsade.de/owa/"
#Define the SMTP Server
$SMTPServer = '10.1.10.30'
#Define Email-Sender
$mailSender = "Passwort Reminder <passwortreminder@rfsmediagroup.de>"
$subject = "Passwort Reminder"
#Define Email-Text
$line1 = "Hallo"
$line2 = "`n`nihr Password läuft in "
$line3 = "Tagen ab.`n"
$line4 = "`nBitte ändern Sie ihr Password rechtzeitig durch STRG + ALT + ENTF oder über " + $owaAdress + "`n`nBei fragen wenden Sie sich bitte an das HelpDesk: " + $helpDeskAdress
$line5 = "`nDies ist eine automatisch generierte Mail, auf diese zu antworten ist nicht notwendig."
#On which days left a user should be informed that his password experies
$WarningDays = @(5, 9, 20)
#Always warn if time is lower than
$WarnAlwaysOnRemainingTime = 3

    $users = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0} -Properties "Name", "mail", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "Name", "mail", "msDS-UserPasswordExpiryTimeComputed"
    foreach($user in $users) {
    try {
    if($user.Name -ne "DORTMUND$") {
        $endDate = [datetime]::FromFileTime($user.'msDS-UserPasswordExpiryTimeComputed')
        $daysLeft = (New-TimeSpan -start $startDate -end $endDate).days
        if($daysLeft -gt 0) {
            if($WarningDays.Contains($daysLeft)) {
                sendMail $daysLeft $user
            } elseif($daysLeft -le $WarnAlwaysOnRemainingTime) {
                sendMail $daysLeft $user
            }
        } 
    }
    } catch {
        $user
        $_
    }

}

function sendMail($days, $user) {
    if($evalMode -eq $True) {
        $EmailBody = $line1, $user.name + ",", $line2, $days, $line3, $line4, $line5 -join ' '
        Send-MailMessage -To $testAdress -From $mailSender -SmtpServer $SMTPServer -Subject $subject -Body $EmailBody -Encoding UTF8
        $user.Name + " | " + $days
    } else {
        $EmailBody = $line1, $user.name + ",", $line2, $days, $line3, $line4, $line5 -join ' '
        Send-MailMessage -To $user.mail -From $mailSender -SmtpServer $SMTPServer -Subject $subject -Body $EmailBody -Encoding UTF8
        $user.Name + " | " + $days

         if(0 -eq $days) {
            $remindForAdminsLine1 = "Hallo "
            $remindForAdminsLine2 = "`n`n, das Passwort von "
            $remindForAdminsLine3 = "ist heute abgelaufen`n"
            $remindForAdminsEmailBody = $remindForAdminsLine1, $remindForAdminsLine2, $user.Name, $remindForAdminsLine3 -join ' '
            Send-MailMessage -To $helpDeskAdress -From $mailSender -SmtpServer $SMTPServer -Subject $subject -Body $EmailBody -Encoding UTF8
        }
    }
}



EXIT
