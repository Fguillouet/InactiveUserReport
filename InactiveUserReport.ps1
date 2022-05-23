#Connexion au Tenant O365
Connect-MsolService
Connect-ExchangeOnline

#Récupération de la liste des utilisateurs et filtrage des résultats
Get-MsolUser | ForEach-Object {
    $upn = $_.UserPrincipalName
    $LastLogonTime = (Get-MailboxStatistics -Identity $upn).lastlogontime
    $DisplayName = $_.DisplayName
    $IsLicensed = $_.IsLicensed
    $MBUserCount++
    Write-Progress -Activity "`n     Processed mailbox count: $MBUserCount "`n"  Currently Processing: $DisplayName"

    [PSCustomObject]@{
        Name = $DisplayName
        Mail = $upn
        LastLogon = $LastLogonTime 
        IsLicensed = $IsLicensed
        } | Export-Csv -path "C:\Admin Rapport\Utilisateurs inactifs $(get-date -f dd-MM-yyyy).csv" -notype -Append 
}
Write-Output "Fin d'exécution de ce script, vous poouvez retrouver le rapport généré ici : C:\Admin Rapport\Utilisateurs inactifs $(get-date -f dd-MM-yyyy).csv"
Pause