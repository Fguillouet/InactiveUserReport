#Réinitialisation de variables
$MBUserCount = 0

#Connexion au Tenant O365
Connect-MsolService
Connect-ExchangeOnline -ShowBanner:$false

Clear-Host

#Récupération de la liste des utilisateurs et filtrage des résultats
Get-MsolUser | ForEach-Object {

    #WIP : exclusion des comptes externes (mention #EXT#) de la liste à exporter 

    $upn = $_.UserPrincipalName
    $LastLogonTime = (Get-MailboxStatistics -Identity $upn -erroraction 'silentlycontinue').lastlogontime
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
Write-Output "Fin d'execution de ce script, vous poouvez retrouver le rapport genere ici : C:\Admin Rapport\Utilisateurs inactifs $(get-date -f dd-MM-yyyy).csv"
Pause