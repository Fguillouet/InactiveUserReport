Function Get-Folder($initialDirectory="")

{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Sélectionnez un dossier"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

#Réinitialisation de variables
$MBUserCount = 0

#Connexion au Tenant O365
Connect-MsolService
Connect-ExchangeOnline -ShowBanner:$false

#Récupérer le dossier de destination
$DestinationFolder = Get-Folder

Clear-Host

#Récupération de la liste des utilisateurs et filtrage des résultats
Get-MsolUser | ForEach-Object {

    #WIP : exclusion des comptes externes (mention #EXT#) de la liste à exporter 

    $upn = $_.UserPrincipalName
    $LastLogonTime = (Get-MailboxStatistics -Identity $upn -erroraction 'silentlycontinue').lastlogontime
    $DisplayName = $_.DisplayName
    $IsLicensed = $_.IsLicensed
    $MBUserCount++
    Write-Progress -Activity "Processed mailbox count: $MBUserCount "`n"Currently Processing: $DisplayName"

    [PSCustomObject]@{
        Name = $DisplayName
        Mail = $upn
        LastLogon = $LastLogonTime 
        IsLicensed = $IsLicensed
        } | Export-Csv -path $DestinationFolder"\Utilisateurs inactifs $(get-date -f dd-MM-yyyy).csv" -notype -Append 
}

Clear-Host

Write-Output "Fin d'execution de ce script, vous poouvez retrouver le rapport genere ici : "$DestinationFolder"\Utilisateurs inactifs $(get-date -f dd-MM-yyyy).csv"
Pause