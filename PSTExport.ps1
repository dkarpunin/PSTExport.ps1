cls

$ErrorActionPreference="SilentlyContinue"

Write-Host "Экспорт п/я в PST `n" -ForegroundColor Green

$Email = Read-Host "Адрес п/я"

$Path="\\<shared_folder_path>\"+$Email+".pst"

$Domains = (Get-ADForest).Domains

foreach ($Domain in $Domains) {
   
    $DCs = (Get-ADDomain -Server $Domain).InfrastructureMaster

    foreach ($DC in $DCs) {
        
        $OriginatingServer = (Get-Mailbox -Identity $Email -DomainController $DC).OriginatingServer

        if($OriginatingServer) {$MailboxDC = $OriginatingServer}

    }

}

New-MailboxExportRequest -DomainController $MailboxDC -Mailbox $Email -FilePath $Path