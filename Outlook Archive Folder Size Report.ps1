$mailboxes = @(Get-EXOMailbox -ResultSize Unlimited)
$report = @()

foreach ($mailbox in $mailboxes)
{
    $inboxstats = Get-MailboxFolderStatistics $mailbox.UserPrincipalName -FolderScope Archive | ? {$_.FolderPath -eq "/Archive"}

    $mbObj = New-Object PSObject
    $mbObj | Add-Member -MemberType NoteProperty -Name "UPN" -Value $mailbox.UserPrincipalName
    $mbObj | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $mailbox.DisplayName
    $mbObj | Add-Member -MemberType NoteProperty -Name "Folder" -Value $inboxstats.FolderPath
    $mbObj | Add-Member -MemberType NoteProperty -Name "Folder Size (MB)" -Value $inboxstats.FolderandSubFolderSize
    $mbObj | Add-Member -MemberType NoteProperty -Name "Number of Items" -Value $inboxstats.ItemsinFolderandSubfolders
    $report += $mbObj
}

$report | Export-CSV c:\temp\allarchivefoldersizereport.csv
