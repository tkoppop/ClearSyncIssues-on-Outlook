$o = New-Object -comobject outlook.application
$n = $o.GetNamespace(“MAPI”)
$Account = $n.Folders | ? { $_.Name -eq "$env:UserName" + "@glotmansimpson.com" };

$TestFolder = $Account.Folders | ? {$_.Name -match 'Test'};

$SyncIssues = $Account.Folders | ? { $_.Name -match 'Sync Issues' };
$LocalFailures = $SyncIssues.Folders | ? { $_.Name -match 'Local Failures' };
$Conflicts = $SyncIssues.Folders | ? { $_.Name -match 'Conflicts' };
while ($SyncIssues.Items.Count -ne 0){
$SyncIssues.Items | % {$_.delete()}
}

while ($LocalFailures.Items.Count -ne 0){
$LocalFailures.Items | % {$_.delete()}
}

while ($Conflicts.Items.Count -ne 0){
$Conflicts.Items | % {$_.delete()}
}
