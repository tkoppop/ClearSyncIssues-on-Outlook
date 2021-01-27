$o = New-Object -comobject outlook.application
$n = $o.GetNamespace(“MAPI”)
$Account = $n.Folders | ? { $_.Name -eq "$env:UserName" + "@glotmansimpson.com" };

$Folder = $Account.Folders | ? {$_.Name -match 'Test'};

$SyncIssues = $Account.Folders | ? { $_.Name -match 'Sync Issues' };
$LocalFailures = $SyncIssues.Folders | ? { $_.Name -match 'Local Failures' };
$Conflicts = $SyncIssues.Folders | ? { $_.Name -match 'Conflicts' };

$Count1 = $SyncIssues.Items().Count
$Count1
$SyncIssues = $SyncIssues.Items()

For ($i = ($Count1);$i -ge 1;$i--) {
    $SyncIssues.Remove($i)
}

$Count2 = $Conflicts.Items().Count
$Count2
$Conflicts = $Conflicts.Items()

For ($i = ($Count2);$i -ge 1;$i--) {
    $Conflicts.Remove($i)
}

$Count3 = $LocalFailures.Items().Count
$Count3 
$LocalFailures = $LocalFailures.Items()

For ($i = ($Count3);$i -ge 1;$i--) {
    $LocalFailures.Remove($i)
}

