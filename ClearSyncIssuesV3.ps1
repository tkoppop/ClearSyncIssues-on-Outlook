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

while ($SyncIssues.Items.Count -ge 0){

try{
	For ($i = ($Count1);$i -ge 1;$i--) {
   		$SyncIssues.Remove($i)
	}
} catch {
break
}
if($Count1 -eq 0){

break
}
}

$Count2 = $Conflicts.Items().Count
$Count2
$Conflicts = $Conflicts.Items()

while ($Conflicts.Items.Count -ge 0){
try{
	For ($i = ($Count2);$i -ge 1;$i--) {
    		$Conflicts.Remove($i)
	}
} catch {
break
}
if($Count2 -eq 0){

break
}
}

$Count3 = $LocalFailures.Items().Count
$Count3 
$LocalFailures = $LocalFailures.Items()

while ($LocalFailures.Items.Count -ge 0){
try{
	For ($i = ($Count3);$i -ge 1;$i--) {
    		$LocalFailures.Remove($i)
	}
} catch {
break
}
if($Count3 -eq 0){

break
}
}
