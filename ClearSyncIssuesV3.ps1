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
[int]$temp = 0
if ($Count1 -eq 0){
$temp = 1
}
while ($SyncIssues.Items.Count -ge 0){
 $global:temp++
try{
	For ($i = ($Count1);$i -ge 1;$i--) {
   		$SyncIssues.Remove($i)
	}
} catch {
break
}
if($temp -eq 1){

break
}
}
$temp = 0
$Count2 = $Conflicts.Items().Count
$Count2
$Conflicts = $Conflicts.Items()
if ($Count2 -eq 0){
$temp = 1
}
while ($Conflicts.Items.Count -ge 0){
try{
	For ($i = ($Count2);$i -ge 1;$i--) {
    		$Conflicts.Remove($i)
	}
} catch {
break
}
if($temp -eq 1){

break
}
}
$temp = 0
$Count3 = $LocalFailures.Items().Count
$Count3 
$LocalFailures = $LocalFailures.Items()
if ($Count3 -eq 0){
$temp = 1
}
while ($LocalFailures.Items.Count -ge 0){
try{
	For ($i = ($Count3);$i -ge 1;$i--) {
    		$LocalFailures.Remove($i)
	}
} catch {
break
}
if($temp -eq 1){

break
}
}
