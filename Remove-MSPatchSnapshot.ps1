# Script      : Remove-MSPatchSnapshot.ps1
# Author      : Burak Vural
# Description : Lists patch snapshots matching pattern and retention.
# Notes       : Public version. Environment specific values are anonymized.

# *********************************************************************
#
# Powershell script that deletes MSPatch snapshots
#
# USAGE: ./Remove-MSPatchSnapshot.ps1 -Retention <NumberOfDays>
#
# ASSUMPTIONS:
#
# *********************************************************************

param([Parameter(Mandatory=$true)][int32]$Retention)

$VMs = Get-VM
$Today = Get-Date
$SnapshotsToDelete = @()

foreach ($VM in $VMs) {
	$Snapshots = Get-Snapshot -VM $VM
	if (!$Snapshots) { continue }
	
	foreach ($Snapshot in $Snapshots) {
		if ($Snapshot.Name -match "smvi_4e23") { 
			$SnapshotCreateDate = $Snapshot.Created
			$DaysOld = $Today.Subtract($SnapshotCreateDate).Days
			if ($DaysOld -gt $Retention) { $SnapshotsToDelete += $Snapshot }
		}
	}
}

#$SnapshotsToDeletePartial = $SnapshotsToDelete | Select -First 20

#foreach ($SnapshotToDelete in $SnapshotsToDeletePartial) {
#	Remove-Snapshot -Snapshot $SnapshotToDelete -RunAsync -Confirm:$false
#	Start-Sleep -seconds 180 
#}

$SnapshotsToDelete | FT VM,Name,Created -AutoSize
$SnapshotsToDelete.Count
