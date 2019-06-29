########################################################################################################
#Purpose: Script to obtain the disk space on remote servers
#Pre-Reqs: 
#  WMI should be enabled on all target servers
#  WMI ports to all target servers, should be opened from the location where the script is run  
#  Excel should be installed in the location where the script is run
########################################################################################################

import-module activedirectory
$erroractionpreference = “SilentlyContinue” 
$a = New-Object -comobject Excel.Application 
$a.visible = $True



$b = $a.Workbooks.Add() 
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = “Machine Name” 
$c.Cells.Item(1,2) = “Drive” 
$c.Cells.Item(1,3) = “Cluster/Local” 
$c.Cells.Item(1,4) = “Total size (GB)” 
$c.Cells.Item(1,5) = “Free Space (GB)” 
$c.Cells.Item(1,6) = “Free Space (%)” 
$c.cells.item(1,7) = "Name "
$c.cells.item(1,8) = "DriveType"
$d = $c.UsedRange 
$d.Interior.ColorIndex = 19 
$d.Font.ColorIndex = 11 
$d.Font.Bold = $True 
$d.EntireColumn.AutoFit()

$intRow = 2

$cred= get-credential

$clusterArray = New-Object System.Collections.ArrayList
$clusterHashTable = New-Object System.Collections.Hashtable

##Provide location of file with list of all target servers to check
$Computers = get-content "computers.txt"
  
foreach ($Computer in $Computers) 
{ 
    ##For cluster
    $clustername= gwmi MSCluster_Cluster -Namespace root/mscluster -ComputerName $Computer -Credential $cred|select name
    if($clustername)
    {
        $clusterDrives = New-Object System.Collections.ArrayList
        ##see if this cluster is processed
        if($clusterArray -notcontains $clustername.name){            
            ##Add to cluster array
            $clusterArray.Add($clustername.name)
            $ClusterDisks=gwmi MSCluster_DiskPartition -Namespace root/mscluster -ComputerName $Computer -Credential $cred| Select Path, TotalSize, Freespace, Volumelabel
            foreach ($objdisk in $ClusterDisks) 
            { 
                $clusterDrives.Add($objDisk.Path)
                $c.Cells.Item($intRow, 1) = $clustername.Name.ToUpper() 
                $c.Cells.Item($intRow, 2) = $objDisk.Path
                $c.Cells.Item($intRow, 3) = "Cluster"
                $c.Cells.Item($intRow, 4) = “{0:N0}” -f ($objDisk.TotalSize/1024) 
                $c.Cells.Item($intRow, 5) = “{0:N0}” -f ($objDisk.FreeSpace/1024) 
                $c.Cells.Item($intRow, 6) = “{0:P0}” -f ([double]$objDisk.FreeSpace/[double]$objDisk.TotalSize) 
                $c.cells.item($introw, 7) = $objdisk.Volumelabel
                $c.cells.item($introw, 8) = 3

                $intRow = $intRow + 1 
            }

            $clusterHashTable.Add($clustername.name,$clusterDrives)
        }
        ##Add local drives of the computer having cluster
        $Disks = Get-wmiobject  Win32_LogicalDisk -computername $Computer -filter "DriveType= 3" -Credential $cred

        foreach ($objdisk in $Disks) 
        { 
            if($clusterHashTable[$clustername.name] -notcontains $objDisk.DeviceID)
            {
                $c.Cells.Item($intRow, 1) = $Computer.ToUpper() 
                $c.Cells.Item($intRow, 2) = $objDisk.DeviceID 
                $c.Cells.Item($intRow, 3) = "Local"
                $c.Cells.Item($intRow, 4) = “{0:N0}” -f ($objDisk.Size/1GB) 
                $c.Cells.Item($intRow, 5) = “{0:N0}” -f ($objDisk.FreeSpace/1GB) 
                $c.Cells.Item($intRow, 6) = “{0:P0}” -f ([double]$objDisk.FreeSpace/[double]$objDisk.Size) 
                $c.cells.item($introw, 7) = $objdisk.volumename
                $c.cells.item($introw, 8) = $objdisk.DriveType

                $intRow = $intRow + 1
            } 
        }
         
    }else{
        $Disks = Get-wmiobject  Win32_LogicalDisk -computername $Computer -filter "DriveType= 3" -Credential $cred

        foreach ($objdisk in $Disks) 
        { 
           $c.Cells.Item($intRow, 1) = $Computer.ToUpper() 
                $c.Cells.Item($intRow, 2) = $objDisk.DeviceID 
                $c.Cells.Item($intRow, 3) = "Local"
                $c.Cells.Item($intRow, 4) = “{0:N0}” -f ($objDisk.Size/1GB) 
                $c.Cells.Item($intRow, 5) = “{0:N0}” -f ($objDisk.FreeSpace/1GB) 
                $c.Cells.Item($intRow, 6) = “{0:P0}” -f ([double]$objDisk.FreeSpace/[double]$objDisk.Size) 
                $c.cells.item($introw, 7) = $objdisk.volumename
                $c.cells.item($introw, 8) = $objdisk.DriveType

            $intRow = $intRow + 1 
        }
    
    }
 
}
$d.EntireColumn.AutoFit()

cls