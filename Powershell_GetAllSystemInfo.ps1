###############################################################################################################################################
# Purpose: To extract system information, OS, Shares, Disk, Partition and Windows Cluster Details from a list of remote Servers
#Pre-Reqs: 
#  WMI should be enabled on all target servers
#  WMI ports to all target servers, should be opened from the location where the script is run
#  Create a list of computers in computer.txt file;Update $location variable
################################################################################################################################################


$location = ".\server_details"
$cred= get-credential
$Computers = get-content "$location\computers.txt"
$clusterArray = New-Object System.Collections.ArrayList
$clusterHashTable = New-Object System.Collections.Hashtable
function export-tasks ($computer,$cred,$location){
    $cred.Username
    $domain=""
    $user=""
    If($cred.Username -match "\\") {
        $userarr=$cred.Username.Split("\\")
        $user=$userarr[1]
        $domain=$userarr[0]
    } Else {
        $user=$cred.Username
        $domain=$env:USERDOMAIN
    }
    #$domain
    #$user

    $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($cred.Password)
    $result = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
    #$result
    
    $sch = New-Object -ComObject("Schedule.Service")
    $sch.Connect($computer,$user,$domain,$result)
    $tasks = $sch.GetFolder("\").GetTasks(0)

    $outfile_temp = "${location}\${Computer}\Task_{0}_${computer}.xml"

    $tasks | %{
	    $xml = $_.Xml
	    $task_name = $_.Name
	    $outfile = $outfile_temp -f $task_name
	    $xml | Out-File $outfile
    }

}

foreach ($Computer in $Computers) 
{
    New-Item -Path "${location}\${Computer}" -ItemType "directory"
    Write-Output "Processing computer: $Computer"
    $ErrorLogMain = New-Item -Path "${location}\${Computer}\ErrorMainBlock_${Computer}.txt" -ItemType "file"
    try{
    
    ##Get disks attached to the server
    gwmi -query "Select * from Win32_DiskDrive" -ComputerName $Computer -Credential $cred | Select-Object *|export-csv -path ${location}\${Computer}\diskDrive_${Computer}.csv -Force
    ##Get disk partitions for the server
    gwmi -query "Select * from Win32_DiskPartition" -ComputerName $Computer -Credential $cred | Select-Object *|export-csv -path ${location}\${Computer}\diskPartition_${Computer}.csv -Force
    ##Get OS details for the server
    gwmi -query "Select * from  Win32_OperatingSystem" -ComputerName $Computer -Credential $cred | Select-Object *|export-csv -path ${location}\${Computer}\OSDetails_${Computer}.csv -Force
    ##Get all computer details
    gwmi -Class Win32_ComputerSystem -computername $Computer -Credential $cred| Select-Object *|export-csv -path ${location}\${Computer}\CompDetails_${Computer}.csv -Force
    ##Get Bios details
    gwmi Win32_BIOS -computername $Computer -Credential $cred| Select-Object *|export-csv -path ${location}\${Computer}\BIOS_${Computer}.csv -Force
    ##Get NIC details
    gwmi Win32_NetworkAdapterConfiguration -computername $Computer -Credential $cred -filter ipenabled="true" | Select-Object *|export-csv -path ${location}\${Computer}\NICs_${Computer}.csv -Force
    ##Get all windows shares
    gwmi Win32_share -computername $Computer -Credential $cred| Select-Object *|export-csv -path ${location}\${Computer}\Fileshares_${Computer}.csv -Force
    ##Get disk information
    $Disks = Get-wmiobject  Win32_LogicalDisk -computername $Computer -filter "DriveType= 3" -Credential $cred    
    New-Item -Path "${location}\${Computer}\LogicalDiskInfo_${Computer}.csv" -ItemType "file"
    "Machine Name,Drive,Cluster/Local,Total size (GB),Free Space (GB),Free Space (%),Name,DriveType"|Add-Content -Path "${location}\${Computer}\LogicalDiskInfo_${Computer}.csv"


    if($?)
    {

    Write-Output "Processed WMI Queries on computer: $Computer"

    $ErrorLog = New-Item -Path "${location}\${Computer}\Error_${Computer}.txt" -ItemType "file"
    
    ##Get Cluster details 
    try{
        $clustername= gwmi MSCluster_Cluster -Namespace root/mscluster -ComputerName $Computer -Credential $cred|select name
        if($?)
        {
            if($clustername)
            {
                    Write-Output "Processing Cluster on computer: $Computer"
                    
                    ##Get cluster resources list
                    New-Item -Path "${location}\${Computer}\Cluster_${Computer}.csv" -ItemType "file"
                    "Computer,ClusterName,ResourceName,OwnerNode,ResourceType,ClusterSharedVolume"|Add-Content -Path "${location}\${Computer}\Cluster_${Computer}.csv"
                    
                    
                    $clusterDrives = New-Object System.Collections.ArrayList
                    ##see if this cluster is processed
                    if($clusterArray -notcontains $clustername.name){            
                        ##Add to cluster array
                        $clusterArray.Add($clustername.name)
                        
                        ##Get cluster resources list
                        $ClusterResources=gwmi MSCluster_Resource -Namespace root/mscluster -ComputerName $Computer -Credential $cred
                        foreach ($objresource in $ClusterResources) 
                        { 
                            
                            $Computer+","+$clustername.Name.ToUpper()+","+$objresource.Name+","+$objresource.OwnerNode+","+$objresource.Type+","+$objresource.IsClusterSharedVolume|Add-Content -Path "${location}\${Computer}\Cluster_${Computer}.csv"
    
                        }

                        ##Get cluster logical volumes list
                        $ClusterDisks=gwmi MSCluster_DiskPartition -Namespace root/mscluster -ComputerName $Computer -Credential $cred| Select Path, TotalSize, FreeSpace, VolumeLabel
                        foreach ($objdisk in $ClusterDisks) 
                        { 
                            $clusterDrives.Add($objDisk.Path)                
                            $clustername.Name.ToUpper() +"," + $objdisk.Path + "," + "Cluster" + "," + "{0:N0}" -f ($objdisk.TotalSize/1024) + "," + "{0:N0}" -f ($objdisk.FreeSpace/1024) + "," + "{0:P0}" -f ([double]$objdisk.FreeSpace/[double]$objdisk.TotalSize) + "," + $objdisk.Volumelabel + "," + "3" |Add-Content -Path "${location}\${Computer}\LogicalDiskInfo_${Computer}.csv"
   
                        }
                        $clusterHashTable.Add($clustername.name,$clusterDrives)

                    }
                    
                    ##Add local drives of the computer having cluster
                    ##This section added here to avoid duplicate entries of cluster logical volumes with local ones               

                    foreach ($objdisk in $Disks) 
                    { 
                        if($clusterHashTable[$clustername.name] -notcontains $objDisk.DeviceID)
                        {
                
                           $Computer.ToUpper() +"," + $objdisk.DeviceID + "," + "Local" + "," + "{0:N0}" -f ($objdisk.Size/1GB) + "," + "{0:N0}" -f ($objdisk.FreeSpace/1GB) + "," + "{0:P0}" -f ([double]$objdisk.FreeSpace/[double]$objdisk.Size) + "," + $objdisk.Volumename + "," + $objdisk.DriveType |Add-Content -Path "${location}\${Computer}\LogicalDiskInfo_${Computer}.csv"

                        } 
                    }
              }
          }
        else{
            throw $error[0].Exception
        }

        if($null -eq $clustername){
            ##Logical disk info for computers not having clusters
            foreach ($objdisk in $Disks) 
            { 
                $Computer.ToUpper() +"," + $objdisk.DeviceID + "," + "Local" + "," + "{0:N0}" -f ($objdisk.Size/1GB) + "," + "{0:N0}" -f ($objdisk.FreeSpace/1GB) + "," + "{0:P0}" -f ([double]$objdisk.FreeSpace/[double]$objdisk.Size) + "," + $objdisk.Volumename + "," + $objdisk.DriveType |Add-Content -Path "${location}\${Computer}\LogicalDiskInfo_${Computer}.csv"

            }    
        }              
           
      }Catch{
	        Write-Output "$ErrorMessage : $FailedItem"
            $ErrorMessage = $_.Exception.Message
	        $FailedItem = $_.Exception.ItemName
            $_|Add-Content -Path "${location}\${Computer}\Error_${Computer}.txt"
	    }

     }else
        {
            throw $error[0].Exception
        }


    export-tasks $Computer $cred $location

    }Catch{
	        Write-Output "$ErrorMessagemain : $FailedItemmain"
            $ErrorMessagemain = $_.Exception.Message
	        $FailedItemmain = $_.Exception.ItemName
            $_|Add-Content -Path "${location}\${Computer}\ErrorMainBlock_${Computer}.txt"
	    }
}