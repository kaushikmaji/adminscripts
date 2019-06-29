#######################################################################################################
#Purpose: The script generates an excel report showing health status of resources in a windows cluster
#Pre-Reqs: 
#  WMI should be enabled on all target servers
#  WMI ports to all target servers, should be opened from the location where the script is run  
#  Excel should be installed in the location where the script is run
########################################################################################################

$erroractionpreference = “SilentlyContinue” 
$a = New-Object -comobject Excel.Application 
$a.visible = $True
$b = $a.Workbooks.Add() 
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = “Machine Name” 
$c.Cells.Item(1,2) = “Cluster Name” 
$c.Cells.Item(1,3) = "Owner Node" 
$c.Cells.Item(1,4) = “Resource Name”
$c.Cells.Item(1,5) = “Status”
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
$Computers = get-content "servers.txt"

$Computers

foreach ($Computer in $Computers) 
{ 
    ##For cluster
    try{
    $clustername= gwmi MSCluster_Cluster -Namespace root/mscluster -ComputerName $Computer -Credential $cred|select name
    if($clustername)
    {

         $clusterDrives = New-Object System.Collections.ArrayList
            ##see if this cluster is processed
            if($clusterArray -notcontains $clustername.name){            
                ##Add to cluster array
                $clusterArray.Add($clustername.name)
                $ClusterResources=gwmi MSCluster_Resource -Namespace root/mscluster -ComputerName $Computer -Credential $cred
                foreach ($objresource in $ClusterResources) 
                { 
                    $clusterDrives.Add($objresource.Path)
                    $c.Cells.Item($intRow, 1) = $Computer 
                    $c.Cells.Item($intRow, 2) = $clustername.Name.ToUpper()
                    $c.Cells.Item($intRow, 3) = $objresource.OwnerNode
                    $c.Cells.Item($intRow, 4) = $objresource.Name
                    $c.Cells.Item($intRow, 5) = $objresource.State

                    if($objresource.State -eq 2)
                    {
                        $c.Cells.Item($intRow, 5).Interior.ColorIndex = 43
                    }
                    if($objresource.State -eq 4)
                    {
                        $c.Cells.Item($intRow, 5).Interior.ColorIndex = 3
                    }
                    if($objresource.State -eq 3)
                    {
                        $c.Cells.Item($intRow, 5).Interior.ColorIndex = 45
                    }

                    $intRow = $intRow + 1 
                }
            }
         }
        }Catch{
	        $ErrorMessage = $_.Exception.Message
	        $FailedItem = $_.Exception.ItemName
            Write-Output "$Computer could not be processed"
            
	    }

}
$d.EntireColumn.AutoFit()
