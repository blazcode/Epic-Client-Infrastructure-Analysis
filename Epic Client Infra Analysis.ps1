# Prerequisites
#     Citrix Powershell Module
#     PowerCLI
#     ImportExcel Module

# References 
#     https://galaxy.epic.com/?#Browse/page=1!68!50!2875795,100047310&from=Galaxy-Redirect
#     https://bronowski.it/blog/2020/06/powershell-into-excelimportexcel-module-part-1/
#     https://github.com/dfinke/ImportExcel


# ----------------------------------------------------
# Configuration
# ----------------------------------------------------

#vCenter to run analysis against; Connect-VIServer can be again against multiple vCenters at once
$vCenter = "tthepicvc01.phsi.promedica.org"

#Citrix Delivery Controller
$citrixDC = "XDDCE01.phsi.promedica.org"

# -----------------------------------------------

#Load Citrix Powershell module
Add-PSSnapin *Citrix*

#Connect to vCenter
$creds = Get-Credential
Connect-VIServer -Server $vCenter -Credential $Creds

$clusters = @()
$hosts = @()
$VMs = @()

#Enumerate vCenter datacenters
foreach($dc in Get-Datacenter){
    Write-Host -foregroundcolor darkgreen $dc.Name
    
    #Enumerate clusters in datacenter 
    #Example for filtering clusters: foreach($cluster in $(Get-Cluster -Location $dc | Where Name -Like "*epicctxclstr*" )){
    foreach($cluster in $(Get-Cluster -Location $dc)){
        Write-Host -foregroundcolor green $cluster.Name
        
        $clusterSessions = 0
        $clusterVMs = 0

        $esxCount = 0
        foreach($esx in Get-VMHost -Location $cluster){
            $esxCount += 1
        }

        #Enumerate hosts in clusters
        foreach($esx in Get-VMHost -Location $cluster){
            Write-Host -foregroundcolor yellow $esx.Name

            $vCPU = Invoke-Expression ((Get-VMHost $esx.Name | Get-VM | Where-Object {$_.PowerState -EQ "PoweredOn"}).NumCPU -join '+')
            $ratio = [math]::round($vCPU / $esx.NumCpu,1)

            $pRAM = $esx.MemoryTotalGB
            $vRAM = Invoke-Expression ((Get-VMHost $esx.Name | Get-VM | Where-Object {$_.PowerState -EQ "PoweredOn"}).MemoryGB -join '+')
       
            $esxVMs = Get-VM -Location $esx  
            
            $esxSessions = 0
            $vmCount = 0

            #Enummerate VMs on hosts
            foreach($esxVM in $esxVMs | Where-Object {$_.PowerState -EQ "PoweredOn"}){
                Write-Host -foregroundcolor gray $esxVM.Name
                $clusterVMs += 1

                $brokerMachine = Get-BrokerMachine -AdminAddress $citrixDC -HostedMachineName $esxVM.Name   
                
                try { 
                    $cpuLoadIndex = $($brokerMachine.LoadIndexes[0].Substring($brokerMachine.LoadIndexes[0].IndexOf(":")+1))
                    $memoryLoadIndex = $($brokerMachine.LoadIndexes[1].Substring($brokerMachine.LoadIndexes[1].IndexOf(":")+1))
                } catch {
                    $cpuLoadIndex = 0
                    $memoryLoadIndex = 0
                }
                        
                $esxSessions += $brokerMachine.SessionCount
                $clusterSessions += $brokerMachine.SessionCount
                $vmCount += 1

                $vmObj = New-Object -TypeName psobject
                $vmObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $esxVM.Name
                $vmObj | Add-Member -MemberType NoteProperty -Name "Host" -Value $esx.Name
                $vmObj | Add-Member -MemberType NoteProperty -Name "vCPU" -Value $($esxVM.NumCpu)
                $vmObj | Add-Member -MemberType NoteProperty -Name "CPU Load Index" -Value $cpuLoadIndex
                $vmObj | Add-Member -MemberType NoteProperty -Name "Target vCPU" -Value 6
                $vmObj | Add-Member -MemberType NoteProperty -Name "RAM" -Value $esxVM.MemoryGB
                $vmObj | Add-Member -MemberType NoteProperty -Name "Target RAM" -Value 30
                $vmObj | Add-Member -MemberType NoteProperty -Name "RAM Load Index" -Value $memoryLoadIndex
                $vmObj | Add-Member -MemberType NoteProperty -Name "Sessions" -Value $brokerMachine.SessionCount
                #$vmObj
                $VMs += $vmObj
            }
            
            $hostsObj = New-Object -TypeName psobject
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Cluster" -Value $cluster.Name
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Host" -Value $esx.Name
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Processor" -Value $esx.ProcessorType
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Running VMs" -Value $vmCount
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Target Running VMs" -Value $([math]::round(($esx.NumCpu * 1.5) / 6))
            $hostsObj | Add-Member -MemberType NoteProperty -Name "pCPUs Available" -Value $($esx.NumCpu)
            $hostsObj | Add-Member -MemberType NoteProperty -Name "vCPUs Allocated" -Value $vCPU
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Target vCPUs Allocated" -Value $($esx.NumCpu * 1.5) 
            $hostsObj | Add-Member -MemberType NoteProperty -Name "CPU Ratio" -Value $ratio
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Target CPU Ratio" -Value 1.5
            $hostsObj | Add-Member -MemberType NoteProperty -Name "RAM (GB)" -Value $([math]::round($pRAM))
            $hostsObj | Add-Member -MemberType NoteProperty -Name "RAM Allocated (GB)" -Value $([math]::round($vRAM))
            $hostsObj | Add-Member -MemberType NoteProperty -Name "RAM Utilized (GB)" -Value $([math]::round($esx.MemoryUsageGB))
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Target RAM Allocated (GB)" -Value $([math]::round(($esx.NumCpu * 1.5) / 6) * 30)
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Sessions" -Value $esxSessions
            $hostsObj | Add-Member -MemberType NoteProperty -Name "Sessions Per Core" -Value $([math]::Round(($esxSessions / $($esx.NumCpu)),1))
            #$hostsObj
            $hosts += $hostsObj
        }

        $clustersObj = New-Object -TypeName psobject
        $clustersObj | Add-Member -MemberType NoteProperty -Name "Name" -Value $cluster.Name
        $clustersObj | Add-Member -MemberType NoteProperty -Name "Host Count" -Value $esxCount
        $clustersObj | Add-Member -MemberType NoteProperty -Name "VM Count" -Value $clusterVMs
        $clustersObj | Add-Member -MemberType NoteProperty -Name "VMs Per Host" -Value $([math]::Round($clusterVMs / $esxCount))
        $clustersObj | Add-Member -MemberType NoteProperty -Name "Sessions" -Value $clusterSessions
        $clustersObj | Add-Member -MemberType NoteProperty -Name "Sessions Per Host" -Value $([math]::Round(($clusterSessions / $esxCount),1))
        #$clustersObj
        $clusters += $clustersObj 
    }
} 

$reportFilename = $("EpicClientComputeAnalysis.xlsx")

#Add cluster analysis to workbook
$clusters | Export-Excel $reportFilename -WorksheetName "Cluster Analysis" -TableName "clusters" -AutoSize

#Go to the temporary location
Set-Location $env:TEMP

#Load the spreadsheet
$excel = Open-ExcelPackage -Path $reportFilename

#Add hosts analysis to workbook
$hosts | Export-Excel -ExcelPackage $excel -WorksheetName "Host Analysis" -TableName "hosts" -AutoSize

#Load the spreadsheet
$excel = Open-ExcelPackage -Path $reportFilename

#Add VM analysis to workbook
$VMs | Export-Excel -ExcelPackage $excel -WorksheetName "VM Analysis" -TableName "vms" -AutoSize -Show
