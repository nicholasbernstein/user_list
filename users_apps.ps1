# nick@nicholasbernstein.com 2022/7/7
# Script expects powershell 5.1 & windows 2016 server but should work in later versions
# Script is called with two mandatory paramaters, one optional
#
# This script checks to see if a process is running, and then logs information about it to a file
#
# INPUT:
# users_apps.ps -Servers <path to csv file> -Apps <path to csv file> -Append
# Servers points to a .csv file which includes a list of servers to check
# Apps points to a .csv file which includes a list of files to check
# Append will append to the output 
#
# OUTPUT:
# csv file, path defined in the $Output variable

[CmdletBinding()]
Param(
	[Parameter(Mandatory = $true)][string]$Servers,
	[Parameter(Mandatory = $true)][string]$Apps,
	[Parameter(Mandatory = $true)][string]$Output,
	[switch]$Append,
    [switch]$Excel
)

# Line ending is either a newline (unix & mac) or a carriage return (windows)
# Newline
$LE =  "`n"
# CR
#$LE = "`r"

# Delimiter can be , or tab (default for excel)
#$del="`t"
$del=","

# Ensure that we are pinned to the intended version of powershell:
# At beginning of .ps1
#$MyPSVersion = "5.1"
if ($PSVersionTable.PSVersion -ne [Version]"$MyPSVersion") {
	# Re-launch as version 5 if we're not already
	Write-Warning -Message "Expected PS version: "
	powershell -Version $MyPSVersion -File $MyInvocation.MyCommand.Definition
	exit
}

Clear-Host
$ErrorActionPreference = "SilentlyContinue"
Write-Host
Write-Host "+--------------------------------------+"
Write-Host "| --- Users and their applications --- |"
Write-Host "+--------------------------------------+"
Write-Host ""

function ensureExcelModuleIsInstalledAndAvailable(){
    if (Get-Module -ListAvailable -Name "importexcel") {
        Import-Module importexcel
    } 
    else {
        Write-Host "installing importexcel module"
        Install-Module importexcel
        Import-Module importexcel
    }
}

ensureExcelModuleIsInstalledAndAvailable


function generateListOfServersFromCSVFile($filePath)
{
    try{
        Write-Host "[+] Reading file: $filePath"
        $csvServers = Import-Csv $filePath -UseCulture
    } 

    catch {
        Write-Host "A Fatal error occurred opening $filePath"
        Write-Host $_
        exit
    }
    
    $ServerList = New-Object Collections.Generic.List[string]
    foreach ($i in $csvServers) { 
        $ServerList.Add($i.'Server Name')
    }
    return $ServerList

}

function generateListofAppsFromCSVFile($filePath){
    try{
        Write-Host "[+] Reading file: $filePath"
        $csvApps = Import-Csv $filePath -UseCulture
    }

    catch {
        Write-Host "A Fata; error occurred opening $filePath"
        Write-Host $_
        exit
    }

    return $csvApps
}
    
$ServerList = generateListOfServersFromCSVFile($Servers)
$csvApps = generateListofAppsFromCSVFile($Apps)

$csvContents = Invoke-Command -ComputerName $ServerList -ScriptBlock {
###############
# This whole section is passed as a script-block to invoke command

    function printCSV($a){
		# $using:LE passes local variable to be used in scriptbllock
        $line = $a.Server + $using:del + $a.Date + $using:del + $a.Time + "`t" + $a.Application + $using:del + $a.User + $using:del + $a.Memory + $using:del + $a.Path + $using:LE
        return $line
        #$a.values | export-csv
    }

    function convertAndFormatInMB($size){
        $size = $size
        $size = ($size/1MB)
        $size = $size.tostring("##.##")
        return $size
    }

    function getProcessInfoFor32B($process) {

        $memInMB = convertAndFormatInMB($process.PrivateMemorySize[0] -as[int])
        $procInfo   = @{
        Server      = $env:computername
        Date        = Get-Date -UFormat "%m/%d/%Y" 
        Time        = $process.StartTime
        Application = $process.Name
        User        =  stripDomainFromUsername($process.UserName)
        Memory      = $memInMB + "MB"
        Path        = $process.Path
        }

        return $procInfo
    }

    function stripDomainFromUsername($name){
        #write-host "Stripping domain"
        $domain, $user = $name.split("\")
        return $user
    }

    function getProcessInfoFor64B($process) {
        $memInMB = convertAndFormatInMB($process.PrivateMemorySize[0] -as[int])
        $procInfo   = @{
        Server      = $env:computername
        Date        = Get-Date -UFormat "%m/%d/%Y" 
        Time        = $process.StartTime
        Application = $process.Name
        User        = stripDomainFromUsername($process.UserName)
        Memory      = $memInMB + "MB"
        Path        = $process.Path
        }
        return $ProcInfo
    }



    function findMatchingProcessesAndReturnCSVFormattedRows($applicationList)
    {
        # Get one process list to search through and then re-use it for performance
        $procList = Get-Process -IncludeUserName
        
        foreach ($myProc in $procList) {
            foreach ($i in $applicationList) { 
                $applicationName = $i.Application
                $exeName = $MyProc.Name + ".exe"
                #write-host $exeName "--- $applicationName"
                if ( $exeName -eq $applicationName) {
                    #write-host "[+] Match for for $applicationname"
                    # 32b programs return a string for their attributes instead of an object
                    if ($myProc.Path.GetType().Name -eq 'String') {

                        $procInfo = getProcessInfoFor32B($myProc)

                    } else {

                       $procInfo = getProcessInfoFor64B($myProc)
                    }

                #$rows += printCSV($procInfo)
                printCSV($procInfo)
                $procInfo = @{}
                }
            }
        }
        #return $rows
    }

    # The $using:csvAPPS passes a local variable to the scriptblock as $csvAPPS

    findMatchingProcessesAndReturnCSVFormattedRows($using:csvAPPS)

    ## End Scriptblock
    }


# Output from servers is in $cvsContents, lets write or append that to a file
    #$csvContents
	$csvHeader="Server" + $del + "Date" + $del + "Time" + $del + "Application" + $del + "User" + $del + "Memory" + $del + "Path" + $LE
	if ($csvContents) {
		$csv = $csvHeader
		$csv += $csvContents
        #$csv = $csv -replace "^`s+", ""
        #write-host $csv
		if ($Append) {
			$csvContents | out-file $Output -NoNewline -Append  -Encoding UTF8 
            #$csvContents | Export-Excel "$Output.xlsx" -Append  -autosize -autofilter
			Write-Host "[+] Successfully appended to" $Output
		}
		else {
			$csvHeader | out-file $Output -NoNewline -Encoding UTF8
            $csvContents | out-file $Output -NoNewline -Append  -Encoding UTF8
            #$Output
            #$csv |  Export-Excel "$Output.xlsx"   -autosize -autofilter
			#Write-Host "[+] Successfully exported to" $Output
            Write-Host "[+] Successfully exported to" $Output
		}
         if ($Excel) {
         # The following converts the final csv to native excel format
            $fname, $extension = $Output.Split(".")
            $ExcelOutputFileName = $fname + ".xlsx"
            $myCSV = Import-Csv -Path $Output -Delimiter $del -Header Server,Date,Time,Application,User,Memory,Path -Encoding UTF8
            $myCSV | Export-Excel -autosize -autofilter -Path $ExcelOutputFileName
            Write-Host "[+] converted csv to $ExcelOutputFIleName"
         }
	}