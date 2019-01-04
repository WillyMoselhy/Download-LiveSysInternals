Param(
    # Path to download the SysInernals tools
    [Parameter(Mandatory = $true)]
    [String]
    $DownloadPath = 'C:\SysInternals',

    # SysInternals Live URL - default is https://live.sysinternals.com/
    [Parameter(Mandatory = $false)]
    [String]
    $SysInternalsLiveURL = 'https://live.sysinternals.com',

    # List of Files / Folders to skip in the live SysInterals website
    [Parameter(Mandatory = $false)]
    [String[]]
    $SkipList = @('/About_This_Site.txt','/Files/','/Tools/','/WindowsInternals/'),

    # Show logging on screen
    [Parameter(Mandatory = $false)]
    [Switch]
    $HostMode
)

#region: Script Configuration
    $ErrorActionPreference = "Stop"

    $DownloadFolder = New-Item -Path $DownloadPath -ItemType Directory -Force 

    $LogPath = "$DownloadFolder\_Log.log"
    if($LogPath){ # Logs will be saved to disk at the specified location
        $ScriptMode=$true
    }
    else{ # Logs will not be saved to disk
        $ScriptMode = $false
    }
    $LogLevel = 0
    $Trace    = ""   
#endregion: Script Configuration

#region: Logging Functions 
    #This writes the actual output - used by other functions
    function WriteLine ([string]$line,[string]$ForegroundColor, [switch]$NoNewLine){
        if($Script:ScriptMode){
            if($NoNewLine) {
                $Script:Trace += "$line"
            }
            else {
                $Script:Trace += "$line`r`n"
            }
            Set-Content -Path $script:LogPath -Value $Script:Trace
        }
        if($Script:HostMode){
            $Params = @{
                NoNewLine       = $NoNewLine -eq $true
                ForegroundColor = if($ForegroundColor) {$ForegroundColor} else {"White"}
            }
            Write-Host $line @Params
        }
    }
    
    #This handles informational logs
    function WriteInfo([string]$message,[switch]$WaitForResult,[string[]]$AdditionalStringArray,[string]$AdditionalMultilineString){
        if($WaitForResult){
            WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)$message" -NoNewline
        }
        else{
            WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)$message"  
        }
        if($AdditionalStringArray){
            foreach ($String in $AdditionalStringArray){
                WriteLine "                    $("`t" * $script:LogLevel)`t$String"     
            }       
        }
        if($AdditionalMultilineString){
            foreach ($String in ($AdditionalMultilineString -split "`r`n" | Where-Object {$_ -ne ""})){
                WriteLine "                    $("`t" * $script:LogLevel)`t$String"     
            }
       
        }
    }

    #This writes results - should be used after -WaitFor Result in WriteInfo
    function WriteResult([string]$message,[switch]$Pass,[switch]$Success){
        if($Pass){
            WriteLine " - Pass" -ForegroundColor Cyan
            if($message){
                WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)`t$message" -ForegroundColor Cyan
            }
        }
        if($Success){
            WriteLine " - Success" -ForegroundColor Green
            if($message){
                WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)`t$message" -ForegroundColor Green
            }
        } 
    }

    #This write highlighted info
    function WriteInfoHighlighted([string]$message,[string[]]$AdditionalStringArray,[string]$AdditionalMultilineString){ 
        WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)$message"  -ForegroundColor Cyan
        if($AdditionalStringArray){
            foreach ($String in $AdditionalStringArray){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String" -ForegroundColor Cyan
            }
        }
        if($AdditionalMultilineString){
            foreach ($String in ($AdditionalMultilineString -split "`r`n" | Where-Object {$_ -ne ""})){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String" -ForegroundColor Cyan
            }
        }
    }

    #This write warning logs
    function WriteWarning([string]$message,[string[]]$AdditionalStringArray,[string]$AdditionalMultilineString){ 
        WriteLine "[$(Get-Date -Format hh:mm:ss)] WARNING: $("`t" * $script:LogLevel)$message"  -ForegroundColor Yellow
        if($AdditionalStringArray){
            foreach ($String in $AdditionalStringArray){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String" -ForegroundColor Yellow
            }
        }
        if($AdditionalMultilineString){
            foreach ($String in ($AdditionalMultilineString -split "`r`n" | Where-Object {$_ -ne ""})){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String" -ForegroundColor Yellow
            }
        }
    }

    #This logs errors
    function WriteError([string]$message){
        WriteLine ""
        WriteLine "[$(Get-Date -Format hh:mm:ss)] ERROR:   $("`t`t" * $script:LogLevel)$message" -ForegroundColor Red
        
    }

    #This logs errors and terminated script
    function WriteErrorAndExit($message){
        WriteLine "[$(Get-Date -Format hh:mm:ss)] ERROR:   $("`t" * $script:LogLevel)$message"  -ForegroundColor Red
        Write-Host "Press any key to continue ..."
        $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
        $HOST.UI.RawUI.Flushinputbuffer()
        Throw "Terminating Error"
    }

#endregion: Logging Functions

#region: Script Functions

function GetLiveSysInternalsFileList ($URL,$SkipList) {

    WriteInfo "Going to $URL to get file list" -WaitForResult
        $LiveSysInternalsHTML = invoke-webrequest -Uri $URL -UseBasicParsing
        $Lines = $LiveSysInternalsHTML.RawContent -split "`n"
        $lines = $Lines[9] -split "<br>"   
    WriteResult -Success
    
    WriteInfo "Converting list from HTML to PS Object"
    $script:LogLevel++
        foreach ($Line in $Lines){
            if($Line -match "(.{38}).+<A HREF=`"(.+)`""){
                if($matches[2] -notin $SkipList){
                    [PSCustomObject]@{
                        TimeStamp = Get-Date -Date $matches[1]
                        URL       = "$URL$($matches[2])"
                        FileName  = $matches[2].replace("/","")
                    }
                }
                else{
                    WriteInfo "Skipping $($matches[2]) as it is on the Skip List"
                }
            }
        }
    $script:LogLevel--
}
function DownloadSysInternalsFile ($URL,$Destination){
    $Script:LogLevel++
    writeinfo "$URL => $Destination"
        try{
            Invoke-WebRequest -Uri $URL -OutFile $Destination 
        }
        catch{
            WriteError $error[0].Exception.Message
        }
    $Script:LogLevel--
}
#endregion

#region: Get SysInternals Live Website
WriteInfo -message "ENTER: Get SysInternals Live Website"
$LogLevel++

    $LiveVersionList = GetLiveSysInternalsFileList -URL $SysInternalsLiveURL -SkipList $SkipList
    WriteInfo -message "Retrieved this list of files" -AdditionalMultilineString ($LiveVersionList |Sort-Object FileName |Out-string)    

$LogLevel--
WriteInfo -message "Exit: Get SysInternals Live Website"
#endregion: Get SysInternals Live Website

#region: Prepare download folder and import current version list
WriteInfo -message "ENTER: Prepare download folder and import current version list"
$LogLevel++

    WriteInfo "Making sure $DownloadPath exists - will create one if needed" -WaitForResult
        $DownloadFolder = New-Item -Path $DownloadPath -ItemType Directory -Force
    WriteResult -Success
$LogLevel--
WriteInfo -message "Exit: Prepare download folder and import current version list"
#endregion: Prepare download folder and import current version list

#region: Check current version list
WriteInfo -message "ENTER: Check current version list"
$LogLevel++

    $CurrentVersionListPath =  "$DownloadFolder\_CurrentVersionList.csv"
    if(Test-Path -Path $CurrentVersionListPath){
        WriteInfo "File exists:$CurrentVersionListPath"
        $CurrentVersionList = Import-Csv -Path $CurrentVersionListPath
        WriteInfo "Imported CSV File"
    }
    else{
        WriteInfo "File not found: $CurrentVersionListPath"
    }    

$LogLevel--
WriteInfo -message "Exit: Check current version list"
#endregion: Check current version list

#region: Download new files
WriteInfo -message "ENTER: Download new files"
$LogLevel++

    if($CurrentVersionList -eq $null){
        WriteInfo "Current Version List does not exist. Treating all files as new"
        WriteInfo "Downloading all files"

        foreach ($File in $LiveVersionList){
            DownloadSysInternalsFile -URL $File.URL -Destination "$DownloadFolder\$($File.FileName)"
        }
    }
    else{
        foreach($File in $LiveVersionList){
            $PathOnDisk = "$DownloadFolder\$($File.FileName)"
            if((Test-Path -Path $PathOnDisk) -eq $false){ #File does not exist on disk => Download it
                WriteInfo "$($File.FileName) does not exist"
                DownloadSysInternalsFile -URL $File.URL -Destination $PathOnDisk
            }
            else{ #File exists => Check Timestamp
                $CurrentVersion = $CurrentVersionList | where-object {$_.FileName -eq $File.FileName}
                if((Get-Date $CurrentVersion.TimeStamp) -lt $File.TimeStamp){ #Existing file is old
                    WriteInfo "Found new version of $($File.FileName): $(Get-Date $CurrentVersion.TimeStamp) > $($File.TimeStamp)"
                    DownloadSysInternalsFile -URL $File.URL -Destination $PathOnDisk
                }
            }
        }
    }

$LogLevel--
WriteInfo -message "Exit: Download new files"
#endregion: Download new files

#region: Update Current Version List
WriteInfo -message "ENTER: Update Current Version List"
$LogLevel++

    $LiveVersionList | Export-Csv -Path $CurrentVersionListPath -NoTypeInformation

$LogLevel--
WriteInfo -message "Exit: Update Current Version List"
#endregion: Update Current Version List