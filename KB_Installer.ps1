# Set the KB number to check for and install
$KBNumber = "KB5001330"
$Global:LogFile = "$PSScriptRoot\Custom_KB_$($KBNumber)_Installation.log" #Log file location

#Function for logging
function Write-Log {
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Validateset("INFO", "ERR", "WARN")]
        [string]$Type = "INFO"
    )

    if( -Not [System.Diagnostics.EventLog]::SourceExists("KBInstall")){
        New-EventLog -LogName Application -Source "KBInstall"
    }

    $DateTime = Get-Date -Format "MM-dd-yyyy HH:mm:ss"
    $FinalMessage = "[{0}]::[{1}]::[{2}]" -f $DateTime, $Type, $Message

    #Storing the output in the log file
    $FinalMessage | Out-File -FilePath $LogFile -Append

    if ($Type -eq "ERR") {
        Write-EventLog -LogName Application -Source "KBInstall" -EntryType Error -EventId 1000 -Message $FinalMessage
        Write-Host "$FinalMessage" -ForegroundColor Red
    }
    else {
        Write-EventLog -LogName Application -Source "KBInstall" -EntryType Information -EventId 1000 -Message $FinalMessage
        Write-Host "$FinalMessage" -ForegroundColor Green
    }
}

try{

    Write-Log "Checking the HotFix Status of $KBNumber"
    $KBInstalled = Get-HotFix | Where-Object {$_.HotFixID -eq $KBNumber}

    # If the KB is not installed, install it
    if (-not $KBInstalled) {

        Write-Log "KB is not installed on this machine, therefore installation initiated"

        $KBFile = "C:\Temp\$KBNumber.msu"

        # Install the KB
        Write-Progress -Activity "Installing KB $KBNumber" -Status "Installing..."

        if(-Not (Test-Path -Path $KBFile -PathType Leaf)){
            write-Log "File $KBFile is not found" -Type ERR
            throw "File $KBFile is not found"
        }

        wusa $KBFile /quiet /norestart
        Write-Progress -Activity "Installing KB $KBNumber" -Completed

        Start-Sleep -Seconds 2

        # Validate the installation
        $KBInstalled = Get-HotFix | Where-Object {$_.HotFixID -eq $KBNumber}

        if ($KBInstalled) {
            Write-Host "KB $KBNumber successfully installed"

            Write-Progress -Activity "Rebooting machine" -Status "Rebooting..."
            Start-Sleep -Seconds 5
            Restart-Computer

        } else {
            Write-Log "KB $KBNumber is not installed, even after install initiation" -Type ERR
        }
    }else{
        Write-Log "KB $KBNumber is already installed on the machine"
    }

}
catch{
    Write-Log "Error installing KB $KBNumber - $_" -Type ERR
}
