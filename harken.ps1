<# 
   .SYNOPSIS
   Creates a .NET HTTP Listener as a Windows service which allows PowerShell Commands to be invoked via HTTP calls.

   .DESCRIPTION
   Creates a .NET HTTP Listener as a Windows service which allows PowerShell Commands to be invoked via HTTP calls.
   More information about the .NET HTTP Listener is available here: https://gallery.technet.microsoft.com/scriptcenter/Simple-REST-api-for-b04489f1

   .PARAMETER ServiceName
   Optional, String: The name you would like to set for the service. 
   Default: Harken

   .PARAMETER HttpListenerURL
   Optional, String: The URL you'd like to use for the HTTP Listener. (http://localhost:8888/harken)
   Default: harken
   
   .PARAMETER HttpListenerPort
   Optional, String: The port number you would like to use for the HTTP Listener. 
   Default: 8888

   .PARAMETER HttpListenerAuth
   Optional, String: The authentication method you would like to use for the HTTP Listener.
   Valid Options: Anonymous, Basic, Digest, IntegratedWindowsAuthentication, Negotiate, None, Ntlm
   Default: IntegratedWindowsAuthentication
   More Info: https://msdn.microsoft.com/en-us/library/system.net.authenticationschemes(v=vs.110).aspx
   
   .PARAMETER HttpListenerVersion
   Optional, String: The version of HTTPListener you'd like to download. 
   Default: 1.0.1

   .PARAMETER HttpListenerDir
   Optional, String: The HttpListener directory. 
   Default: C:\Windows\System32\WindowsPowerShell\v1.0\Modules

   .PARAMETER NssmVersion
   Required, String: The version of nssm you'd like to download.
   Default: 2.24
   
   .PARAMETER NssmDir
   Optional, String: The nssm directory.
   Default: C:\nssm\
   
   .PARAMETER NssmArch
   Optional, String: The Architecture of nssm you'd like to use. 
   Valid Options: win32, win64
   Default: win64

   .EXAMPLE Install Harken with all the defaults.
   Install-Harken

   .EXAMPLE Uninstall Harken
   Uninstall-Harken

   .EXAMPLE Using Harken
   Visit this URL in your browser http://localhost:8888/harken?command=get-service winmgmt&format=text"
#> 

# Function to unzip files.
function Unzip-File {
    param(
        [Parameter(mandatory=$true)][string]$File,
        [Parameter(mandatory=$true)][string]$Destination
    )
    Write-Host "Extracting file: $File to $Destination"
    Add-Type -assembly 'system.io.compression.filesystem'
    if (!(Get-ChildItem -Path $Destination)) {
        [io.compression.zipfile]::ExtractToDirectory($File, $Destination)
    }
    Write-Host 'Extraction Complete!'
}

# Function to download files
function Download-File {
    param(
        [Parameter(mandatory=$true)][string]$URL,
        [Parameter(mandatory=$true)][string]$Destination
    )
    $StartTime = Get-Date
    Invoke-WebRequest -Uri $URL -OutFile $Destination
    Write-Host "Downloading file from: $URL"
    Write-Host "Download took: $((Get-Date).Add(-$StartTime).Millisecond) ms."
    Write-Host ' '
}


# Install Harken Function
function Install-Harken {
    param(
        [Parameter(mandatory=$false)][string]$ServiceName = 'Harken',
        [Parameter(mandatory=$false)][string]$HttpListenerURL = 'harken',
        [Parameter(mandatory=$false)][string]$HttpListenerPort = 8888,
        [Parameter(mandatory=$false)][System.Net.AuthenticationSchemes] $HttpListenerAuth = [System.Net.AuthenticationSchemes]::IntegratedWindowsAuthentication,
        [Parameter(mandatory=$false)][string]$HttpListenerVersion = '1.0.1',
        [Parameter(mandatory=$false)][string]$HttpListenerDir = 'C:\Windows\System32\WindowsPowerShell\v1.0\Modules\',
        [Parameter(mandatory=$false)][string]$NssmVersion = '2.24',
        [Parameter(mandatory=$false)][string]$NssmDir = 'C:\nssm\',
        [Parameter(mandatory=$false)][ValidateSet('win32','win64')][string]$NssmArch = 'win64'
    )
    try {
        # Create the NssmDir if it does not exist.
        if (!(Test-Path -Path $NssmDir -PathType Container)) {
            New-Item -Path $NssmDir -ItemType Directory
        }
        else {
            Write-Host ' ' 
            Write-Host ($NssmDir + 'already exists, continuing installation...')
        }

        # Create temporary download location for files.
        $TempDir = 'C:\temp'
        if (!(Test-Path -Path $TempDir -PathType Container)) {
            New-Item -Path $TempDir -ItemType Directory | Out-Null
        }

        # Go download the latest nssm (Non-Sucking Service Manager) if it's not already installed.
        $NssmURL = ('http://nssm.cc/release/nssm-' + $NssmVersion + '.zip')
        $NssmZip = ($TempDir + '\nssm.zip')
        Download-File -URL $NssmURL -Destination $NssmZip

        # Unzip the download into the $NssmDir location.
        Unzip-File -File $NssmZip -Destination $NssmDir 
        
        # Create the HttpListenerDir if it does not exist.
        if (!(Test-Path -Path $HttpListenerDir -PathType Container)) {
            New-Item -Path $HttpListenerDir -ItemType Directory
        }
        else {
            Write-Host ' '
            Write-Host ($HttpListenerDir + 'already exists, continuing installation...')
        }

        # Go download the lastest PowerShell HTTP Listener.
        $HttpListenerFileURL = ('https://gallery.technet.microsoft.com/Simple-REST-api-for-b04489f1/file/126130/1/HttpListener_' + $HttpListenerVersion + '.zip')
        $HttpListenerZip = ($TempDir + '\HttpListener.zip')
        Download-File -URL $HttpListenerFileURL -Destination $HttpListenerZip

        # Unzip into the $HttpListenerDir location.
        Unzip-File -File $HttpListenerZip -Destination $HttpListenerDir

        # Create the .NET HTTPListener as a Windows Service using nssm.exe (Harken)
        $NssmCmd = ('install ' + $ServiceName + ' "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-command "& { Import-Module -Name HTTPListener ; Start-HTTPListener -URL ' + $HttpListenerURL + ' -Port ' + $HttpListenerPort + ' -Auth ' + $HttpListenerAuth + ' }"" ')
        switch ($NssmArch) {
            # Create and install the service using the 32bit version of nssm.exe
            'win32' {
                Set-Location ($NssmDir + '\nssm-2.24\' + $NssmArch) 
                Start-Process -FilePath .\nssm.exe -ArgumentList $NssmCmd -NoNewWindow -Wait 
                Write-Host "Service $ServiceName installed with $NssmArch nssm.exe" 
            }
            # Create  and install the service using the 64bit version of nssm.exe
            'win64' {
                Set-Location ($NssmDir + '\nssm-2.24\' + $NssmArch)
                Write-Host $NssmCmd
                Start-Process -FilePath .\nssm.exe -ArgumentList $NssmCmd -NoNewWindow -Wait
                Write-Host "Service $ServiceName installed with $NssmArch nssm.exe"
            }
        }

        # Test that the service was successfully installed
        if (Get-Service -Name $ServiceName) {
            Write-Host "The $ServiceName service was created successfully."
        }

        # Check if the service is stopped. If it is, start it!
        if ((Get-Service -Name $ServiceName).Status -eq 'Stopped') {
            Get-Service -Name $ServiceName | Start-Service
        }
        
        # Test basic HTTP call using Invoke-RestMethod
        
        # Do some basic clean up.
        if (Test-Path -Path $TempDir -PathType Container) {
            Remove-Item -Path $TempDir -Recurse -Force
        }
        
        # Once everything looks good, notify the user where the service is currently running. 
        Write-Host ' '
        Write-Host '-------------------------------------------------'
        Write-Host ('Harken is now running at http://localhost:' + $HttpListenerPort + '/' + $HttpListenerURL )
        Write-Host '-------------------------------------------------'
        Write-Host ' '
    }
    catch {
        Write-Warning $_ 
    }       
}

# Uninstall Harken Function
function Uninstall-Harken {
    param(
        [Parameter(mandatory=$false)][string]$ServiceName = 'Harken', # ServiceName
        [Parameter(mandatory=$false)][string]$NssmDir = 'C:\nssm\', # NssmInstallDir (where you want to install nssm)
        [Parameter(mandatory=$false)][ValidateSet('win32','win64')][string]$NssmArch = 'win64' # NssmArch (32 or 64 bit .exe file)
    )
    try {
        # Stop the Service
        Get-Service -Name $ServiceName | Stop-Service
        # Uninstall the Service
        switch ($NssmArch) {
            # Uninstall the service using the 32bit version of nssm.exe
            'win32' {
                Set-Location ($NssmDir + '\nssm-2.24\' + $NssmArch) 
                $cmd =  ('remove ' + $ServiceName + ' confirm')
                Start-Process -FilePath .\nssm.exe -ArgumentList $cmd  -NoNewWindow
                Start-Sleep 5
                Write-Host "Service $ServiceName removed with $NssmArch nssm.exe" 
            }
            # Uninstall the service using the 64bit version of nssm.exe
            'win64' {
                Set-Location ($NssmDir + '\nssm-2.24\' + $NssmArch)
                $cmd =  ('remove ' + $ServiceName + ' confirm')
                Start-Process -FilePath .\nssm.exe -ArgumentList $cmd  -NoNewWindow
                Start-Sleep 5
                Write-Host "Service $ServiceName removed with $NssmArch nssm.exe"
            }
        }        
    }
    catch {
        Write-Warning $_ 
    }       
}
