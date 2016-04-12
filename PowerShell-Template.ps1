#region Header
<#  
.SYNOPSIS  
    .

.DESCRIPTION
    .
                
.PARAMETER 
    [boolean] Logging
    If true starts transcription logging and verbose output.

.EXAMPLE
    .

.INPUTS
    None. You cannot pipe objects to this function.

.OUTPUTS
    None. 

.NOTES  
    Author     : Glen Buktenica
	Change Log : 2016XXXX Initial Build  
#> 
Param(
    [switch]$Logging #= $true # decomment line to force logging.
     ) 
#region Functions
Function Configure-Logs
{
<#  
.SYNOPSIS  
    Configure standard and error log file paths for use by Export-Logs.

.EXAMPLE
    Configure-Logs should be the first function called.

.INPUTS
    [boolean] Logging
    If Logging True starts transcription logging and verbose output.

.OUTPUTS
    [string] LogPath            Transcripting output
    [string] ErrorPath          Standard error output
    [string] VerbosePreference  Enables verbose screen output

.NOTES  
    Author     : Glen Buktenica
	Change Log : 20160329 Initial Build 
               : 20160412 Bug fix: Global error path
#> 
    If ($Logging){$VerbosePreference = "Continue"} else {$VerbosePreference = "SilentlyContinue"} 
    If ($MyInvocation.ScriptName) #Confirm script has been saved
    {
        # Create log paths in the same location as the script
        $CurrentPath = Split-Path $MyInvocation.ScriptName
        $ScriptName = [io.path]::GetFileNameWithoutExtension($MyInvocation.ScriptName)
        $LogPathDate = Get-Date -Format yyyyMMdd
        $global:LogPath = "$CurrentPath\$ScriptName$LogPathDate.log"       # Example c:\scripts\MyScript20160329.log
        $global:ErrorPath = "$CurrentPath\$ScriptName$LogPathDate.err.log" # Example c:\scripts\MyScript20160329.err.log
        Try
        {
            # Test that log file location is able to be written to
            " " | Out-File "$LogPath.test"
        }
        Catch
        {
            # If log file location not writable save to current user Temp
            $global:LogPath = "$env:LOCALAPPDATA\temp\$ScriptName$LogPathDate.log"       # Example C:\Users\bukteng0\AppData\Local\temp\MyScript20160329.log
            $global:ErrorPath = "$env:LOCALAPPDATA\temp\$ScriptName$LogPathDate.err.log" # Example C:\Users\bukteng0\AppData\Local\temp\MyScript20160329.err.log
            {$error.Remove($error[0])}
        }
        Finally
        {
            # Clean up test file
            Remove-Item "$LogPath.test" -Force -ErrorAction SilentlyContinue
        }
        Write-Verbose $LogPath
        Write-Verbose $ErrorPath
        # Start standard logging if NOT running in PowerShell ISE and logging enabled and the log path is valid.
        If(($Host.UI.RawUI.BufferSize.Height -gt 0) -and ($Logging) -and ($LogPath)) 
        {
            $Transcripting = $true
            Start-Transcript -path $LogPath -append
        }
        # Clear standard error in case populated by previous scripts
        $Error.Clear()
    }
}
Function Export-Logs
{
<#  
.SYNOPSIS  
    If the standard error variable is populated then all errors will be saved to a text file.

.EXAMPLE
    Export-Logs should be the last line in a script and/or called before Exit command.

.INPUTS
    None. You cannot pipe objects to this function.

.OUTPUTS
    [text file] In the current script path or if not writeable in the user local temp path.

.NOTES  
    Author     : Glen Buktenica
	Change Log : 20160329 Initial Build 
               : 20160412 Bug Fix: Export Variables 
#> 
    If ($Error -and $ErrorPath)
    {
        Write-Verbose "Errors were found"
        Get-Date -Format "dddd, d MMMM yyyy HH:mm" | Out-File -FilePath $ErrorPath -append
        "User name: $env:USERNAME"  | Out-File -FilePath $ErrorPath -append
        "Computer name: $env:COMPUTERNAME" | Out-File -FilePath $ErrorPath -append
        $Error.Reverse()
        # Loop through each error, getting detailed error information and write to error log file
        $Error | ForEach-Object {"--------------------------------------------------"
            "Exception:"
            $_.Exception
            "FullyQualifiedErrorId:"
            $_.FullyQualifiedErrorId
            "ScriptStackTrace:"
            $_.ScriptStackTrace
            $a = $_.InvocationInfo.ScriptLineNumber 
            "Error line:  $a"
            $_.CategoryInfo } | Out-File -FilePath $ErrorPath -append
    }
    # If transcripting started, stop the transcript
    Try{If($Transcripting)
    {Stop-Transcript | Out-Null}}
    Catch [System.InvalidOperationException]{$error.Remove($error[0])}

    Write-Verbose "Script End"
    If ($Logging){Start-Sleep -s 10}
    $Logging = $false
}
Function Select-GUI
{
<#  
.SYNOPSIS  
    Open or save files or open folders using Windows forms.

.PARAMETER Start
    [string] The start directory for the form.
    Default value: Last path used

.PARAMETER Description
    [string] Text that is included on the chrome of the form and writen to the console window.
    Default value: Select <Folder/File> to <Open/Save>

.PARAMETER Ext
    [string] Adds an extension filter to file open or save forms. 
    Default value *.* 

.PARAMETER File
    [switch] When present this switch uses file open or save forms. Used with the Save switch.

.PARAMETER Save
    [switch] When used with the File switch launches a file save form. When not present launches a file open form.

.PARAMETER UNC
    [switch] When used with no File swich and the required dll is missing will use Read-Host so a UNC path can be entered instead of failing back to the native non-unc path form.

.EXAMPLE
    Select-GUI -Start ([Environment]::GetFolderPath('MyDocuments')) -Description "Save File" -Ext "csv" -File -Save
    Windows save file form filtered for *.csv defaulting to the current users my documents folder and the description "Save File"

.EXAMPLE
    Select-GUI -Start "C:\" -Description "Open Folder" -UNC
    Windows open folder form starting in C drive root. If folder.dll not present Read-Host will be used.

.INPUTS
    folder.dll is required to be in the same path as the script calling this function.

.OUTPUTS
    Full path of Folder or file    

.NOTES  
    Author     : Glen Buktenica
	Change Log : 20150130 Initial Build  
                 20151005 Public Release 
                 20160304 Description written to console
                 20160407 Updated comments and updated for Export-Logs functions
#> 
Param (
    [parameter(Position=1)][string] $Start,
    [parameter(Position=2)][String] $Description,
    [String] $Ext,
    [Switch] $File,
    [Switch] $Save,
    [Switch] $UNC
)
    Add-Type -AssemblyName System.Windows.Forms
    If ($File)
    {
        If ($Save)
        {
            $OpenForm = New-Object System.Windows.Forms.SaveFileDialog
            If (!$Description)
            {
                $Description = "Select file to save"
            }
        }
        Else
        {
            $OpenForm = New-Object System.Windows.Forms.OpenFileDialog
            If (!$Description)
            {
                $Description = "Select file to open"
            }
        }
        Write-Host $Description
        $OpenForm.InitialDirectory = $Start
        If ($Ext.length -gt 0)
        {
            $OpenForm.Filter = "$Ext files (*.$Ext)|*.$Ext|All files (*.*)|*.*"
        }
        If ($OpenForm.showdialog() -eq "Cancel")
        {
            Write-Error "You pressed cancel, script will now terminate." 
            Export-Logs
            Start-Sleep -Seconds 2
            Exit
        }
        $OpenForm.filename
        $OpenForm.Dispose()
    }
    Else #Open Folder
    {
        $DllPath = (Split-Path $script:MyInvocation.MyCommand.Path) + "\FolderSelect.dll"
        If (!$Description)
        {
            $Description = "Select folder to open"
        }
        Write-Host $Description
        If (Test-Path $DllPath -ErrorAction SilentlyContinue)
        {
            Add-Type -Path $DllPath
            $OpenForm = New-Object -TypeName FolderSelect.FolderSelectDialog -Property @{ Title = $Description; InitialDirectory = $Start }
            $A = $OpenForm.showdialog([IntPtr]::Zero)
            If (!($OpenForm.FileName))
            {
                Write-Error "You pressed cancel, script will now terminate." 
                Export-Logs
                Start-Sleep -Seconds 2
                Exit
            }
            Else
            {
                $OpenForm.FileName
            }
        }
        #If FolderSelect.dll missing fall back to .NET form or Read-Host if UNC forced
        Elseif($UNC)
        {
            $OpenForm = Read-Host $Description
            Return $OpenForm
        }
        Else
        {
            $OpenForm = New-Object System.Windows.Forms.FolderBrowserDialog
            $OpenForm.Rootfolder = $Start
            $OpenForm.Description = $Description

            If ($OpenForm.showdialog() -eq "Cancel")
            {
                Write-Error "You pressed cancel, script will now terminate." 
                Export-Logs
                Start-Sleep -Seconds 2
                Exit
            }
            $OpenForm.SelectedPath
            $OpenForm.Dispose()
        }
    }
}
#endregion Functions
#endregion Header
#region Main
Configure-Logs
Clear-Host

    ##################
    # CODE GOES HERE #
    ##################

#endregion Main
Export-Logs
