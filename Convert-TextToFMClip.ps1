<#
.SYNOPSIS
	Convert text on clipboard to FM clipboard format.
.NOTES
	Author     : Daniel Smith dansmith65@gmail.com
	Requires   : Powershell to be running in single threaded mode (powershell.exe -sta)
.LINK
	https://github.com/dansmith65/FileMaker-PowerShell-Clipboard
.PARAMETER 	Format
	Mac-XMTB = table
	Mac-XMFD = field
	Mac-XMSC = script
	Mac-XMSS = script step
	Mac-XMFN = custom function
	Mac-XMLO = layout object (.fp7)
	Mac-XML2 = layout object (.fmp12)
	Mac-     = Theme
#>

param (
    [Parameter(Mandatory=$True,Position=0)]
    [string]$Format
)


##########################################################
# dependencies
##########################################################
Add-Type -AssemblyName System.Windows.Forms



##########################################################
# functions
##########################################################
function Show-Message {
    param (
        [string]$Message,
        [int]$Milliseconds = 0
            # if >0, will automatically timeout after specified time
    )
    Write-Host $Message
    Write-Host
    Write-Host "press any key to continue"
    $start = $(Get-Date)
    do {
        Start-Sleep -milliseconds 100
        If ($Milliseconds -gt 0)
        {
            $elapsed = New-Timespan $start $(Get-Date)
            If ($elapsed.TotalMilliseconds -gt $Milliseconds) {break}
        }
    } until ([console]::KeyAvailable)
    $Host.UI.RawUI.FlushInputBuffer()
}



##########################################################
# create header
##########################################################
Write-Host "#############################################################################"
Write-Host "#" $MyInvocation.MyCommand.Path
Write-Host "# Convert text on clipboard to FM clipboard format."
Write-Host "#############################################################################"
Write-Host 
if ([threading.thread]::CurrentThread.GetApartmentState() -eq "MTA") {
   Show-Message "must be called in single threaded mode (powershell.exe -sta)"
   Exit
}



##########################################################
# read data from clipboard
##########################################################
If (! [System.Windows.Forms.Clipboard]::ContainsData("Text"))
{
    Show-Message "NOT CONVERTED: Text not on the clipbard!" -Milliseconds 3000
    Exit
}
$clipText = [System.Windows.Forms.Clipboard]::GetText("Text")




##########################################################
# setup
##########################################################
$encoding = [System.Text.Encoding]::UTF8
# for FileMaker types, the first 4 bytes on the clipboard is the size of the data on the clipboard
$offset = 4
$clipLength = $clipText.length
$totalLength = $offset + $clipText.length



##########################################################
# write to stream
##########################################################
$lengthAsBytes = [BitConverter]::GetBytes($clipLength)
$clipAsBytes = $encoding.GetBytes($clipText)

$memStream = New-Object System.IO.MemoryStream($totalLength)

$memStream.Write(
    $lengthAsBytes,
    0,
    $offset
)

$memStream.Write(
    $clipAsBytes,
    0,
    $clipLength
)



##########################################################
# Save data to the clipboard
#    https://msdn.microsoft.com/en-us/library/system.windows.forms.clipboard.setdata(v=vs.110).aspx?
##########################################################

[System.Windows.Forms.Clipboard]::SetData($format, $memStream)

# Show-Message "converted text to $format" -Milliseconds 2000
