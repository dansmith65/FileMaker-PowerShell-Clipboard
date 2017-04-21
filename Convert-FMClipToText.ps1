<#
.SYNOPSIS
	Convert FileMaker clipboard format to text.
.NOTES
	Author     : Daniel Smith dansmith65@gmail.com
	Requires   : Powershell to be running in single threaded mode (powershell.exe -sta)
.LINK
	https://github.com/dansmith65/FileMaker-PowerShell-Clipboard
#>


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
Write-Host "# Convert FileMaker clipboard format to text."
Write-Host "#############################################################################"
Write-Host 
if ([threading.thread]::CurrentThread.GetApartmentState() -eq "MTA") {
   Show-Message "must be called in single threaded mode (powershell.exe -sta)"
   Exit
}


##########################################################
# check if FileMaker format exists on clipboard
##########################################################
If ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMTB"))
{
    $format = "Mac-XMTB"
    $formatDescription = "Table"
}
ElseIf ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMFD"))
{
    $format = "Mac-XMFD"
    $formatDescription = "Field"
}
ElseIf ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMSC"))
{
    $format = "Mac-XMSC"
    $formatDescription = "Script"
}
ElseIf ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMFN"))
{
    $format = "Mac-XMFN"
    $formatDescription = "Custom Function"
}
ElseIf ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMSS"))
{
    $format = "Mac-XMSS"
    $formatDescription = "Script Step"
}
ElseIf ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XML2"))
{
    $format = "Mac-XML2"
    $formatDescription = "Layout Object .fmp12"
}
ElseIf ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMLO"))
{
    $format = "Mac-XMLO"
    $formatDescription = "Layout Object .fp7"
}
ElseIf ([System.Windows.Forms.Clipboard]::ContainsData("Mac-"))
{
    $format = "Mac-"
    $formatDescription = "Theme"
}
ElseIf ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMVL"))
{
    $format = "Mac-XMVL"
    $formatDescription = "value list (FM16)"
}
Else
{
    Show-Message "NOT CONVERTED: FileMaker formatted data not on clipbard!" -Milliseconds 3000
    Exit
}



##########################################################
# read data from clipboard
#    https://msdn.microsoft.com/en-us/library/system.windows.forms.clipboard.getdata
##########################################################
$clipStream = [System.Windows.Forms.Clipboard]::GetData($format)

# for FileMaker types, the first 4 bytes on the clipboard is the size of the data on the clipboard
$offset = 4
$clipStream.Position = $offset

# reading a stream:
#    https://msdn.microsoft.com/en-us/library/system.io.stream.read(v=vs.110).aspx
$length = $clipStream.Length - $offset
$buffer = New-Object byte[]($length)
$readLength = $clipStream.Read($buffer, 0, $length)
$dispose = $clipStream.Dispose

$encoding = [System.Text.Encoding]::UTF8
$clip = $encoding.GetString($buffer)



##########################################################
# save back to clipboard as text
##########################################################
[System.Windows.Forms.Clipboard]::SetText($clip)


#Show-Message "converted $formatDescription to text" -Milliseconds 5000
