<#
.SYNOPSIS
	Convert FileMaker clipboard format both to and from text. Either input
	will result in clipboard being able to be pasted into either FileMaker
	or a text editor.
	Input/output formats will be automatically detected.
.NOTES
	Author   : Daniel Smith dan@filemaker.consulting
	Requires : Powershell to be running in single threaded mode (powershell.exe -sta)
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
function Throw-Message {
<#
	This script is expected to be called directly via hotkey. This means the console will close as
	soon as the script exits. That wouldn't allow users to see error messages before the window
	closed. This function is meant to be used instead of throwing an error directly, which will
	allow users to view the message before the console closes.
#>
	param (
		[string]$Message
	)
	Write-Host $Message -BackgroundColor Black -ForegroundColor Red
	Write-Host
	
	# alternative to Pause, which allows user to press any key
	if (! $psISE)
	{
		Write-Host -NoNewLine 'Press any key to continue...';
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
		Write-Host
	}
	
	Throw $Message
}



##########################################################
# create header
##########################################################
Write-Host "#############################################################################"
Write-Host "#" $MyInvocation.MyCommand.Path
Write-Host "# Convert FileMaker clipboard format both to and from text."
Write-Host "#############################################################################"
Write-Host 
if ([threading.thread]::CurrentThread.GetApartmentState() -eq "MTA") {
	Throw-Message "must be called in single threaded mode (powershell.exe -sta)"
}



##########################################################
# check if FileMaker format exists on clipboard
# NOTE: $formatType isn't really needed; it's mostly for documentation and messages to user
##########################################################
if ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMTB"))
{
	$fromFormat = "Mac-XMTB"
	$formatType = "BaseTable"
}
elseif ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMFN"))
{
	$fromFormat = "Mac-XMFN"
	$formatType = "CustomFunction"
}
elseif ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMFD"))
{
	$fromFormat = "Mac-XMFD"
	$formatType = "Field"
}
elseif ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XML2"))
{
	$fromFormat = "Mac-XML2"
	$formatType = "Layout"
}
elseif ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMSS"))
{
	$fromFormat = "Mac-XMSS"
	$formatType = "Step"
}
elseif ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMSC"))
{
	$fromFormat = "Mac-XMSC"
	$formatType = "Script"
}
elseif ([System.Windows.Forms.Clipboard]::ContainsData("Mac-"))
{
	$fromFormat = "Mac-"
	$formatType = "Theme"
}
elseif ([System.Windows.Forms.Clipboard]::ContainsData("Mac-XMVL"))
{
	$fromFormat = "Mac-XMVL"
	$formatType = "ValueList"
}

if ($fromFormat)
{
	Write-Host "FM Format Detected: $fromFormat ($formatType)"
}



##########################################################
# check if text xml format that appears to be FM compatible exists on the clipboard
##########################################################
if (! $fromFormat)
{
	if (! [System.Windows.Forms.Clipboard]::ContainsText())
	{
		Throw-Message "No FM or text format on clipboard!"
	}
	$textClip = [System.Windows.Forms.Clipboard]::GetText()

	Try
	{
		$result = Select-Xml -Content $textClip -Xpath "/fmxmlsnippet[@type='FMObjectList']/*[1]" -ErrorAction Stop
		if (! $result) {
			$result = Select-Xml -Content $textClip -Xpath "/fmxmlsnippet[@type='LayoutObjectList']/*[1]"-ErrorAction Stop
		}
		if ($result) {
			$formatType = $result.Node.ToString()
			Write-Host "Text FormatType Detected: $formatType"
		}
		if ($formatType -eq "BaseTable")
		{
			$toFormat = "Mac-XMTB"
		}
		elseif ($formatType -eq "CustomFunction")
		{
			$toFormat = "Mac-XMFN"
		}
		elseif ($formatType -eq "Field")
		{
			$toFormat = "Mac-XMFD"
		}
		elseif ($formatType -eq "Layout")
		{
			$toFormat = "Mac-XML2"
		}
		elseif ($formatType -eq "Step")
		{
			$toFormat = "Mac-XMSS"
		}
		elseif ($formatType -eq "Script")
		{
			$toFormat = "Mac-XMSC"
		}
		elseif ($formatType -eq "Group") # script folder
		{
			$toFormat = "Mac-XMSC"
		}
		elseif ($formatType -eq "Theme")
		{
			$toFormat = "Mac-"
		}
		elseif ($formatType -eq "ValueList")
		{
			$toFormat = "Mac-XMVL"
		}
	}
	Catch
	{
		Throw-Message "Clipboard did not contain FM compatible format"
	}

	if ($toFormat)
	{
		Write-Host "Text Format Detected: $toFormat"
	}
	else
	{
		Throw-Message "Could not determine clipboard format"
	}
}



##########################################################
# get data from clipboard
# https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.clipboard.getdataobject
##########################################################
$oldClip = [System.Windows.Forms.Clipboard]::GetDataObject()
# for some unknown reason, I can't directly edit the oldClip so I'm copying each format to a new clip
$newClip = New-Object System.Windows.Forms.Dataobject
foreach($format in $oldClip.GetFormats())
{
	# Don't add "Clipboard Viewer Ignore" to the new clipboard because I want Ditto,
	# a clipboard manager, to record the new clip
	if ($format -ne "Clipboard Viewer Ignore")
	{
		$data = $oldClip.GetData($format)
		$newClip.SetData($format, $data)
	}
}




##########################################################
# convert between FM and Text
##########################################################
if ($fromFormat)
{
	$fmClip = $oldClip.GetData($fromFormat)

	# first 4 bytes on the clipboard is the size of the data on the clipboard
	$offset = 4
	$encoding = [System.Text.Encoding]::UTF8

	# reading a stream:
	# https://msdn.microsoft.com/en-us/library/system.io.stream.read(v=vs.110).aspx
	$fmClip.Position = $offset
	$length = $fmClip.Length - $offset
	$buffer = New-Object byte[]($length)
	$readLength = $fmClip.Read($buffer, 0, $length)
	$dispose = $fmClip.Dispose
	$textClip = $encoding.GetString($buffer)
	
	# add textClip to data object
	# auto-convert to other formats (like text, oemtext, etc.)
	$newClip.SetData([System.Windows.Forms.DataFormats]::UnicodeText, $true, $textClip)
}
elseif ($toFormat)
{
	# $textClip should have been set in above section before determining the "toFormat"
	# first 4 bytes on the clipboard is the size of the data on the clipboard
	$offset = 4
	$encoding = [System.Text.Encoding]::UTF8
	
	# write to stream
	$clipLength = $textClip.length
	$totalLength = $offset + $textClip.length
	$lengthAsBytes = [BitConverter]::GetBytes($clipLength)
	$clipAsBytes = $encoding.GetBytes($textClip)
	$fmClip = New-Object System.IO.MemoryStream($totalLength)
	$fmClip.Write($lengthAsBytes, 0, $offset)
	$fmClip.Write($clipAsBytes, 0, $clipLength)
	
	# add fmClip to data object
	# don't auto-convert to other formats
	$newClip.SetData($toFormat, $false, $fmClip)
}
else
{
	Throw-Message "No FM or Text format detected, but I don't think this section of code should ever run; it's just a safety-catch"
}



##########################################################
# save data back to the clipboard
##########################################################
# Write-Host ""; Write-Host "save these formats to the clipboard:"; $newClip.GetFormats();
[System.Windows.Forms.Clipboard]::SetDataObject($newClip, $true)
