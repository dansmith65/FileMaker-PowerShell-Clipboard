<#
.SYNOPSIS
	Convert FileMaker clipboard format both to and from text. Either input
	will result in clipboard being able to be pasted into either FileMaker
	or a text editor.
	Input/output formats will be automatically detected.

.PARAMETER 	prettyPrint
	When converting from FM to text format, pretty print the XML.
	Convert-FMClip.prettyPrint environment variable will override this parameter.

.PARAMETER 	prettyPrintChar
	Character to use for indentation: "space", "tab", "cr", or "lf"
	Convert-FMClip.prettyPrintChar environment variable will override this parameter.

.PARAMETER 	prettyPrintIndentation
	Number of prettyPrintChar to use for each level.
	Convert-FMClip.prettyPrintIndentation environment variable will override this parameter.

.NOTES
	Version:   2.0.0
	Author:    Daniel Smith dan@filemaker.consulting
	Requires:  Powershell to be running in single threaded mode (powershell.exe -sta)

.LINK
	https://github.com/dansmith65/FileMaker-PowerShell-Clipboard
#>

param (
	[bool]$prettyPrint = $true,
	[string]$prettyPrintChar = "tab",
	[int]$prettyPrintIndentation = 1
)

Try {
<#
	This script is expected to be called directly via hotkey. This means the console will close as
	soon as the script exits. That wouldn't allow users to see error messages before the window
	closed. Therefore, all code in this script was put into a Try block which allows the Catch
	block to display the message then pause to allow users to view it before the console closes.
#>



##########################################################
# functions
##########################################################
function Format-XML {
<# source: https://stackoverflow.com/a/39271782/1327931 #>
	Param (
		[Parameter(ValueFromPipeline=$true,Mandatory=$true)][string]$xmlcontent,
		[char]$indentChar,
		[int]$indentation
	)
	$xmldoc = New-Object -TypeName System.Xml.XmlDocument
	$xmldoc.LoadXml($xmlcontent)
	$sw = New-Object System.IO.StringWriter
	$writer = New-Object System.Xml.XmlTextwriter($sw)
	$writer.Formatting = [System.XML.Formatting]::Indented
	if ($indentChar) {
		$writer.IndentChar = $indentChar
	}
	if ($indentation) {
		$writer.Indentation = $indentation
	}
	$xmldoc.WriteContentTo($writer)
	$sw.ToString()
}
function Test-IsInteractiveShell {
<#
	.SYNOPSIS
		Returns boolean determining if prompt was run noninteractive.
	.DESCRIPTION
		First, we check `[Environment]::UserInteractive` to determine if we're if the shell if running 
		interactively. An example of not running interactively would be if the shell is running as a service.
		If we are running interactively, we check the Command Line Arguments to see if the `-NonInteractive` 
		switch was used; or an abbreviation of the switch.
	.LINK
		https://github.com/UNT-CAS/Test-IsNonInteractiveShell
		(function was modified from this version before adding to GetSSL.ps1)
#>
	return ([Environment]::UserInteractive -and (-not ([Environment]::GetCommandLineArgs() | ?{ $_ -like '-NonI*' })))
}



##########################################################
# create header
##########################################################
Write-Host "#############################################################################"
Write-Host "#" $MyInvocation.MyCommand.Path
Write-Host "# Convert FileMaker clipboard format both to and from text."
Write-Host "#############################################################################"



##########################################################
# validate
##########################################################
if ([threading.thread]::CurrentThread.GetApartmentState() -eq "MTA") {
	throw "must be called in single threaded mode (powershell.exe -sta)"
}



##########################################################
# set default parameters
##########################################################
if (${Env:Convert-FMClip.prettyPrint} -ne $null)
{
	$prettyPrint = [System.Convert]::ToBoolean(${Env:Convert-FMClip.prettyPrint})
	Write-Debug "prettyPrint loaded from environment: $prettyPrint"
}

if (${Env:Convert-FMClip.prettyPrintChar} -ne $null)
{
	$prettyPrintChar =${Env:Convert-FMClip.prettyPrintChar}
	Write-Debug "prettyPrintChar loaded from environment: $prettyPrintChar"
}

if (${Env:Convert-FMClip.prettyPrintIndentation} -ne $null)
{
	$prettyPrintIndentation =${Env:Convert-FMClip.prettyPrintIndentation}
	Write-Debug "prettyPrintIndentation loaded from environment: $prettyPrintIndentation"
}

if ($prettyPrint)
{
	Write-Host "pretty print $prettyPrintIndentation $prettyPrintChar"
	$prettyPrintCharMap = @{
		tab   = 9
		space = 32
		lf    = 10
		cr    = 13
	}
	if ($prettyPrintCharMap[$prettyPrintChar] -eq $null)
	{
		throw "invalid prettyPrintChar"
	}
}
else
{
	Write-Host "don't pretty print"
}
Write-Host



##########################################################
# dependencies
##########################################################
Add-Type -AssemblyName System.Windows.Forms



##########################################################
# check if FileMaker format exists on clipboard
# NOTE: $formatType isn't really needed; it's mostly for documentation and messages to user
##########################################################
$fromFormat = $formatType = $null
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
$toFormat = $null
if (! $fromFormat)
{
	if (! [System.Windows.Forms.Clipboard]::ContainsText())
	{
		throw "No FM or text format on clipboard!"
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
		throw "Clipboard did not contain FM compatible format"
	}

	if ($toFormat)
	{
		Write-Host "Text Format Detected: $toFormat"
	}
	else
	{
		throw "Could not determine clipboard format"
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
	if ($prettyPrint)
	{
		$textClip = (Format-XML $textClip -indentation $prettyPrintIndentation -indentChar $prettyPrintCharMap[$prettyPrintChar])
	}
	
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
	throw "No FM or Text format detected, but I don't think this section of code should ever run; it's just a safety-catch"
}



##########################################################
# save data back to the clipboard
##########################################################
# Write-Host ""; Write-Host "save these formats to the clipboard:"; $newClip.GetFormats();
[System.Windows.Forms.Clipboard]::SetDataObject($newClip, $true)


} Catch {
	# show error message to user
	Write-Host $error[0].Exception.Message -BackgroundColor Black -ForegroundColor Red
	Write-Host

	# alternative to Pause, which allows user to press any key
	if (!($psISE) -and (Test-IsInteractiveShell))
	{
		Write-Host -NoNewLine 'Press any key to continue...';
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
		Write-Host
	}

	# throw the error again so a calling script can access it
	Throw
}
