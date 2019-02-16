<# Save configuration for the Convert-FMClip.ps1 script to environment variables.
 # Once reviewing the values below, run this script to set them.
 #
 # NOTE: You do NOT need to run this if you choose to use the default options.
 ###############################################################################>


setx Convert-FMClip.prettyPrint $null
	# $true or $false, Default is $true
	# If True, will pretty-print the XML after converting from FM to XML

setx Convert-FMClip.prettyPrintChar "tab"
	# "space", "tab", "cr", "lf".
	# Character to use for indentation. Default is "tab".

setx Convert-FMClip.prettyPrintIndentation 1
	# Number of prettyPrintChar to use for each level. The default is 1.


<###############################################################################>
Write-Host
Write-Host "You will need to re-launch any currently open application before they will see these new values."
pause
