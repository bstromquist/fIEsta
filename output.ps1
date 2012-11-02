# Permission is hereby granted, free of charge, to any person obtaining
# a copy of this software and associated documentation files (the
# "Software"), to deal in the Software without restriction, including
# without limitation the rights to use, copy, modify, merge, publish,
# distribute, sublicense, and/or sell copies of the Software, and to
# permit persons to whom the Software is furnished to do so, subject to
# the following conditions:

# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
# LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
# OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
# WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

<#  
.SYNOPSIS  

.DESCRIPTION
    
.PARAMETER x
    ...
.EXAMPLE 
    ...  
#>  
function Write-Banner( $url)
{
	$strComputer = gc env:computername;

	$ieversion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Internet Explorer').Version
#	$colItems = get-wmiobject -class "MicrosoftIE_Summary" -namespace "root\CIMV2\Applications\MicrosoftIE" `
#	-computername $strComputer

	$objItem = @($colItems)[0];
	write-host "###########################################################"  	-foregroundcolor White -backgroundcolor DarkGreen	
	write-host "# Date:           " (Get-Date)		  						 	-foregroundcolor White -backgroundcolor DarkGreen	
	write-host "# Computer:       " $strComputer								-foregroundcolor White -backgroundcolor DarkGreen	
	write-host "# Version:        " $ieversion									-foregroundcolor White -backgroundcolor DarkGreen	
	write-host "# URL:            " $url										-foregroundcolor White -backgroundcolor DarkGreen	
	write-host "###########################################################"	-foregroundcolor White -backgroundcolor DarkGreen	
	
	Add-HtmlLogEntry -t 'bannerStart'
	Add-HtmlLogEntry -t 'bannerEntry' -label "Date" -value (Get-Date)
	Add-HtmlLogEntry -t 'bannerEntry' -label "Computer" -value $strComputer
	Add-HtmlLogEntry -t 'bannerEntry' -label "IE Version" -value $ieversion
	Add-HtmlLogEntry -t 'bannerEntry' -label "URL" -value $url
	Add-HtmlLogEntry -t 'bannerEnd' 
	Write-HtmlLog
}

###########################################################
Set-Alias Comment Write-Comment
function Write-Comment( $str, $type="plain")
{
	Write-TAPComment $str $type
	Add-HtmlLogEntry -t 'comment' -v $str -l $type
}

###########################################################
Set-Alias tapComment Write-TAPComment
function Write-TAPComment( $str, $type="")
{
	$color = "White"
	
	if( $type -eq 'error')
	{
		$color = 'Red'
		write-host "# ERROR: " -NoNewline   -foregroundcolor $color
	}
	elseif( $type -eq 'trace')
	{
		$color = 'Yellow'
		write-host "# TRACE: " -NoNewline   -foregroundcolor $color
	}
	else
	{
		$color = "White"
		write-host "# " -NoNewline   -foregroundcolor $color
	}
	
	write-host $str  -foregroundcolor $color
}

###########################################################
Set-Alias trace Trace-Message
function Trace-Message
{
	param( 
		[alias("t")][string]$text=$null, 
		[switch]$on, 
		[switch]$off
	)
	
	if( $on)
	{
		$script:trace = $true
	}
	
	if( $script:trace -and $text)
	{
		tapComment "$($testScript):$testLine $text" -type "trace"
	}
	
	if( $off)
	{
		$script:trace = $false
	}
}

###########################################################
$testCount = 0
$passCount = 0

###########################################################
Set-Alias testResult Write-TestResult
function Write-TestResult( $pass, $description)
{
	$script:testCount += 1
	
	if( !$description)
	{
		$description = $cmdLine
	}
	
	if( $pass)
	{
		$script:passCount += 1;
	}
	
	Write-TAPTestResult $pass $description
	Write-HtmlTestResult $pass $description
}	

###########################################################
Set-Alias tapTestResult Write-TAPTestResult
function Write-TAPTestResult( $pass, $description)
{
	if( $pass)
	{
		write-host "ok $testCount -" "$testScript`:$testLine $testName $description"  -foregroundcolor green
	}
	else
	{
		write-host "not ok $testCount -" "$testScript`:$testLine $description"  -foregroundcolor red
	}
}	

###########################################################
Set-Alias htmlTestResult Write-HtmlTestResult
function Write-HtmlTestResult( $pass, $description)
{
	if( $autoSnap)
	{
		Write-Screenshot -noLog
	}
	
	if( $pass)
	{
		Add-HtmlLogEntry -t 'testEntry' -l "ok" -v $description
	}
	else
	{
		Add-HtmlLogEntry -t 'testEntry' -l "not ok" -v $description
	}	
}	

###########################################################
function Add-HtmlLogEntry
{
	param( 
		[alias("t")][string]$type=$null,
		[alias("l")][string]$label=$null,
		[alias("v")][string]$value=$null
	)
	
	if( !$sectionOutput)
	{
		$script:sectionOutput = @()
	}
	
	$html = $script:sectionOutput
	
	if( $type -eq 'bannerStart')
	{
		$html += "<table class='bannerTable'>"
	}
	elseif( $type -eq 'bannerEnd')
	{
		$html += "</table>"
	}
	elseif( $type -eq 'bannerEntry')
	{
		$html += "<tr class='bannerEntry'><th class='bannerLabel'>$label</th><td class='bannerValue'>$value</td></tr>"
	}
	elseif( $type -eq 'sectionStart')
	{
		$html += "<h2 class='sectionTitle'>$label</h2>"
		$html += "<table class='sectionTable'>"
	}
	elseif( $type -eq 'comment')
	{
		$html += "<tr>"
		$html +=   "<td></td>"
		$html +=   "<td></td>"
		$html +=   "<td>$testScript</td>"
		$html +=   "<td>$testLine</td>"
		$html +=   "<td class='comment-$label'>$value</td>"
		$html +=   "<td></td>"
		$html += "</tr>"
	}
	elseif( $type -eq 'testEntry')
	{
		$html += "<tr class='testEntry'>"
		if( $label -eq 'ok')
		{
			$html += "<td class='pass'>ok</td>"
		}
		else
		{
			$html += "<td class='fail'>not ok</td>"
		}
		$html += "<td class='testNumber'>$testCount</td>"
		$html += "<td>$testScript</td>"
		$html += "<td>$testLine</td>"
		$html += "<td>$value</td>"
		if( $autoSnap)
		{
			$html += "<td class='thumb'>"
			$html += "  <a href='screenshots/$ssFileName'><img src='screenshots/$ssFileName' title='$ssFileName'/></a>"
#			$html += "  <a href='$ssFileName'><img src='file:///$logDir/$ssFileName' title='$ssFileName'/></a>"
			$html += "</td>"
		}
		else
		{
			$html += "<td></td>"
		}
		$html += "</tr>"
	}
	elseif( $type -eq 'sectionEnd')
	{
		$html += "</table>"
	}
	elseif( $type -eq 'screenshot')
	{
		$html += "<tr class='testEntry'>"
		$html +=   "<td></td>"
		$html +=   "<td class='testNumber'>$testCount</td>"
		$html +=   "<td>$testScript</td>"
		$html +=   "<td>$testLine</td>"
		$html +=   "<td>$value</td>"
		$html +=   "<td class='thumb'>"
		$html +=     "<a href='screenshots/$ssFileName'><img src='screenshots/$ssFileName' title='$ssFileName'/></a>"
#		$html +=     "<a href='$ssFileName'><img src='file:///$logDir/$ssFileName' title='$ssFileName'/></a>"
		$html +=   "</td>"
		$html += "</tr>"
	}
	elseif( $type -eq 'start')
	{
		$html += '<!doctype html>'
		$html += '<html><head><title>Test Summary</title></head>'
		$html += '<body>'
		$html += '<style>'
		$html += 'table{border-style:none;border-width:0px;font-size:8pt;background-color:#ccc;width:100%;}'
		$html += 'th{text-align:right;}'
		#$output += 'td{background-color:#fff;border-style:dotted;border-width:1px;}'
		$html += 'td.thumb{height:200px; float:right;}'
		$html += 'tr.testEntry td.pass, td.fail{width:50px}'
		$html += 'td.pass{background-color:#00FF00;}'
		$html += 'td.fail{background-color:#FF0000;}'
		$html += 'td.comment-error{background-color:#FF0000;}'
		$html += 'td.comment-trace{background-color:#FFFF00;}'
		$html += 'body{font-family:verdana;font-size:8pt;}'
		$html += 'h1{font-size:14pt;}'
		$html += 'h2{font-size:12pt;}'
		$html += 'img{height: 100%; vericle-align: middle}'
		$html += '</style>'
		$html += '<h1>Test Log</h1>'
	}
	elseif( $type -eq 'end')
	{
		$html += "<h1 class='$type'>$passCount of $testCount tests passed.</h1>"
		$html += '</body></html>'
	}
	
	$script:sectionOutput = $html
}

###########################################################
function Write-HtmlLog
{
	$count = @($sectionOutput).length
	$script:sectionOutput | Out-File "$logFileName.html" -Append -Force
	$script:sectionOutput = $null
	
	create-MHT "$logFileName.html" "$logFileName.mht"
}

###########################################################
function Show-HtmlLog
{
	if( $showHtmlLog)
	{
		$ie.navigate2("$logFileName.html", 2048, "Test Summary")
	}
}

###########################################################
<#  
.SYNOPSIS  

.DESCRIPTION
    
.PARAMETER x
    ...
.EXAMPLE 
    ...  
#>  
Set-Alias screenshot Write-Screenshot
function Write-Screenshot( $description, [switch]$noLog)
{
	if( !$noLog)
	{
		# only do this if called from a user's script
		$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
		$script:testLine = $MyInvocation.ScriptLineNumber
		$script:cmdLine = $MyInvocation.Line
	}
	$script:ssFileName = "$testScript-$testLine.png"
	$count = 1;
	
	if( !(test-path "$logDir\screenshots" -pathType container))
	{
		New-Item "$logDir\screenshots" -type directory
	}
	
	while( test-path "$logDir\screenshots\$ssFileName") 
	{
		$count++
		$script:ssFileName = "$testScript-$($testLine)_$count.png"
	}
	
	WaitForIE
	Get-ScreenShot -ie $ie -file "$logDir\screenshots\$ssFileName" 
#	Get-ScreenShot -ie $ie -file "$logDir\$ssFileName" 
	
	if( !$noLog)
	{
		Add-HtmlLogEntry -t "screenshot" -v $description
	}
}


function Create-MHT($htmlPath,$mhtPath)
{
	$adSaveCreateNotExist = 1
	$adSaveCreateOverWrite = 2
	$adTypeBinary = 1
	$adTypeText = 2
	
	$msg = New-Object -ComObject CDO.Message
	$msg.CreateMHTMLBody( $htmlPath, 0)
	
	$strm = New-Object -ComObject ADODB.Stream
	# $strm.Type = $adTypeBinary
	$strm.Type = $adTypeText
	$strm.Charset = "US-ASCII"
	$strm.Open()
	$dsk = $msg.DataSource
	$dsk.SaveToObject( $strm, "_Stream")

	$strm.SaveToFile($mhtPath, $adSaveCreateOverWrite)
}


