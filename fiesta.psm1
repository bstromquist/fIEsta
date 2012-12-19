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


# This file contains the public routines for the module.  All supporting reoutines are included
# at the end following the export statement.


###########################################################
# Initailize-Script uses this to call AppActivate().  When IE launches, input focus is not always
# properly established.  fIEsta resorts to the equivalent of clicking IE and typing a key to get
# it in the right state to start.
Add-Type -AssemblyName microsoft.VisualBasic

###########################################################
#  this is used to call SendKeys()
Add-Type -AssemblyName System.Windows.Forms

###########################################################
# This is used for finding IE windows
$shell = New-Object -comObject Shell.Application

$trace = $false

# by default, results are stored in a $pwd\log\yyyyMMdd_hhmmss directory
#$defaultLogDirBase = "$pwd\log"
#$defaultLogSubdir = Get-Date -format yyyyMMdd_hhmmss
$defaultCfg = @{
	"url"="http://google.com";
	"resultsDir"="$pwd\log";
	"resultsSubdir"= Get-Date -format yyyyMMdd_hhmmss
	"closeAll"=$false;
	"trace"=$false;
}

$logFileName = ""
$showHtmlLog = $true
$autoSnap = $true
$sectionOutput = $null


###########################################################
<#  
.SYNOPSIS  
	Get everthing ready to start testing. 
.DESCRIPTION
    
.PARAMETER x
    ...
.EXAMPLE 
    ...  
#>  
Set-Alias initialize Initialize-Script
function Initialize-Script ( $config=$defaultCfg )
{
	# take some defaults if none given to the passed hash
	if ($config['resultsDir']) {
		$resultsDir = $config['resultsDir'];
	}
	else {
		$resultsDir = $defaultCfg['resultsDir'];
	}
	if ($config['resultsSubdir']) {
		$resultsSubdir = $config['resultsSubdir'];
	}
	else {
		$resultsSubdir = $defaultCfg['resultsSubdir'];
	}

	$script:logDir = "{0}\{1}" -f $resultsDir, $resultsSubdir;

	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line
	$script:trace = $config['trace']
	$script:showIE = !$hidden

	$count = 1;

	# if it doesn't exist, this will (recursively) create $logDir/screenshots 
	# this creates all dirs needed for the test in one fell swoop
	if( !(test-path ${script:logDir}\screenshots -pathType container))
	{
		New-Item ${script:logDir}\screenshots -type directory
	}

	$script:logFileName = "$logDir\$($testScript)";
	while( test-path "$logFileName.html") 
	{
		$count++
		$script:logFileName = "${script:logDir}\$($testScript)_$count"
	}
	
	Add-HtmlLogEntry -t 'start' 
	Write-HtmlLog $config
	
	Write-Banner $config
	
	section "Initialize"
	
	if( $config['closeAll'])
	{
		Close-IE
	}

	# Create an ie com object
	trace "Launching IE..."
	$ie = New-Object -com internetexplorer.application
	$ie.top = 0
	$ie.left = 0
	$ie.visible = $true 
	$script:ie = $ie
	$script:mainHandle = $ie.hwnd

	# Wait for IE to fully launch
	Start-Sleep -milliseconds 500
	$script:proc = Get-Process | where {$_.mainwindowtitle -eq "Windows Internet Explorer"}		
	
	if( $script:proc)
	{
		# For some unknown reason IE dosen't always get focus so force it.
		[void]$script:proc.waitForInputIdle()
		Start-Sleep -milliseconds 500
		
		#TODO put this in front of every sendkeys call
		[Microsoft.VisualBasic.Interaction]::AppActivate( $script:proc.mainwindowtitle)
	}

}

###########################################################
<#  
.SYNOPSIS  
	Close any open IE windows
.DESCRIPTION
    
.PARAMETER x
    ...
.EXAMPLE 
    ...  
#> 
Set-Alias close Initialize-Script
function Close-IE 
{
		# close all IE instances
		trace "Closing all open IE instances..."
		foreach( $win in @($shell.windows()))
		{
			if( $win -and ($win.path -match "Internet Explorer"))
			{
				trace "Closing $($win.path)"
				$win.quit();
			}
		}
		
		# wait until they are all closed
		$try = 0;
		$allClosed = $false
		do 
		{
			$allClosed = $true
			foreach( $win in @($shell.windows()))
			{
				if( $win -and ($win.path -match "Internet Explorer"))
				{
					write-host "IE still closing"
					$allClosed = $false
				}
			}
			
			$try ++
			if ($try -gt 60) 
			{
				break;
			}
			
			Start-Sleep -milliseconds 500
		} 
		while (!$allClosed)
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
Set-Alias pause Set-PauseScript
function Set-PauseScript( $title)
{
	Write-Host "Press any key to continue ..."
	[void]$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
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
Set-Alias selectTab Select-BrowserTab
function Select-BrowserTab( $title)
{
	$try = 0
	$ie2 = $null
	
	$caller = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]

	if( $caller -ne "fiesta") 
	{
		# only do this if called from a user's script
		$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
		$script:testLine = $MyInvocation.ScriptLineNumber
		$script:cmdLine = $MyInvocation.Line
	}
	
	trace "Search for handle $mainHandle with text '$title'"
	
	# There is a separate browser com object for each tab
	# find the explorer instance that matches the given title
	do 
	{
		foreach( $win in @($shell.windows()))
		{
			trace "Testing $($win.hwnd) '$($win.locationName)' '$($win.locationUrl)'"
			if(($win.hwnd -eq $mainHandle) -and ($win.locationName -like $title))
			{
				$ie2 = $win
				trace "found '$title' at $mainHandle"
				break;
			}
			if(($win.hwnd -eq $mainHandle) -and ($win.locationUrl -eq $title))
			{
				$ie2 = $win
				trace "found '$title' at $mainHandle"
				break;
			}
		}
		
		$try ++
		if ($try -gt 2) #60) 
		{
			break;
		}
		else
		{
			Start-Sleep -milliseconds 500
		}
	} 
	while ($ie2 -eq $null)
	
	if( $ie2)
	{
		# found the requested tab, make it the current tab for this script
		$ie2.visible = $true 
		$script:ie = $ie2
		waitForIE
		
		trace "bring '$($ie.document.title)' to the front"
		if( Select-IETab $ie $ie.document.title )
		{
			Write-TestResult $true 		
		}
		else
		{
			Write-TestResult $false 		
		}
	}
	else
	{
		Write-TestResult $false 		
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
Set-Alias finalize Show-Summary
function Show-Summary( $config=$defaultCfg )
{
	Add-HtmlLogEntry -t 'sectionEnd'

	tapComment "$passCount of $testCount tests passed." -type "result"
	
	Add-HtmlLogEntry -t 'end' 
	Write-HtmlLog

	if ($config['closeAll']) {
		Close-IE
	}
	else {
		Show-HtmlLog
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
Set-Alias visit Set-BrowserLocation
function Set-BrowserLocation
{
	param( [alias("l")]  $url=$null,
		   [alias("u")]  $user=$null,
		   [alias("p")]  $pass=$null,
		   [alias("n")]  [switch]$newTab
		 )
		 
	if( $newTab)
	{
		$ie.navigate2( $url, 2048)
		waitForIE
		selectTab $url
	}
	else
	{	
		$ie.navigate( $url)
	}

	if( $user)
	{
		$keys = "$user"
		wait 
		Start-Sleep -milliseconds 500
		
		if( $pass)
		{
			$keys += "{TAB}$pass"
		}
		
		$keys += "{ENTER}"
		trace "Sending key string '$keys'"
		Send-Keys( $keys) 
	}
	
	waitForIE
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
Set-Alias sendKeys Send-Keys
function Send-Keys( $keys)
{
	[Microsoft.VisualBasic.Interaction]::AppActivate( $script:proc.mainwindowtitle)
	[System.Windows.Forms.SendKeys]::SendWait( $keys)
	waitForIE
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
Set-Alias section Set-TestName
function Set-TestName( $name)
{
	if( $testName)
	{
		Add-HtmlLogEntry -t 'sectionEnd'
		Write-HtmlLog
	}
	
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line
	$script:testName = $name

	Add-HtmlLogEntry -t 'sectionStart' -l $name
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
Set-Alias refresh Send-BrowserRefresh
function Send-BrowserRefresh
{
	$ie.refresh()
	waitForIE
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
Set-Alias back Send-BrowserBack
function Send-BrowserBack
{
	$ie.goback()
	waitForIE
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
Set-Alias forward Send-BrowserForward
function Send-BrowserForward
{
	$ie.goforward()
	waitForIE
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
Set-Alias resize Set-BrowserSize
function Set-BrowserSize
{
	param( [alias("w")]  [int32]$width=$null,
		   [alias("h")]  [int32]$height=$null
		 )
		 
	if( $width)
	{
		$ie.width = $width
	}
	
	if( $height)
	{
		$ie.height = $height
	}
	
	waitForIE
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
Set-Alias waitForText Wait-Text
Set-Alias wait Wait-Text
function Wait-Text
{
	param( [alias("t")]  [string]$text=$null,
		   [alias("i")]  [string]$id=$null,
		   [alias("n")]  [string]$name=$null,
		   [alias("a")]  [string]$attr=$null,
		   [alias("av")] [string]$attrValue=$null,
		   [alias("tg")] [string]$tag=$null,
		   [alias("rt")] [string]$rootTag=$null,
		   [alias("ra")] [string]$rootAttr=$null,
		   [alias("rav")][string]$rootAttrValue=$null,
		   [alias("ne" )][switch]$ignoreError,
		   [alias("l"  )][string]$limit=2,
		   [alias("IE" )][switch]$IEBusy
		   )
	
	if( $IEBusy)
	{
		waitForIE
		return
	}
	
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line
	
	if( !$text -and !$id -and !$name -and !$attr -and !$attrValue)
	{
		trace "Waiting for $limit seconds."
		Start-Sleep -milliseconds (1000 * $limit)
		return
	}

	# busy loop for any time ie is loading a page
	$count = 0
	$el =  $null
	
	do
	{
		$el = findElement -t $text -i $id -n $name -a $attr -av $attrValue -tg $tag -rt $rootTag -ra $rootAttr -rav $rootAttrValue
		
		if( $el -and ($id -or $name -or $attr) -and $text)
		{
			# text hasn't been verified yet
			if( !($el.innerText -like $text))
			{
				trace "text match failed"
				$el = $null
			}
			else
			{
				trace "text match passed"
			}
		}
		
		$count++
		trace "waiting $count Sec for '$text$id' to appear."
		Start-Sleep -Milliseconds 1000
	}
	while (!$el -and ( $count -lt $limit))
	
	if( !$el)
	{
		trace "wait timed out"
	}
	else
	{
		trace "wait succeeded"
	}
	
	if( !$ignoreError)
	{
		Write-TestResult ($el -ne $null)
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
Set-Alias waitNo Wait-NoText
function Wait-NoText
{
	param( [alias("t")]  [string]$text=$null,
		   [alias("i")]  [string]$id=$null,
		   [alias("n")]  [string]$name=$null,
		   [alias("a")]  [string]$attr=$null,
		   [alias("av")] [string]$attrValue=$null,
		   [alias("tg")] [string]$tag=$null,
		   [alias("rt")] [string]$rootTag=$null,
		   [alias("ra")] [string]$rootAttr=$null,
		   [alias("rav")][string]$rootAttrValue=$null,
		   [alias("l"  )][string]$limit=2
		   )
	
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line

	# busy loop for any time ie is loading a page
	$count = 0
	$el =  $null
	
	do
	{
		$el = findElement -t $text -i $id -n $name -a $attr -av $attrValue -tg $tag -rt $rootTag -ra $rootAttr -rav $rootAttrValue
		
		if( $el -and ($id -or $name -or $attr) -and $text)
		{
			# text hasn't been verified yet
			if( !($el.innerText -like $text))
			{
				trace "text match failed"
				$el = $null
			}
			else
			{
				trace "text match passed"
			}
		}
		
		$count++
		trace "waiting $count Sec for '$text$id' to disappear."
		Start-Sleep -Milliseconds 1000
	}
	while ($el -and ( $count -lt $limit))
	
	if( $el)
	{
		trace "wait timed out"
	}
	else
	{
		trace "wait succeeded"
	}
		
	Write-TestResult ($el -eq $null)
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
Set-Alias click Send-Click
function Send-Click
{		   
	param( [alias("t"  )][string]$text=$null,
		   [alias("i"  )][string]$id=$null,
		   [alias("n"  )][string]$name=$null,
		   [alias("a"  )][string]$attr=$null,
		   [alias("av" )][string]$attrValue=$null,
		   [alias("tg" )][string]$tag=$null,
		   [alias("rt" )][string]$rootTag=$null,
		   [alias("ra" )][string]$rootAttr=$null,
		   [alias("rav")][string]$rootAttrValue=$null,
		   [alias("dbl")][switch]$double,
		   [alias("ne" )][switch]$ignoreError,
		   [string]$tab=$null
		   )
		   
	$result = $true
	$el = $null
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line
	
	if( $tab)
	{
		Select-BrowserTab $tab
		return
	}

	$el = findElement $text $id $name $attr $attrValue $tag $rootTag $rootAttr $rootAttrValue
	
	if( !$el)
	{
		trace "click() - no matching element found"
	}
	
	if( $el)
	{
		if( $el.tagName -eq 'OPTION')
		{
			$combo = $el.parentNode
			trace "clicking $($combo.tagName) id='$($combo.id)' name='$($combo.name)'."
			$combo.selectedIndex = $el.index
			[void]$combo.fireEvent( "onchange")
		}
		elseif( $el.tagName -eq 'SELECT')
		{
			trace "clicking $($el.tagName) id='$($el.id)' name='$($el.name)'."
			$idx = -1
			foreach( $op in @($el.childNodes))
			{
				if( $op.innerText -eq $text)
				{
					$idx = $op.index
					break
				}
			}

			if( $idx -ne -1)
			{
				trace "select combobox id='$($el.id)' item $idx"
				$el.selectedIndex = $idx
				[void]$el.fireEvent( "onchange")
			}
			else
			{
				trace "'$text' is not a valid selection for select id='$($el.id)'"
				$result = $false
			}
		}
		else
		{
			trace "clicking $($el.tagName) id='$($el.id)' name='$($el.name)'."
			try 
			{
				[void]$el.fireEvent( "onmouseover")
				Start-Sleep -milliseconds 50
				[void]$el.fireEvent( "onmousedown")
				Start-Sleep -milliseconds 50
				[void]$el.fireEvent( "onmouseup")
				Start-Sleep -milliseconds 50
			}
			catch 
			{
				trace "fireEvent threw an error"
			}
			
			
			if( !$double)
			{
				try 
				{
					$el.click()
				}
				catch 
				{
					trace "click threw an error"
				}
			}
			else
			{
				try 
				{
					[void]$el.fireEvent( "ondblclick")
				}
				catch 
				{
					trace "fireEvent ondblclick threw an error"
				}
			}
			
			try 
			{
				[void]$el.fireEvent( "onmouseout")
			}
			catch 
			{
				trace "fireEvent onmouseout threw an error"
			}
		}

		waitForIE
	}
	elseif( !$ignoreError)
	{
		Write-Comment "click() could not find a matching element."	-type "error"	
		$result = $false
	}

	if( !$ignoreError)
	{
		Write-TestResult $result 
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
Set-Alias fill Set-ElementValue
function Set-ElementValue #( $id, $val)
{
 	param( [alias("i"  )][string]$id=$null,
		   [alias("t"  )][string]$text=$null,
		   [alias("n"  )][string]$name=$null,
		   [alias("a"  )][string]$attr=$null,
		   [alias("av" )][string]$attrValue=$null,
		   [alias("tg" )][string]$tag=$null,
		   [alias("rt" )][string]$rootTag=$null,
		   [alias("ra" )][string]$rootAttr=$null,
		   [alias("rav")][string]$rootAttrValue=$null
		  )
			
	$result = $true
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line
	trace "call findElement"
	$el = findElement $null $id $name $attr $attrValue $tag $rootTag $rootAttr $rootAttrValue
	
	# see if id is adjacent text
	if( !$el)
	{
		trace "fill() searching for input element preceded by '$id'"
		$elList = $ie.document.getElementsByTagName( "input")
		$count = @($elList).length
		trace "$count input elements found"
		
		foreach( $tmpEl in @($elList))
		{
			$atext = $tmpEl.getAdjacentText( "beforeBegin")
			trace "fill() testing el with id = '$($tmpEl.id)' after text '$atext'"
		}
		
		$el = $elList | Where-Object { $_.getAdjacentText( "beforeBegin") -like $id}
		
		if( !$el)
		{
			trace "fill() searching for select element preceded by '$id'"
			$el = $ie.document.getElementsByTagName( "select") | Where-Object { $_.getAdjacentText( "beforeBegin") -like $id}
		}
		
		if( $el)
		{
			trace "found $($el.getType())"
			if($el -is [array])
			{
				$el = @($el)[0]
			}
		}
		
		if( $el)
		{
			trace "fill() '$id' found as adjacent text."
		}
	}

	if( $el)
	{
		trace "setting $($el.tagName) id='$($el.id)' name='$($el.name)' to '$text'"
		
		if( $el.tagName -eq "select")
		{
			$idx = -1
			foreach( $op in @($el.childNodes))
			{
				if( $op.innerText -eq $text)
				{
					$idx = $op.index
					break
				}
			}

			if( $idx -ne -1)
			{
				trace "select combobox id='$($el.id)' item $idx"
				$el.selectedIndex = $idx
			}
			else
			{
				trace "'$val' is not a valid selection for select id='$($el.id)'"
				$result = $false
			}
		}
		else
		{
			$el.value = $text
		}
		
		if( $result)
		{
			# fire onkeyup and onchange events here so listeners know data has changed
			[void]$el.fireEvent( "onkeyup")
			[void]$el.fireEvent( "onchange")
		}
	}
	else
	{
		Write-Comment "fill() '$id' not found." -type "error"
		$result = $false
	}

	Write-TestResult $result 
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
Set-Alias assert Test-Assertion
Set-Alias test Test-Assertion
function Test-Assertion
{
	param( [alias("t")]  [string]$text=$null,
		   [alias("i")]  [string]$id=$null,
		   [alias("n")]  [string]$name=$null,
		   [alias("a")]  [string]$attr=$null,
		   [alias("av")] [string]$attrValue=$null,
		   [alias("tg")] [string]$tag=$null,
		   [alias("rt")] [string]$rootTag=$null,
		   [alias("ra")] [string]$rootAttr=$null,
		   [alias("rav")][string]$rootAttrValue=$null,
		                 [switch]$title,
		   [alias("nf")] [switch]$notFound,
		   [alias("v")]  [switch]$visible,
		   [alias("nv")]  [switch]$notVisible
		   )
		   
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line
		
	if( $title)
	{
		trace "assert document title = $text"
		$result = ($ie.document.title -like $text)
	}
	else
	{
		$el = findElement -t $text -id $id -n $name -a $attr -av $attrValue -tg $tag -rt $rootTag -ra $rootAttr - rav $rootAttrValue
		$result = $el
		
		if( $el -and ($id -or $name -or $attr) -and $text)
		{
			# text hasn't been verified yet
			$result = ($el.innerText -like $text)
			trace "test expected '$text' and got '$($el.innerText)'"
		}
	}
	
	if( $visible -and $result)
	{
		$el = $result
		trace "testing visiblity of $($e.tag) element id = '$($el.id)' text = '$($el.innerText)'"
		$result = ($el.offsetWidth -gt 0) -or ($el.offsetHeight -gt 0)
	}

	if( $notVisible -and $result)
	{
		$el = $result
		trace "testing invisiblity of $($e.tag) element id = '$($el.id)' text = '$($el.innerText)'"
		$result = ! $el -or ( ($el.offsetWidth -eq 0) -and ($el.offsetHeight -eq 0) )
	}
			
	if( $notFound)
	{
		$result = !$result
	}
	
	Write-TestResult $result 
}

<###########################################################
 All of the above functions are exported.  Everything below
 This statement in not available to scripts.
#>
Export-ModuleMember -Alias * -Function *

###########################################################
. $psScriptRoot\support.ps1
. $psScriptRoot\output.ps1

###########################################################
. $psScriptRoot\Get-Screenshot.ps1

###########################################################
. $psScriptRoot\Get-IETab.ps1
