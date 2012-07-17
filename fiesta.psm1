###########################################################
Add-Type -AssemblyName microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

###########################################################
$shell = New-Object -comObject Shell.Application
$jpegCodec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.FormatDescription -eq "JPEG" }
$trace = $false
$logDir = "$pwd\log"
$logFileName = ""
$showHtmlLog = $true
$autoSnap = $true
$sectionOutput = $null

###########################################################
Set-Alias banner Write-Banner
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
Set-Alias initialize Initialize-Script
function Initialize-Script
{
	param(
		[string]$url,
		[bool]$closeAll=$false,
		[bool]$trace=$false
	)
	
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line
	$script:trace = $trace
	$script:showIE = !$hidden

	$count = 1;
	if( !(test-path $logDir -pathType container))
	{
		New-Item $logDir -type directory
	}
	
	$script:logFileName = "$logDir\$($testScript)_log.html"
	while( test-path $logFileName) 
	{
		$count++
		$script:logFileName = "$logDir\$($testScript)_log_$count.html"
	}
	
	Add-HtmlLogEntry -t 'start' 
	Write-HtmlLog
	
	banner $url
	
	section "Initialize"
	
	if( $closeAll)
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
	}
	
	# Create an ie com object
	trace "Launching IE..."
	$ie = New-Object -com internetexplorer.application
	$ie.top = 0
	$ie.left = 0
	$ie.navigate( "about:blank")
	$ie.visible = $true 
	$script:ie = $ie
	$script:mainHandle = $ie.hwnd
}

###########################################################
Set-Alias pause Set-PauseScript
function Set-PauseScript( $title)
{
	Write-Host "Press any key to continue ..."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

###########################################################
Set-Alias selectTab Select-BrowserTab
function Select-BrowserTab( $title)
{
	$try = 0
	$ie2 = $null
	
	write-host "Search for handle $mainHandle with text '$title'"
	
	# There is a separate browser com object for each tab
	# find the explorer instance that matches the given title
	do 
	{
		Start-Sleep -milliseconds 500
		
		foreach( $win in @($shell.windows()))
		{
			write-host "Testing $($win.hwnd) '$($win.locationName)'"
			if(($win.hwnd -eq $mainHandle) -and ($win.locationName -like $title))
			{
				$ie2 = $win
				#break;
				write-host "found '$title' at $mainHandle"
			}
		}
		
		$try ++
		if ($try -gt 2) #60) 
		{
			break;
		}
	} 
	while ($ie2 -eq $null)
	
	if( $ie2)
	{
		# found the requested tab, make it the current tab for this script
		$ie2.visible = $true 
		$script:ie = $ie2
		waitForIE
		Select-IETab $ie $title | out-null
	
		Write-TestResult $true 		
	}
	else
	{
		Write-TestResult $false 		
	}
}

###########################################################
Set-Alias finalize Show-Summary
function Show-Summary
{
	Add-HtmlLogEntry -t 'sectionEnd'

	tapComment "$passCount of $testCount tests passed." -type "result"
	
	Add-HtmlLogEntry -t 'end' 
	Write-HtmlLog
	Show-HtmlLog
}

###########################################################
Set-Alias visit Set-BrowserLocation
function Set-BrowserLocation( $url)
{
	$ie.navigate( $url)
	waitForIE
}

###########################################################
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
Set-Alias waitForIE Wait-IEBusy
function Wait-IEBusy
{
	# busy loop for any time ie is loading a page
	while( $ie.Busy -eq $true)
	{
		trace "IE is busy..."
		Start-Sleep -Milliseconds 500
	}
}

###########################################################
Set-Alias refresh Send-BrowserRefresh
function Send-BrowserRefresh
{
	$ie.refresh()
	waitForIE
}

###########################################################
Set-Alias back Send-BrowserBack
function Send-BrowserBack
{
	$ie.goback()
	waitForIE
}

###########################################################
Set-Alias forward Send-BrowserForward
function Send-BrowserForward
{
	$ie.goforward()
	waitForIE
}

###########################################################
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

function getProperty( [System.__ComObject]$obj, [string]$prop)
{
	[System.__ComObject].InvokeMember( $prop, [System.Reflection.BindingFlags]::GetProperty, $null, $obj, $null)
}

###########################################################
<#
	findElement "Login"   
					search for "Login" as innerText of common tags, 
					then ID value of any element 
					then name value of any element
					??? then value of any attr of any element
					
	findElement -text "Login"   
	findElement -text "Login" -id "btn1"  
	findElement -text "Login" -name "button1"
	findElement -text "Next" -attr "title" -attrValue "Press to continue"  
	
	findElement "Login" -attr "title"
					search for any element where the title attribute is "Login"
					and the innerText is "Login"
					
	findElement "Login" -tag "div"
					search for DIV element where the innerText is "Login"
					
	findElement "Login" -attr "title" -tag "div"
					search for DIV element where the title attribute is "Login"
					
	findElement -tag "div"
					search for the first DIV element
#>
Set-Alias findElement Find-Element
function Find-Element
{
	param( [alias("t")]  [string]$text=$null,
		   [alias("i")]  [string]$id=$null,
		   [alias("n")]  [string]$name=$null,
		   [alias("a")]  [string]$attr=$null,
		   [alias("av")] [string]$attrValue=$null,
		   [alias("tg")] [string]$tag=$null,
		   [alias("rt")] [string]$rootTag=$null,
		   [alias("ra")] [string]$rootAttr=$null,
		   [alias("rav")][string]$rootAttrValue=$null
		   )
		   
	if( $rootTag -or $rootAttr)
	{
		trace "Searching for root element tag='$rootTag' attr $rootAttr='$rootAttrValue'."
		# search for the root element
		$root = findElement -av $rootAttrValue -a $rootAttr -tag $rootTag
		
		if( !$root)
		{
			# error
			write-comment "$($testScript):$testLine Failed to find root element." -type "error"
		}
		else
		{
			trace "findElement root element is $($root.tagName) id='$($root.id)'."
		}
	}
	
	if( !$root)
	{
		# search the whole page
		trace "searching entire DOM"
		$root = $ie.document
	}
	
	if( $attr -and ($attr -eq "class"))
	{
		# class doen't work - use className instead
		$attr = "className"
	}
	
	if( $id)
	{
		trace "findElement searching for id = '$id'."
		$el = $ie.document.getElementByID( $id)
	}
	elseif( $attr -and ($attr -eq "id"))
	{
		# shortcut for id searches - should be faster than the generic on below
		trace "findElement searching for attr $attr = '$attrValue'."
		$el = $ie.document.getElementByID( $attrValue)
	}
	elseif( $name)
	{
		trace "findElement searching for name = '$name'."
		$el = @($ie.document.getElementsByName( $name))[0]
	}
	elseif( $attr -and ($attr -eq "name"))
	{
		# shortcut for name searches - should be faster than the generic on below
		$el = @($ie.document.getElementsByName( $attrValue))[0]
	}
	elseif( $tag)
	{
		$items = $root.getElementsByTagName( $tag)
		$count = @($items).length
		
		if( $attr)
		{
			# qualify search by tag and attr
			trace "findElement searching $count $tag elements for attr $attr = '$attrValue'."
			foreach( $i in $items)
			{ 
				if( (getProperty $i $attr) -eq $attrValue) 
				{
					$el = $i
					break
				}
			}
		}
		elseif( $text)
		{
			# qualify search by tag and innerText
			trace "findElement searching $count $tag elements for text = '$text'."
			foreach( $i in $items)
			{ 
				if( $i.innerText -eq $text) 
				{
					$el = $i
					break
				}
			}
		}
		else
		{
			# qualify search by tag only
			trace "findElement $tag found $count elements."
			if( $items.length)
			{
				$el = $items[0];
			}
		}
	}
	else
	{
		if( $attr)
		{
			# qualify search by attribute only
			$items = $root.getElementsByTagName('*')
			$count = @($items).length
			
			trace "findElement searching $count elements for attr $attr = '$attrValue'."
			foreach( $i in $items)
			{ 
				$tmpTag = $i.tagName
				$tmpAttr = getProperty $i $attr
				#trace "testing $tmpTag for attr $attr does '$tmpAttr' = $attrValue"
				
				if( $tmpAttr -eq $attrValue) 
				{
					$el = $i
					break
				}
			}
		}
		elseif( $text)
		{
			trace "findElement by '$text' alone."
			# search by innerText, id, name
			$tags = @( "input", "button", "a", "select", "option", "div", "td", "li", "span")
			
			foreach( $tag in $tags)
			{
				$items = $root.getElementsByTagName( $tag)
				
				foreach( $i in $items)
				{ 
					$textVal = ''
					
					if( $tag -eq "input")
					{
						$textVal = $i.value
					}
					else
					{
						$textVal = $i.innerText
					}
					
					#trace "testing $tag element with text '$textVal' against '$text'"
					if( $textVal -like $text) 
					{
						$el = $i
						break
					}
				}
				
				if( $el)
				{
					break;
				}
			}
			
			# is it an id attribute
			if( !$el)
			{
				$el = $ie.document.getElementByID( $text)
			}
			
			# is it a name attribute
			if( !$el)
			{
				$el = @($ie.document.getElementsByName( $text))[0]
			}
		}
		else
		{
			# error
			Write-Comment "findElement bad combination of params." -type "error"
		}
	}
	
	# using 'return $el' causes powershell to 'unroll' enumerable objects
	# using ',$el' foils the unrolling business
	,$el
}

###########################################################
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
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line
	
	if( $tab)
	{
		Select-BrowserTab( $tab)
		return
	}

	$el = findElement $text $id $name $attr $attrValue $tag $rootTag $rootAttr $rootAttrValue
	
	if( !$el)
	{
		trace "no matching element found"
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
				#trace "item '$($op.innerText)'"
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
			catch {}
			
			try 
			{
				if( !$double)
				{
					$el.click()
				}
				else
				{
					[void]$el.fireEvent( "ondblclick")
				}
			}
			catch {}
			
			try 
			{
				[void]$el.fireEvent( "onmouseout")
			}
			catch {}
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
Set-Alias fill Set-ElementValue
function Set-ElementValue( $id, $val)
{
	$result = $true
	$script:testScript = $MyInvocation.ScriptName.split("\")[-1].split(".")[0]
	$script:testLine = $MyInvocation.ScriptLineNumber
	$script:cmdLine = $MyInvocation.Line

	$el = findElement -id $id
	
	if( !$el)
	{
		$el = findElement -name $id
	}
	
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
		trace "setting $($el.tagName) id='$($el.id)' name='$($el.name)' to '$val'"
		
		if( $el.tagName -eq "select")
		{
			$idx = -1
			foreach( $op in @($el.childNodes))
			{
				#trace "item '$($op.innerText)'"
				if( $op.innerText -eq $val)
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
				trace "'$val' is not a valid selection for select id='$($el.id)'"
				$result = $false
			}
		}
		else
		{
			$el.value = $val
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
		   [alias("v")]  [switch]$visible
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
		
	if( $notFound)
	{
		$result = !$result
	}
	
	Write-TestResult $result 
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
			$html += "  <a href='./screenshots/$ssFileName'><img src='./screenshots/$ssFileName' title='$ssFileName'/></a>"
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
		$html +=     "<a href='./screenshots/$ssFileName'><img src='./screenshots/$ssFileName' title='$ssFileName'/></a>"
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
	$script:sectionOutput | Out-File $logFileName -Append -Force
	$script:sectionOutput = $null
}

###########################################################
function Show-HtmlLog
{
	if( $showHtmlLog)
	{
		$ie.navigate2($logFileName, 2048, "Test Summary")
	}
}

###########################################################
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
	
	Get-ScreenShot -ie $ie -file "$logDir\screenshots\$ssFileName" 
	
	if( !$noLog)
	{
		Add-HtmlLogEntry -t "screenshot" -v $description
	}
}

###########################################################
. $psScriptRoot\Get-Screenshot.ps1

###########################################################
. $psScriptRoot\Get-IETab.ps1

###########################################################
Export-ModuleMember -Alias * -Function *