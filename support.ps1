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

###########################################################
<#  
.SYNOPSIS  

.DESCRIPTION
    
.PARAMETER x
    ...
.EXAMPLE 
    ...  
#>  
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
<#  
.SYNOPSIS  

.DESCRIPTION
    
.PARAMETER x
    ...
.EXAMPLE 
    ...  
#>  
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
	
	trace "findElement text = '$text' id = '$id' name = '$name' tag = '$tag' attr = '$attr' attrValue = '$attrValue'"
	
	$el = $null
		
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
		trace "findElement searching by tag name '$tag'"
		$items = $root.getElementsByTagName( $tag)
		$count = @($items).length
		
		if( $attr)
		{
			# qualify search by tag and attr
			trace "findElement searching $count $tag elements for attr $attr = '$attrValue'."
			foreach( $i in $items)
			{ 
				if( (getProperty $i $attr) -like $attrValue) 
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
				if( $i.innerText -like $text) 
				{
					$el = $i
					break
				}
			}
		}
		else
		{
			# qualify search by tag only
			trace "findElement found $count $tag elements."
			if( $count)
			{
				$el = @($items)[0];
				trace "returning $($el.tagName)"
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
				
				if( $tmpAttr -like $attrValue) 
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
			
			# is it an id attribute
			$el = $ie.document.getElementByID( $text)
			
			# is it a name attribute
			if( !$el)
			{
				$el = @($ie.document.getElementsByName( $text))[0]
			}

			# is it innerText
			if( !$el)
			{
				# search in reverse to get the inner-most item that matches the text
				$items = $root.getElementsByTagName( "*")
				$count = @($items).length
				
				for( $index = $count-1; $index -ne 0; $index--)
				{ 
					$i = @($items)[$index]
					$textVal = ''
					
					if( $i.tagName -eq "input")
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
