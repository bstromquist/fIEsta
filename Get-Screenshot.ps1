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
    Used to take a screenshot of an IE tab window and save as a png/bmp/jpeg file. 
.PARAMETER screen
    Screenshot of the entire screen
.PARAMETER ie
    Internet Explorer object
.PARAMETER file
    Name of the file to save as. Default is image.bmp
.PARAMETER imagetype
    Type of image being saved. Can use JPEG,BMP,PNG. Default is bitmap(bmp)    
#>  
#Requires -Version 2
Function Get-ScreenShot 
{
	Param (
       [Parameter(Mandatory = $True) ][object]$ie,
       [Parameter(Mandatory = $False)][string]$file="$pwd\image.bmp", 
       [Parameter(Mandatory = $False)][string][ValidateSet("bmp","jpeg","png")]$imagetype = "png"           
    )

# C# code
$code = @'
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
namespace IETabScreenshot
{
  public class ScreenCapture
  {
	public IntPtr FindFirstChildWindow( IntPtr parent, string windowClass, string windowTitle){
		return Win32.FindWindowEx( parent, IntPtr.Zero, windowClass, windowTitle);
	}
	
	public IntPtr FindNextChildWindow( IntPtr parent, IntPtr child, string windowClass, string windowTitle){
		return Win32.FindWindowEx( parent, child, windowClass, windowTitle);
	}
	
    public int CaptureWindowToFile(IntPtr handle, string filename, ImageFormat format)
    {
		IntPtr hdcSrc = Win32.GetWindowDC( handle);
		Win32.RECT windowRect = new Win32.RECT();
		Win32.GetWindowRect(handle,ref windowRect);
		int width = windowRect.right - windowRect.left;
		int height = windowRect.bottom - windowRect.top;

		// create a device context we can copy to
		IntPtr hdcDest = Win32.CreateCompatibleDC( hdcSrc);
		// create a bitmap we can copy it to,
		// using GetDeviceCaps to get the width/height
		IntPtr hBitmap = Win32.CreateCompatibleBitmap(hdcSrc, width, height);
		// select the bitmap object
		IntPtr hOld = Win32.SelectObject(hdcDest, hBitmap);
		// bitblt over
		//Gdi32.BitBlt(hdcDest, 0, 0, 135, 60, hdcSrc, 10, 38, Gdi32.TernaryRasterOperations.SRCCOPY);
		Win32.PrintWindow( handle, hdcDest, 0x1);
		// restore selection
		Win32.SelectObject(hdcDest, hOld);
		// clean up
		Win32.DeleteDC(hdcDest);
		Win32.ReleaseDC(handle, hdcSrc);
		// get a .NET image object for it
		Bitmap bmp = Image.FromHbitmap(hBitmap);
		bmp.Save(filename,format);
		
		// free up the Bitmap object
		Win32.DeleteObject(hBitmap);

	    return 0;
    }
	
    /// <summary>
    /// Helper class containing User32 API functions
    /// </summary>
    private class Win32
    {
	[StructLayout(LayoutKind.Sequential)]
	public struct RECT
	{
		public int left;
		public int top;
		public int right;
		public int bottom;
	}
	[DllImport("user32.dll")]
	public static extern IntPtr GetWindowDC(IntPtr hWnd);
	[DllImport("user32.dll")]
	public static extern IntPtr ReleaseDC(IntPtr hWnd,IntPtr hDC);
	[DllImport("user32.dll")]
	public static extern IntPtr GetWindowRect(IntPtr hWnd,ref RECT rect);
	[DllImport("user32.dll")]
	public static extern IntPtr PrintWindow(IntPtr hWnd, IntPtr dc, uint flags);     
	[DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
	public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string lclassName, string windowTitle);
	[DllImport("gdi32.dll")]
	public static extern bool DeleteObject(IntPtr hObject);
	[DllImport("gdi32.dll", SetLastError=true)]
	public static extern IntPtr CreateCompatibleDC(IntPtr hdc);
	[DllImport("gdi32.dll")]
	public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);
	[DllImport("gdi32.dll", ExactSpelling=true, PreserveSig=true, SetLastError=true)]
	public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);
	[DllImport("gdi32.dll")]
	public static extern bool DeleteDC(IntPtr hdc);
    }
  }
}
'@
	#User Add-Type to import the code
	add-type $code -ReferencedAssemblies 'System.Windows.Forms','System.Drawing'
	#Create the object for the Function
	$capture = New-Object IETabScreenshot.ScreenCapture

	if( $ie)
	{
		$windowHandle = $ie.hWnd
		$tabTitle = "$($ie.document.title) - Windows Internet Explorer"
		
		#Write-Host "Taking screenshot of IE tab with title '$tabTitle'"
		#Write-host "main window handle = $($ie.hwnd)"
		$hwndFrame = $capture.FindFirstChildWindow( $ie.hWnd, "Frame Tab", "")
		
		if( $hwndFrame -eq 0)
		{
			# IE7 doesn't have the frame tab window
			$hwndFrame = $ie.hWnd
		}
		
		for( $i = 0; $i -lt 20; $i++) # search the first 20 tabs
		{
			#Write-host "tab frame handle = $hwndFrame"
			$hwnd = $capture.FindFirstChildWindow( $hwndFrame, "TabWindowClass", $tabTitle)
			
			if( $hwnd -ne 0)
			{
				break
			}
			
			$hwndFrame = $capture.FindNextChildWindow( $ie.hWnd, $hwndFrame, "Frame Tab", "")
		}

		if( $hwnd.value -ne 0)
		{
			$hwnd = $capture.FindFirstChildWindow( $hwnd, "Shell DocObject View", "")
		}
		
		if( $hwnd.value -ne 0)
		{
			$hwnd = $capture.FindFirstChildWindow( $hwnd, "Internet Explorer_Server", "")
		}
		
		if( $hwnd.value -ne 0)
		{
			#Save to a file
			If ($file) 
			{
				If ($file -eq "") 
				{
					$file = "$pwd\image.bmp"
				}
				
				[void]$capture.CaptureWindowToFile( $hWnd, $file,$imagetype)
			}
		}
		else
		{
			write-host "window not found"
		}
	}
}