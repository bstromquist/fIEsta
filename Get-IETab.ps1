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

#Requires -Version 2
Function Select-IETab 
{
	Param (
       [object]$ie,
       [Parameter(Mandatory = $False)][string]$name
	)
	
# C# code
$code = @'
using System;
using System.Collections.Generic;
using Accessibility;
using System.Runtime.InteropServices;
using System.Diagnostics;
 
namespace IEAccessibleTabs
{
    public class IEAccessible
    {
        private enum OBJID : uint
        {
            OBJID_WINDOW = 0x00000000,
        }
		
        private const int IE_ACTIVE_TAB = 2097154;
        private const int CHILDID_SELF = 0;
        private const int SELFLAG_TAKEFOCUS = 1;
        private IAccessible accessible;
        private IEAccessible[] Children
        {
            get
            {
                int num = 0;
                object[] res = GetAccessibleChildren(accessible, out num);
                if (res == null)
                    return new IEAccessible[0];
 
                List<IEAccessible> list = new List<IEAccessible>(res.Length);
                foreach (object obj in res)
                {
                    IAccessible acc = obj as IAccessible;
                    if (acc != null)
                        list.Add(new IEAccessible(acc));
                }
                return list.ToArray();
            }
        }
		
        private string Name
        {
            get
            {
                string ret = accessible.get_accName(CHILDID_SELF);
                return ret;
            }
        }
		
        private int ChildCount
        {
            get
            {
                int ret = accessible.accChildCount;
                return ret;
            }
        }
 
        public bool Activate(IntPtr ieHandle, string tabCaptionToActivate)
        {
            AccessibleObjectFromWindow( GetDirectUIHWND(ieHandle), OBJID.OBJID_WINDOW, ref accessible);
			
            if (accessible == null)
                throw new Exception();
 
            IEAccessible ieDirectUIHWND = new IEAccessible( accessible);

            foreach (IEAccessible accessor in ieDirectUIHWND.Children)
            {
                foreach (IEAccessible child in accessor.Children)
                {
                    foreach (IEAccessible tab in child.Children)
                    {
                        if (tab.Name == tabCaptionToActivate)
                        {
                            tab.Activate();
							//tab.Select();
                            return true;
                        }
                    }
                }
            }
			
			return false;
        }
 
        private IntPtr GetDirectUIHWND(IntPtr ieFrame)
        {			
			// For IE 8:
				IntPtr directUI = FindWindowEx(ieFrame, IntPtr.Zero, "CommandBarClass", null);
				directUI = FindWindowEx(directUI, IntPtr.Zero, "ReBarWindow32", null);
				directUI = FindWindowEx(directUI, IntPtr.Zero, "TabBandClass", null);
				directUI = FindWindowEx(directUI, IntPtr.Zero, "DirectUIHWND", null);

			if (directUI == IntPtr.Zero)
			{
				// For IE 9:
				directUI = FindWindowEx(ieFrame, IntPtr.Zero, "WorkerW", "Navigation Bar");
				directUI = FindWindowEx(directUI, IntPtr.Zero, "ReBarWindow32", null);
				directUI = FindWindowEx(directUI, IntPtr.Zero, "TabBandClass", null);
				directUI = FindWindowEx(directUI, IntPtr.Zero, "DirectUIHWND", null);
			}
            return directUI;
 
        }
		
        public IEAccessible()
        {
        }
		
        private IEAccessible(IAccessible acc)
        {
            if (acc == null)
                throw new Exception();
 
            accessible = acc;
        }
		
        private void Activate()
        {
            accessible.accDoDefaultAction( CHILDID_SELF);
        }
		
        private void Select()
        {
            accessible.accSelect( SELFLAG_TAKEFOCUS, CHILDID_SELF);
        }
		
        private static object[] GetAccessibleChildren( IAccessible ao, out int childs)
        {
            childs = 0;
            object[] ret = null;
            int count = ao.accChildCount;
            if (count > 0)
            {
                ret = new object[count];
                AccessibleChildren(ao, 0, count, ret, out childs);
            }
            return ret;
        }
 
        #region Interop
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass,
        string lpszWindow);
        private static int AccessibleObjectFromWindow(IntPtr hwnd, OBJID idObject, ref IAccessible acc)
        {
            Guid guid = new Guid("{618736e0-3c3d-11cf-810c-00aa00389b71}"); // IAccessible
 
            object obj = null;
            int num = AccessibleObjectFromWindow(hwnd, (uint)idObject, ref guid, ref obj);
            acc = (IAccessible)obj;
            return num;
        }
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);
        [DllImport("oleacc.dll")]
        private static extern int AccessibleChildren(IAccessible paccContainer, int iChildStart, int cChildren, [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] rgvarChildren, out int pcObtained);
        #endregion
    }
}
'@

	#User Add-Type to import the code
	add-type $code -ReferencedAssemblies 'Accessibility'
	
	#Create the object for the Function
	$obj = New-Object IEAccessibleTabs.IEAccessible
	if( $name)
	{
		$obj.Activate( $ie.hwnd, $name)
	}
}