' ----------------------------------------------------------------------------
' Script5.vbs
'
' The is a simple VBScript example to create a shaded relief map.
'
' To run this from the command line:
' cscript Script5.vbs
'
' This requires Windows Scripting Host, available free as part of
' Microsoft Internet Explorer 5.0, or from Microsoft's web site at:
' http://msdn.microsoft.com/scripting
'
' Note: As of 7-30-99, VBScript does not support accessing SAFEARRAY's that
'       contain non-VARIANT elements.  This makes it unsuitable for several 
'       of the Surfer Automation methods and properties that operate on a
'       SAFEARRAY containing doubles, longs, etc.  Microsoft has documented
'       this at:
'       http://support.microsoft.com/support/kb/articles/q165/9/67.asp?FR=0
'
'       Here is an extract from the above article:
'       The VBSCRIPT active scripting engine supplied by Microsoft only 
'       supports the indexing of SAFEARRAYs of VARIANTs. While VBSCRIPT is 
'       capable of accepting arrays of non-variant type for the purposes of 
'       boundary checking and passing it to other automation objects, the 
'       engine does not allow manipulation of the array contents at this time. 
'
' ----------------------------------------------------------------------------

' Create an instance of the Surfer Application object and assign it to the
' variable named "SurferObj"
Set SurferObj = WScript.CreateObject("Surfer.Application")

' Makes the Surfer window visible
SurferObj.Visible = True

' Creates a plot document in Surfer and assigns it to the variable named "DocObj"
Set DocObj = SurferObj.Documents.Add

' Assigns the location of the grid file to the variable named "Path"
GridFile = SurferObj.Path + "\samples\Helens2.grd"
	
' Create the shaded relief map from Helens2.grd
Set MapFrameObj = DocObj.Shapes.AddReliefMap(GridFile)
