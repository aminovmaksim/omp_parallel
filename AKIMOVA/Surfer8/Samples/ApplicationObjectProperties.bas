'ApplicationObjectProperties.bas demonstrates the properties
' of the Surfer Application Object.
' See ApplicationObjectMethods.bas for the methods of the
' Surfer Application Object.
' TB - 17 Oct 01.
Sub Main
	Debug.Print "----- ";Time;" -----"

	'Get existing Surfer instance, or create a new one If none exists.
	On Error Resume Next 'Turn off error reporting.
	Set surf = GetObject(,"Surfer.Application")
	If Err.Number<>0 Then
		Set surf = CreateObject("Surfer.Application")
		surf.Documents.Add(srfDocPlot)	
	End If
	On Error GoTo 0 'Turn on error reporting.

	surf.Visible = True
	surf.WindowState = srfWindowStateNormal
	surf.Width = 600
	surf.Height = 400
	surf.Windows(1).Zoom(srfZoomPage)

 	Debug.Print "-----------------------------------------------------"
	Debug.Print "Surfer ";surf.Version;" Application Object Properties"
 	Debug.Print "-----------------------------------------------------"

	Debug.Print "Application= ";surf.Application

 	Debug.Print "ActiveDocument= ";surf.ActiveDocument; '"Plot1.srf" (read-only)
 	'Set the active document to "Plot1". (Can omit ".srf").
 	surf.Documents("Plot1").Activate
 	'Set the active document to the first doc in the Documents Collection.
 	surf.Documents(1).Activate
 	
	'ActiveWindow is "Plot1" or "Plot1:1" (read-only)
 	Debug.Print ". ActiveWindow= ";surf.ActiveWindow
 	'Set the active window to the first window in the Windows Collection.
	surf.Windows(1).Activate
	
	Debug.Print "BackupFiles= ";surf.BackupFiles;
	'Change to File | Preferences is saved when Surfer closes.
	surf.BackupFiles = True 
	Debug.Print ". New BackupFiles= ";surf.BackupFiles

	Debug.Print "Caption= ";surf.Caption;
	surf.Caption = "Surfer "+surf.Version
	Debug.Print ". New Caption= ";surf.Caption

	Debug.Print "DefaultFilePath=";surf.DefaultFilePath;
	'New default path is used during current session of Surfer, 
	' but does not change default path in File | Preferences.
	surf.DefaultFilePath = "D:\INCOMING\"
	Debug.Print ". New DefaultFilePath=";surf.DefaultFilePath

	Debug.Print "FullName = "; surf.FullName
	Debug.Print "Name = ";surf.Name
	Debug.Print "Path = ";surf.Path

	AppActivate "Surfer"
	Debug.Print "Surfer Application Window Height = ";surf.Height;
	surf.Height = surf.Height + 10
	Debug.Print ". New height = ";surf.Height

	Debug.Print "Surfer Application Window Width = "; surf.Width;
	surf.Width = surf.Width + 10  '0 = top of screen.
	Debug.Print ". New width = ";surf.Width

	Debug.Print "Surfer Application Window Left = ";surf.Left;
	surf.Left = surf.Left + 10 '0 = far left edge of screen.
	Debug.Print ". New left = ";surf.Left

	Debug.Print "Surfer Application Window Top = ";surf.Top;
	surf.Top = surf.Top + 10
	Debug.Print ". New Top = ";surf.Top
	AppActivate "ApplicationObjectProperties"
	Debug.Print "Page Units ="; surf.PageUnits;". PageUnits Name = ";
	If surf.PageUnits = 1 Then Debug.Print "srfUnitsInch"
	If surf.PageUnits = 2 Then Debug.Print "srfUnitsCentimeter"

	'The application is the topmost level object, so it is it's own parent.
	Debug.Print "Surfer Application Parent = ";surf.Parent

	Debug.Print "ScreenUpdating = ";surf.ScreenUpdating;
	surf.ScreenUpdating = True
	Debug.Print ". New ScreenUpdating = ";surf.ScreenUpdating

	Debug.Print "ShowStatusBar = ";surf.ShowStatusBar;
	surf.ShowStatusBar = True
	Debug.Print ". New ShowStatusBar = ";surf.ShowStatusBar

	Debug.Print "ShowToolbars =";surf.ShowToolbars;
	Debug.Print ", sum of srfTBMain =";srfTBMain;
	Debug.Print " + srfTBDraw =";srfTBDraw;
	Debug.Print " + srfTBMap =";srfTBMap
	surf.ShowToolbars = srfTBMain+srfTBDraw+srfTBMap 'Turn them all on.

	Debug.Print "Application Visible = ";surf.Visible;
	surf.Visible = True
	Debug.Print ". Now Application Visible = ";surf.Visible

	Debug.Print "There are";surf.Windows.Count;" windows in the Windows Collection."

	Debug.Print "WindowState =";surf.WindowState
	Debug.Print " srfWindowStateMaximized =";srfWindowStateMaximized
	Debug.Print " srfWindowStateMinimized =";srfWindowStateMinimized
	Debug.Print " srfWindowStateNormal =";srfWindowStateNormal
	Debug.Print  'a blank line

End Sub
