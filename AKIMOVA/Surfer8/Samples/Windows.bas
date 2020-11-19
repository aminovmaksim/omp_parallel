'Windows.bas illustrates the Windows Collection and Window Object.
Sub Main
	Debug.Print "----- Windows.bas - ";Time;" -----"

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
 	Debug.Print "Surfer ";surf.Version

	Set plotdoc1 = surf.Documents(1)
  Set plotwin1 = surf.Windows(1)
	path1 = surf.Path+"\samples\"
	file1 = path1+"demogrid.dat"
	'Set wksdoc1 = surf.Documents.Open(file1)

	Set shapes1 = plotdoc1.Shapes

	'=============================================================
	'The Windows Collection consists of all the Windows in Surfer.
	'=============================================================
	Set wins = surf.Windows
	Debug.Print "Windows Collection Application: ";wins.Application
	Debug.Print "Windows Collection Count: ";wins.Count
	Debug.Print "Windows Collection Parent: ";wins.Parent

	'--------------------------------------------------------------------
	'Arrange method tiles, cascades, and arranges minimized window icons.
	'--------------------------------------------------------------------
	wins.Arrange(srfCascade)

	'==============================================================
	'The Window Object is the base object for the PlotWindow,
	' WorksheetWindow, and GridWindow.
	'==============================================================
	'------------------------------------------------------------------------
	'wins(i) is the short way of writing wins.Item(i) and wins.Item(Index:=i)
	'------------------------------------------------------------------------
	'Active returns a boolean describing the active state of the window.
	For i = 1 To wins.Count
		Debug.Print "  ";wins(i).Caption;" Active? ";wins(i).Active
	Next i
	Debug.Print "Window 1 Application: ";wins(1).Application

	'Caption is the default Window Object property.
	Debug.Print "Window 1 Caption: ";wins(1).Caption
	Debug.Print "Window 1 Caption: ";wins(1)

	'Document returns the document object associated with the window.
	Debug.Print "Window 1 Document: ";wins(1).Document.Name

	'Parent returns the window's parent object.
	Debug.Print "Window 1 Parent: ";wins(1).Parent

	'Index returns the 1-based index in the Windows collection.
	surf.Documents.Add(srfDocWks) 'Add a worksheet window.
	surf.Documents.Open(path1+"demogrid.grd") 'Open a grid editor window.
	For Each wnd In wins
		'See end of script for wintypename() function.
		Debug.Print "Window #";wnd.Index;" ";wnd.Caption;
		Debug.Print " Type:"; wintypename(wnd.Type);
		Debug.Print " State:";winstatename(wnd.WindowState)
		Debug.Print " left:";wnd.Left;" top:";wnd.Top;
		Debug.Print " height:";wnd.Height;" width:";wnd.Width
	Next

	'Move and resize window 1.
	wins(1).Height = wins(1).Height - 10 'decrease height by 10 units.
	wins(1).Width = wins(1).Width - 10
	wins(1).Top = wins(1).Top + 5
	wins(1).Left = wins(1).Left + 5

	'====================
	'WindowObject Methods
	'====================
	'Activate Window 1.
	wins(1).Activate

	'Close the grid editor window containing "demogrid.grd"
	AppActivate "Surfer"
	wins("demogrid.grd").Activate
	Wait 1
	wins("demogrid.grd").Close

End Sub
Function wintypename(x)
	'Array is 0-based, but enums are 1-based, so pad 0-index.
	wtn = Array ("0", "srfWinPlot", "srfWinWks", "srfWinGrid")
	wintypename = wtn(x)
End Function
Function winstatename(x)
	wsn = Array("0","maximized", "minimized",	"normal")
 	winstatename = wsn(x)
End Function
