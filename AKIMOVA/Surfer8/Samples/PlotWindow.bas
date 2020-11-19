'PlotWindow.bas illustrates the PlotWindow Object.
' The PlotWindow is derived from the Window Object.
' Refer to the Windows.bas file for the methods and properties that
'  are shared by the PlotWindow and WorksheetWindow Objects.
Sub Main
	Debug.Print "----- PlotWindow.bas - ";Time;" -----"

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

	Set shapes1 = plotdoc1.Shapes

	'=====================
	'PlotWindow Properties
	'=====================
	'--------------------------------------------------------
	'AutoRedraw returns a Boolean indicating whether
	' Automatic Redraw is enabled.
	' It can also enable or disable Automatic Redraw.
	'--------------------------------------------------------
	Debug.Print "AutoRedraw enabled? ";plotwin1.AutoRedraw;
	plotwin1.AutoRedraw = True
	Debug.Print ". AutoRedraw set to ";plotwin1.AutoRedraw

	'--------------------------------------------------------
	'The ShowRulers Property returns a Boolean to indicate
	' whether the rulers are visible.  It can be used to
	' enable or disable the ruler visibility.
	'--------------------------------------------------------
	Debug.Print "Rulers visible? ";plotwin1.ShowRulers;
	plotwin1.ShowRulers = True
	Debug.Print ". Rulers now set to ";plotwin1.ShowRulers

	'--------------------------------------------------------
	'HorizontalRuler Property returns a Ruler Object,
	' which can be used to set GridDivisions, RulerDivisions,
	' ShowPosition, and SnapToRuler.  Not saved when Surfer
	' is closed.
	'--------------------------------------------------------
	Debug.Print "Horizontal Ruler Divisions per page unit = "; _
		plotwin1.HorizontalRuler.RulerDivisions;
	plotwin1.HorizontalRuler.RulerDivisions = 8
	Debug.Print " New div/PU = "; plotwin1.HorizontalRuler.RulerDivisions

	'--------------------------------------------------------
	'VerticalRuler Property returns a Ruler Object,
	' which can be used to set GridDivisions, RulerDivisions,
	' ShowPosition, and SnapToRuler.  Not saved when Surfer
	' is closed.
	'--------------------------------------------------------
	Debug.Print "Vertical Ruler Divisions per page unit = "; _
		plotwin1.VerticalRuler.RulerDivisions;
	plotwin1.VerticalRuler.RulerDivisions = 8
	Debug.Print ". New div/PU = "; plotwin1.VerticalRuler.RulerDivisions

	'--------------------------------------------------------------
	'The ShowGrid Property returns a Boolean to indicate if the
	' drawing grid is visible.  It can be used to enable or disable
	' the drawing grid.
	'--------------------------------------------------------------
	Debug.Print "Drawing Grid visible? ";plotwin1.ShowGrid;
	plotwin1.ShowGrid = True
	Debug.Print ". Drawing grid is now set to ";plotwin1.ShowGrid

	'--------------------------------------------------------------
	'The ShowMargins Property returns a Boolean to indicate if the
	' drawing Margins are visible.  It can be used to enable or disable
	' the drawing Margins.
	'--------------------------------------------------------------
	Debug.Print "Margins visible? ";plotwin1.ShowMargins;
	AppActivate "Surfer"
	plotwin1.ShowMargins = False
	shapes1.AddText(1,1,"Margins OFF").Font.Size=40
	Wait 2
	shapes1("Text").Delete
	plotwin1.ShowMargins = True
	shapes1.AddText(1,1,"Margins ON").Font.Size=40
	Wait 2
	shapes1("Text").Delete
	AppActivate "PlotWindow"
	Debug.Print ". Margins are now set to ";plotwin1.ShowMargins

	'--------------------------------------------------------------
	'The ShowPage Property returns a Boolean to indicate if the
	' Page outline is visible.  It can be used to enable or
	' disable the Page outline
	'--------------------------------------------------------------
	Debug.Print "Page Outline visible? ";plotwin1.ShowPage;
	AppActivate "Surfer"
	plotwin1.ShowPage = False
	shapes1.AddText(1,1,"Page Outline OFF").Font.Size=40
	Wait 2
	shapes1("Text").Delete
	plotwin1.ShowPage = True
	shapes1.AddText(1,1,"Page Outline ON").Font.Size=40
	Wait 2
	shapes1("Text").Delete
	AppActivate "PlotWindow"
	Debug.Print ". Page Outline now set to ";plotwin1.ShowPage

	'==================
	'PlotWindow Methods
	'==================
	'-----------------------------------------------------
	'The Redraw method redraws the contents of the window.
	'-----------------------------------------------------
	Debug.Print "Redraw"
	AppActivate "Surfer"
	'Using plotdoc1 here to emphasize difference between
	' plot document and plot window.
	plotdoc1.Shapes.AddContourMap(path1+"helens2.grd")
	plotdoc1.Shapes.AddText(1,1,"Redrawing...").Font.Size=40
	Wait 2
	plotwin1.Redraw
	For Each shp In shapes1
		shp.Delete
	Next shp
	plotwin1.Zoom(srfZoomPage)

	'---------------------------------------------------
	'The Zoom method has the following parameters:
	' srfZoomFitToWindow, srfZoomPage, srfZoomActualSize,
	' srfZoomSelected, srfZoomFullScreen
	'---------------------------------------------------
	Debug.Print "Zoom"
	plotdoc1.Shapes.AddContourMap(path1+"demogrid.grd")
	plotdoc1.Shapes.AddText(1,1,"Zoom Fit To Window").Font.Size = 40
	Wait 2
	plotwin1.Zoom(srfZoomFitToWindow)
	Wait 2
	For Each shp In shapes1
		shp.Delete
	Next shp
	plotwin1.Zoom(srfZoomPage)

	'---------------------------------------------------------
	'The ZoomPoint method zooms in or out at the specified
	' factor and point XY location in page units.
	'---------------------------------------------------------
	Debug.Print "ZoomPoint (1,1) 300%"
	plotdoc1.Shapes.AddContourMap(path1+"demogrid.grd")
	plotdoc1.Shapes.AddText(1,1,"Zoom 300% at (1,1)").Font.Size = 40
	Wait 2
	plotwin1.ZoomPoint(1,1,Scale:=3)
	Wait 2
	For Each shp In shapes1
		shp.Delete
	Next shp
	plotwin1.Zoom(srfZoomPage)

	'----------------------------------------------
	'ZoomRectangle zooms in or out so the specified
	' rectangle fills the screen.
	'----------------------------------------------
	Debug.Print "ZoomRectangle (5.5,4,6.5,3)"
	plotdoc1.Shapes.AddContourMap(path1+"demogrid.grd")
	plotdoc1.Shapes.AddText(1,1,"ZoomRectangle (5.5,4,6.5,3)").Font.Size = 40
	Wait 2
	plotwin1.ZoomRectangle (5.5,4,6.5,3)
	Wait 2
	For Each shp In shapes1
		shp.Delete
	Next shp
	plotwin1.Zoom(srfZoomPage)

End Sub
