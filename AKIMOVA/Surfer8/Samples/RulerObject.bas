'RulerObject.bas illustrates the properties of the Ruler Object.
Sub Main
	Debug.Print "----- RulerObject.bas - ";Time;" -----"

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
	Set shapes1 = plotdoc1.Shapes
	path1 = surf.Path+"\samples\"

	'=======================
	'RulerObject Properties
	'=======================
	Set hruler = plotwin1.HorizontalRuler

	'------------------------------------------------------------
	'The Application returns the application object, i.e. Surfer.
	'------------------------------------------------------------
	Debug.Print "Horizontal Ruler Application: ";hruler.Application

	'-----------------------------------------------------------
	'The Parent Property returns the parent object, i.e. Surfer.
	'-----------------------------------------------------------
	Debug.Print "Ruler Parent: ";hruler.Parent

	'------------------------------------------------------------
	'The GridDivisions Property returns or sets the the number of
	' grid divisions per page unit. (integer).
	'------------------------------------------------------------
	Debug.Print "The horizontal ruler has";hruler.GridDivisions; _
		" grid division per page unit."
	hruler.GridDivisions = 2.
	Debug.Print "The horizontal ruler has";hruler.GridDivisions; _
		" grid division per page unit."

		'------------------------------------------------------------
	'The RulerDivisions Property returns or sets the number of
	' ruler divsisions per page unit (integer).
	'------------------------------------------------------------
	Debug.Print "The horizontal ruler has";hruler.RulerDivisions; _
		" ruler divisions per page unit."
	hruler.RulerDivisions = 10
	Debug.Print "Now the horizontal ruler has";hruler.RulerDivisions; _
		" ruler divisions per page unit."

	'------------------------------------------------------------
	'The ShowPosition Property returns or sets the "show cursor
	' postition" state on the ruler (Boolean).
	'------------------------------------------------------------
	Debug.Print "Show cursor position on horizontal ruler? "; _
		hruler.ShowPosition
	hruler.ShowPosition = False
	Debug.Print "Show cursor position on horizontal ruler? "; _
		hruler.ShowPosition
	hruler.ShowPosition = True
	Debug.Print "Show cursor position on horizontal ruler? "; _
		hruler.ShowPosition

	'------------------------------------------------------------
	'The SnapToRuler Property returns or sets the
	' "snap to ruler" state. (Boolean).
	'------------------------------------------------------------
	Debug.Print "The horizontal ruler SnapToRuler state is "; _
		hruler.SnapToRuler
	hruler.SnapToRuler = True
	Debug.Print "Now the horizontal ruler SnapToRuler state is "; _
		hruler.SnapToRuler

End Sub
