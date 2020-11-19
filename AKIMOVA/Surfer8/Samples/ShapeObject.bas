'ShapesObject.bas illustrates the properties of the
' following objects:
'  Rectangle (includes rounded rectangles, squares)
'  Ellipse (includes circles)
'  Symbol
'  Text
'  Polyline
'  Polygon
'  Composite
'See VariogramObject.bas, specific map object BAS files
' for examples of variogram and mapframe based objects.
'------------------------------------------------------------
Sub Main
	Debug.Print "----- ShapesObject.bas - ";Time;" -----"

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

	Debug.Print "Surfer ";surf.Version
	surf.Caption = "Surfer "+surf.Version
	AppActivate "Surfer "+surf.Version

	Set plotdoc1 = surf.Documents(1)
	If plotdoc1.Type <> srfDocPlot Then plotdoc1 = surf.Documents.Add
 	Set plotwin1 = surf.Windows(1)

	With plotdoc1.PageSetup
		.Orientation = srfLandscape
		.Height = 8.5
		.Width = 11
	End With
	plotwin1.Zoom(srfZoomPage)
	path1 = surf.Path+"\samples\"

	Set shapes1 = plotdoc1.Shapes
	shapes1.SelectAll
	plotdoc1.Selection.Delete

	'===============================================================
	'Shape Object Properties
	' The Shape Object is the base object for all graphical objects
	' within the plot document.  All objects derived from the
	' Shape Object share the following properties and methods.
	' Refer to the Rectangle Object below for examples.
	'===============================================================

	'===========================
	'Rectangle Object Properties
	'===========================
	'The Rectangle object is created with the Shapes.AddRectangle method.
	Debug.Print "AddRectangle"
	shapes1.AddText(1,1,"AddRectangle").Font.Size = 40
	Set rectangle1 = shapes1.AddRectangle( Left:=1,Top:=6, _
		Right:=4, Bottom:=4)

	AppActivate "ShapesObject"
	'------------------------------
	'Shared Shape Object Properties
	'------------------------------
	Debug.Print "Shape Object Properties"
	With rectangle1
		Debug.Print " Rectangle height:"; .Height; ", width:"; .Width; _
		", left:"; .Left; ", top:";.Top
		Debug.Print " Application: "; .Application
		Debug.Print " Name: "; .Name
		Debug.Print " Parent: "; .Parent
		AppActivate "Surfer "+surf.Version
	End With

	Debug.Print " Current rotation angle:"; rectangle1.Rotation
	rectangle1.Rotation = 30
	Debug.Print " New rotation angle:"; rectangle1.Rotation

	Debug.Print "Rectangle selected: "; rectangle1.Selected
	Wait 1
	rectangle1.Selected = True 'Property
	rectangle1.Select          'Method. Same effect as Selected = True"
	Debug.Print "Rectangle selected? "; rectangle1.Selected

	Debug.Print "Rectangle visible? ";rectangle1.Visible
	Wait 1
	rectangle1.Visible = False
	Debug.Print "Rectangle visible? ";rectangle1.Visible
	Wait 1
	rectangle1.Visible = True

	'----------------------------
	'Shared Shape Object methods.
	'----------------------------
	Set rectangle2 = shapes1.AddRectangle(Left:=3, Top:=5, _
		Right:=5, Bottom:=3)
	rectangle1.Fill.ForeColor = srfColorRed
	rectangle1.Fill.Pattern = "Solid"
	rectangle2.Fill.ForeColor = srfColorBlue
	rectangle2.Fill.Pattern = "Solid"

	rectangle2.SetZOrder(srfZOToBack)
	rectangle2.Select
	Debug.Print "Rectangle2 selected? ";rectangle2.Selected
	Wait 1
	rectangle2.Select 'or rectangle2.Selected=True
	Debug.Print "Rectangle2 selected? ";rectangle2.Selected
	Wait 1
	rectangle2.Deselect 'or rectangle2.Selected=False
	Debug.Print "Rectangle2 selected? ";rectangle2.Selected

	rectangle2.Delete
	Debug.Print "Rectangle2 deleted."

	Wait 1
	shapes1("Text").Delete

	Debug.Print "AddRectangle - rounded rectangle"
	shapes1.AddText(1,1,"AddRectangle - rounded rectangle").Font.Size = 40
	Set roundedrect1 = shapes1.AddRectangle(Left:=5, Right:=8, _
		Top:=6, Bottom:=4, xRadius:=0.5, yRadius:=0.5)

	Wait 1
	shapes1("Text").Delete
	With rectangle1
		'The Fill property returns the FillFormat object, which can
		' be used to set the fill foreground color, pattern,
		' background color (for vector pattern types), and
		' background transparency.
		Debug.Print " Fill Property"
		shapes1.AddText(1,1,"Fill Property").Font.Size = 40
		.Fill.ForeColor = srfColorYellow
		.Fill.Pattern = "Solid"
		Wait 1
		shapes1("Text").Delete
		'The Line property returns the LineFormat object, which can
		' be used to set the line foreground color, style, and width.
		Debug.Print " Line Property"
		shapes1.AddText(1,1,"Line Property").Font.Size = 40
		.Line.ForeColor = srfColorBlue
		.Line.Width = 0.03
		.Line.Style = "Dash Dot"
		'The selected line style name is listed in the Style
		' dropdown list.  Use the Up and Down arrow keys to display
		' each name.  A complete list is in the attrib.ini file.

		Wait 1
		shapes1("Text").Delete
	End With

	'The xRadius and yRadius properties control the anount of
	' rounding at the corners of the rounded rectangle.
	' Units are page units (surf.PageUnits), a double.  These
	' properties can only be specified via automation.
	' There is no equivalent control in the user interface.
	With roundedrect1
		Debug.Print " xRadius, yRadius"
		'vbCrLf is the Visual BASIC enumeration constant for the
		' carriage return + line feed or chr(13) + chr(10).
		shapes1.AddText(1,2, _
			"Rounded Rectangle radii:"+vbCrLf+ _
			"  xRadius: " + .xRadius + _
			", yRadius: " + .yRadius).Font.Size = 40
	End With

	Wait 1
	shapes1.SelectAll
	plotdoc1.Selection.Delete

	'===========================
	'Ellipse Object Properties
	'===========================
	'The Ellipse object is created with the Shapes.AddEllipse method.
	Debug.Print "AddEllipse"
	shapes1.AddText(1,1,"AddEllipse").Font.Size = 40
	Set ellipse1 = shapes1.AddEllipse( Left:=1, Right:=4, _
		Top:=6, Bottom:=4)
	plotdoc1.Windows(1).Zoom(srfZoomPage)

	Wait 1
	shapes1("Text").Delete

	With ellipse1
		'The Fill property returns the FillFormat object, which can
		' be used to set the fill foreground color, pattern,
		' background color (for vector pattern types), and
		' background transparency.
		Debug.Print " Fill Property"
		shapes1.AddText(1,1,"Fill Property").Font.Size = 40
		.Fill.ForeColor = srfColorYellow
		.Fill.Pattern = "Solid"
		Wait 1
		shapes1("Text").Delete
		'The Line property returns the LineFormat object, which can
		' be used to set the line foreground color, style, and width.
		Debug.Print " Line Property"
		shapes1.AddText(1,1,"Line Property").Font.Size = 40
		.Line.ForeColor = srfColorBlue
		.Line.Width = 0.03
		.Line.Style = "Dash Dot"
		'The selected line style name is listed in the Style
		' dropdown list.  Use the Up and Down arrow keys to display
		' each name.  A complete list is in the attrib.ini file.

		Wait 1
		shapes1.SelectAll
		plotdoc1.Selection.Delete

	End With

	'=======================
	'Symbol Object property.
	'=======================
	Debug.Print "Symbol Object - Marker Property"
	shapes1.AddText(1,1,"Symbol Object - Marker Property").Font.Size = 40
	Set sym1 = shapes1.AddSymbol(4,4)
	sym1.Marker.Size =1
	sym1.Marker.Index=2

	Wait 1
	shapes1.SelectAll
	plotdoc1.Selection.Delete

	'=======================
	'Text Object properties.
	'=======================
	Debug.Print "Text Object Properties"
	Set text1 = shapes1.AddText(1,1,"Text Object Properties")
	text1.Font.Size = 40
	Wait 1
	text1.Text="Text Object - Font and Text Properties"

	Wait 1
	text1.Delete

	'==========================
	'Polyline Object Properties
	'==========================
	surf.Caption = "Surfer " + surf.Version
	AppActivate "Surfer " + surf.Version
	Debug.Print "Polyline Object Properties"
	shapes1.AddText(1,1,"Polyline Object Properties").Font.Size = 40
	Dim coords(0 To 7) As Double 'double array
	coordsarray = Array(4,4, 5,5, 6,4, 7,5) 'variant array
	For i = 0 To 7
		coords(i) = coordsarray(i)
	Next i
	Set polyline1 = shapes1.AddPolyLine(coords)
	Wait 1
	'The Line Property sets the line ForeColor, Width, and Style.
	Debug.Print " Polyline Object - Line Property"
	shapes1("Text").Text = " Polyline Object - Line Property"
	polyline1.Line.ForeColor = srfColorRed
	polyline1.Line.Width = 0.1
	polyline1.Line.Style = "Dash Dot"

	Wait 1

	Debug.Print " Polyline Object - StartArrow Property"
	shapes1("Text").Text = " Polyline Object - StartArrow Property"
	polyline1.StartArrow = srfASTriangle 'Arrow Style is Triangle
	Wait 1

	polyline1.ArrowScale = 5
	Debug.Print " Polyline Object - ArrowScale Property"
	shapes1("Text").Text = " Polyline Object - ArrowScale Property"
	Wait 1

	Debug.Print " Polyline Object - EndArrow Property"
	shapes1("Text").Text = " Polyline Object - EndArrow Property"
	polyline1.EndArrow = srfASSimple

	Debug.Print " Polyline Object - Vertices Property"
	shapes1("Text").Text = " Polyline Object - Vertices Property"
	Dim verts() As Double
	verts() = polyline1.Vertices
	Wait 1
	Debug.Print " Polyline starts at ";verts(0);", ";verts(1)
	shapes1("Text").Text = " Polyline starts at"+Str(verts(0))+","+Str(verts(1))

	Wait 1
	shapes1.SelectAll
	plotdoc1.Selection.Delete

	'=========================================
	'The Polygon Object Properties and Methods
	'=========================================
	Debug.Print "Polygon Object"
	shapes1.AddText(1,1,"Polygon Object").Font.Size = 40

	Dim coords2(0 To 9) As Double
	coordsarray2 = Array( _
		2.8, 2.2, _
		3.4, 2.7, _
		3.7, 2.2, _
		3.5, 1.8, _
		2.8, 2.2) 'Variant Array
	For i = 0 To 9 'Copy Variant array to Double array.
		coords2(i) = coordsarray2(i)
	Next i

	Set poly1 = shapes1.AddPolygon(coords2)
	'The Fill property sets the fill forecolor and pattern.
	poly1.Fill.ForeColor = srfColorLightYellow
	poly1.Fill.Pattern = "Solid"
	'The Line Property sets the line forecolor, style and width.
	poly1.Line.ForeColor = srfColorBlue
	poly1.Line.Width = 0.1
	poly1.Line.Style = "Solid"

	'The PolyCounts Property returns an array containing the number
	' of vertices per sub-polygon (for complex polygons).
	' It returns an array of longs.
	Set poly1 = shapes1("Polygon")
	Dim polycnts() As Long
	polycnts() = poly1.PolyCounts
	firstsubpoly = LBound(polycnts)
	lastsubpoly = UBound(polycnts)
	For i = firstsubpoly To lastsubpoly
		Debug.Print " Subpolygon #";i;" has";polycnts(i);" vertices.
	Next i

	'The Vertices Property returns an array containing
	' the coordinates of the vertices in the polygon.
	' It returns an array of doubles.
	Dim verts1() As Double
	verts1() = poly1.Vertices
	Debug.Print "Polygon Vertices"
	For i = LBound(verts1) To UBound(verts1) Step 2
		Debug.Print verts1(i);verts1(i+1)
	Next i
	Wait 1

	'The SetVertices Method changes the vertices in an existing
	' polygon.
	Debug.Print "Change vertex #3."
	shapes1("Text").Text = "Change Vertex #3"
	Dim verts2() As Double
	Dim polycnt2() As Long
	verts2() = poly1.Vertices
	polycnt2() = poly1.PolyCounts
	'Vertex 3 coordinates are at indices 4 (x) and 5 (y).
	Debug.Print "Old coordinates: ";verts2(4);verts2(5)
	'Change the third vertex coordinates.
	verts2(4) = 5
	verts2(5) = 3
	Debug.Print "New coordinates: ";verts2(4);verts2(5)
	poly1.SetVertices( _
		Vertices:=verts2(), _
		PolyCounts:=polycnt2() )
	Wait 1

	'The Composite Object contains one or more Shape objects.
	' It is created with the Selection.Combine method.
	Debug.Print "Composite Object"
	shapes1("Text").Text = "Composite Object"
	Set rect1 = shapes1.AddRectangle(4,5,6,7)
	rect1.Select
	poly1.Select
	Set composite1 = plotdoc1.Selection.Combine
	Wait 1

	'The BreakApart Method breaks a Composite Object into
	' its component objects.
	shapes1("Text").Text = "Break Apart"
	composite1.BreakApart


End Sub
