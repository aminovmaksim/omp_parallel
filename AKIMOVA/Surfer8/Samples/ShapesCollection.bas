'ShapesCollection.bas illustrates the properties and methods
' of the Shapes Collection.
Sub Main
	Debug.Print "----- ShapesCollection.bas - ";Time;" -----"

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


	'============================
	'Shapes Collection Properties
	'============================
	Set shapes1 = plotdoc1.Shapes
	If shapes1.Count >0 Then
		shapes1.SelectAll
		plotdoc1.Selection.Delete
	End If

	'-------------------------------------------------------
	'The Application Property returns the application object.
	'-------------------------------------------------------
	Debug.Print "Shapes Collection Aapplication: ";shapes1.Application

	'------------------------------------------------------------------
	'The Parent Property returns the shapes collection parent. (object)
	'------------------------------------------------------------------
	Debug.Print "Shapes Collection Parent: ";shapes1.Parent

	'--------------------------------------------------------
	'The Count Property returns the number of items in the
	' shapes collection (integer).
	'--------------------------------------------------------
	Debug.Print "The Shapes Collection has";shapes1.Count;" items."

	'========================
	'ShapesCollection Methods
	'========================---------------------------------
	'The AddBaseMap method creates a new base map.  It returns
	' a MapFrame object.
	'---------------------------------------------------------
	surf.Caption = "Surfer "+surf.Version
	AppActivate "Surfer "+surf.Version
	Debug.Print "AddBaseMap"
	shapes1.AddText(1,1,"AddBaseMap").Font.Size = 40
	Set mapframe1 = shapes1.AddBaseMap(path1+"ca.gsb")
	Set base1 = mapframe1.Overlays("Base")
	base1.Fill.ForeColor = srfColorPaleYellow
	base1.Fill.Pattern = "Solid"
	mapframe1.BackgroundFill.ForeColor = srfColorBabyBlue
	mapframe1.BackgroundFill.Pattern = "Solid"
	For Each Axis In mapframe1.Axes
		Axis.ShowLabels = False
		Axis.MajorTickType = srfTickNone
		Axis.MinorTickType = srfTickNone
	Next Axis
	With mapframe1
		.xLength = 2.5
		.yLength = 2
		.Top = 8.25
		.Left = 0.25
		.Axes("Bottom Axis").Title = "Base"
		.Axes("Bottom Axis").TitleFont.Size = 25
	End With

	Wait 1
	shapes1("Text").Delete

	'---------------------------------------------------------
	'The AddClassedPostMap method adds a new classed post map.
	' It returns a MapFrame object.
	'---------------------------------------------------------
	Debug.Print "AddClassedPostMap"
	shapes1.AddText(1,1,"AddClassedPostMap").Font.Size = 40
	Set mapframe2 = shapes1.AddClassedPostMap(path1+"demogrid.dat", _
		xCol:=1, yCol:=2, zCol:=3)
	mapframe2.BackgroundFill.ForeColor = srfColorWhite
	mapframe2.BackgroundFill.Pattern = "Solid"
	Set clpost1 = mapframe2.Overlays("Classed Post")
	clpost1.ShowLegend = False
	For i = 1 To 5
		With clpost1.BinSymbol(i)
			.Set = "GSI Default Symbols"
			.Index = 12
			.Size = 0.1
		End With
	Next i
	clpost1.BinSymbol(1).Color = srfColorPastelBlue
	clpost1.BinSymbol(2).Color = srfColorGrassGreen
	clpost1.BinSymbol(3).Color = srfColorDeepYellow
	clpost1.BinSymbol(4).Color = srfColorLightOrange
	clpost1.BinSymbol(5).Color = srfColorBrickRed
	For Each Axis In mapframe2.Axes
		Axis.ShowLabels = False
		Axis.MajorTickType = srfTickNone
		Axis.MinorTickType = srfTickNone
	Next Axis
	With mapframe2
		.xLength = 2.5
		.yLength = 2
		.Top = 8.25
		.Left = 2.9
		.Axes("Bottom Axis").Title = "Classed Post"
		.Axes("Bottom Axis").TitleFont.Size = 25
	End With

	Wait 2
	shapes1("Text").Delete

	'---------------------------------------------------------
	'The AddContourMap method adds a contour map.  It returns
	' a MapFrameObject.
	'---------------------------------------------------------
	Debug.Print "AddContourMap"
	shapes1.AddText(1,1,"AddContourMap").Font.Size = 40
	Set mapframe3 = 	shapes1.AddContourMap(path1+"demogrid.grd")
	Set contours1 = mapframe3.Overlays("Contours")
	contours1.ShowColorScale = False
	contours1.FillContours = True

	For Each Axis In mapframe3.Axes
		Axis.ShowLabels = False
		Axis.MajorTickType = srfTickNone
		Axis.MinorTickType = srfTickNone
	Next Axis
	With mapframe3
		.xLength = 2.5
		.yLength = 2
		.Top = 8.25
		.Left = 5.6
		.Axes("Bottom Axis").Title = "Contours"
		.Axes("Bottom Axis").TitleFont.Size = 25
	End With

	Wait 2
	shapes1("Text").Delete

	'-------------------------------------------------
	'The AddImageMap method creates a new image map.
	' It returns a MapFrame object.
	'-------------------------------------------------
	Debug.Print "AddImageMap"
	shapes1.AddText(1,1,"AddImageMap").Font.Size = 40
	Set mapframe4 = shapes1.AddImageMap(path1+"demogrid.grd")
	Set imagemap1 = mapframe4.Overlays("Image Map")
	imagemap1.ShowColorScale = False
	For Each Axis In mapframe4.Axes
		Axis.ShowLabels = False
		Axis.MajorTickType = srfTickNone
		Axis.MinorTickType = srfTickNone
	Next Axis
	With mapframe4
		.xLength = 2.5
		.yLength = 2
		.Top = 8.25
		.Left = 8.25
		.Axes("Bottom Axis").Title = "Image Map"
		.Axes("Bottom Axis").TitleFont.Size = 25
	End With

	Wait 2
	shapes1("Text").Delete

	'------------------------------------------
	'The AddPostMap method adds a new post map.
	' It returns a MapFrame Object.
	'------------------------------------------
	Debug.Print "AddPostMap"
	shapes1.AddText(1,1,"AddPostMap").Font.Size = 40
	Set mapframe5 = shapes1.AddPostMap(path1+"demogrid.dat")
	mapframe5.BackgroundFill.ForeColor = srfColorWhite
	mapframe5.BackgroundFill.Pattern = "Solid"
	Set postmap1 = mapframe5.Overlays("Post")

	For Each Axis In mapframe5.Axes
		Axis.ShowLabels = False
		Axis.MajorTickType = srfTickNone
		Axis.MinorTickType = srfTickNone
	Next Axis
	With mapframe5
		.xLength = 2.5
		.yLength = 2
		.Top = 5.5
		.Left = 0.25
		.Axes("Bottom Axis").Title = "Post Map"
		.Axes("Bottom Axis").TitleFont.Size = 25
	End With

	Wait 2
	shapes1("Text").Delete

	'-----------------------------------------------------
	'The AddReliefMap method adds a new shaded relief map.
	' It returns a MapFrame Object.
	'-----------------------------------------------------
	Debug.Print "AddReliefMap"
	shapes1.AddText(1,1,"AddReliefMap").Font.Size = 40
	Set mapframe6 = shapes1.AddReliefMap(path1+"helens2.grd")
	For Each Axis In mapframe6.Axes
		Axis.ShowLabels = False
		Axis.MajorTickType = srfTickNone
		Axis.MinorTickType = srfTickNone
	Next Axis
	With mapframe6
		.xLength = 2.5
		.yLength = 2
		.Top = 5.5
		.Left = 2.9
		.Axes("Bottom Axis").Title = "Shaded Relief Map"
		.Axes("Bottom Axis").TitleFont.Size = 25
	End With

	Set relief1 = mapframe6.Overlays("Shaded Relief Map")

	Wait 2
	shapes1("Text").Delete

	'----------------------------------------------
	'The AddVectorMap method adds a new vector map.
	' It returns a MapFrame Object.
	'----------------------------------------------
	Debug.Print "AddVectorMap"
	shapes1.AddText(0.5,1,"AddVectorMap").Font.Size = 40
	Set mapframe7 = shapes1.AddVectorMap(path1+"demogrid.grd")
	mapframe7.BackgroundFill.ForeColor = srfColorWhite
	mapframe7.BackgroundFill.Pattern = "Solid"
	For Each Axis In mapframe7.Axes
		Axis.ShowLabels = False
		Axis.MajorTickType = srfTickNone
		Axis.MinorTickType = srfTickNone
	Next Axis
	With mapframe7
		.xLength = 2.5
		.yLength = 2
		.Top = 5.5
		.Left = 5.6
		.Axes("Bottom Axis").Title = "Vector Map"
		.Axes("Bottom Axis").TitleFont.Size = 25
	End With

	Set vectors1 = mapframe7.Overlays("Vectors")

	Wait 2
	shapes1("Text").Delete

	'----------------------------------------------
	'The AddWireframe method adds a new wireframe map.
	' It returns a MapFrame Object.
	'----------------------------------------------
	Debug.Print "AddWireframe"
	shapes1.AddText(0.5,1,"AddWireframe").Font.Size = 40
	Set mapframe8 = shapes1.AddWireframe(path1+"demogrid.grd")
	mapframe8.BackgroundFill.ForeColor = srfColorWhite
	mapframe8.BackgroundFill.Pattern = "Solid"
	Set wireframe1 = mapframe8.Overlays("Wireframe")
	wireframe1.ShowColorScale = False
	For Each Axis In mapframe8.Axes
		Axis.ShowLabels = False
		Axis.MajorTickType = srfTickNone
		Axis.MinorTickType = srfTickNone
	Next Axis
	With mapframe8
		.xLength = 2
		.yLength = 1.5
		.zLength = .5
		.Top = 5.5
		.Left = 8.25
	End With
	mapframe8.ViewTilt = 20
	mapframe8.ViewProjection = srfOrthographic
	Set wireframetext = shapes1.AddText(8.75,3.25,"Wireframe")
	wireframetext.Font.Size = 25
	wireframetext.Name = "Wireframe Text"

	Wait 2
	shapes1("Text").Delete

	'----------------------------------------------
	'The AddSurface method adds a new surface map.
	' It returns a MapFrame Object.
	'----------------------------------------------
	If Val(Left(surf.Version,1)) >7 Then
		Debug.Print "AddSurface"
		shapes1.AddText(0.5,1,"AddSurface").Font.Size = 40
		Set mapframe9 = shapes1.AddSurface(path1+"demogrid.grd")
		Set surface1 = mapframe9.Overlays("3D Surface")
		surface1.ShowColorScale = False
		For Each Axis In mapframe9.Axes
			Axis.ShowLabels = False
			Axis.MajorTickType = srfTickNone
			Axis.MinorTickType = srfTickNone
		Next Axis
		With mapframe9
			.xLength = 2
			.yLength = 1.5
			.zLength = .5
			.Top = 2.75
			.Left = 8.00
		End With
		mapframe9.ViewTilt = 20
		mapframe9.ViewProjection = srfOrthographic
		Set surfacetext = shapes1.AddText(8.75,0.75,"Surface")
		surfacetext.Font.Size = 25
		surfacetext.Name = "Surface Text"

		Set surface1 = mapframe9.Overlays("3D Surface")

		Wait 2
		shapes1("Text").Delete
	End If
	'-------------------------------------------------------------------------
	'The AddComplexPolygon method adds a new complex polygon using page units.
	' It returns a Polygon Object.
	'-------------------------------------------------------------------------
	Debug.Print "AddComplexPolygon"
	shapes1.AddText(1,1,"AddComplexPolygon").Font.Size = 40
	Dim coords (0 To 23) As Double
	Dim numpolys(0 To 2) As Long
	coordarray = Array( _
						3.44, 4.06, _
					 	1.10, 6.39, _
						3.44, 8.73, _
						5.75, 6.39, _
						3.36, 8.07, _
						5.01, 6.42, _
						3.36, 4.75, _
						1.71, 6.42, _
						1.71, 8.07, _
						5.01, 8.07, _
						5.01, 4.75, _
						1.71, 4.75 ) 'Array returns a variant.
	For i = 0 To 23
		coords(i) = coordarray(i) 'Copy variant array to double array.
	Next i
	numpolys(0) = 4
	numpolys(1) = 4
	numpolys(2) = 4
	Set complexpoly = shapes1.AddComplexPolygon(vertices:=coords, _
		PolyCounts:=numpolys)
	With complexpoly
		.Fill.ForeColor = srfColorBlue
		.Fill.Pattern = "Solid"
		.Left = 6.25
		.Top = 2.75
		.Height = 1
		.Width = 1
	End With
	Set complexpolytext = shapes1.AddText(6.25,1.5,"Complex"+vbCrLf+"Polygon")
	With complexpolytext
		.Name = "Text Complex Polygon"
		.Font.Size = 25
	End With

	Wait 2
	shapes1("Text").Delete

	'---------------------------------------------------------
	'The AddEllipse adds a new ellipse shape using page units.
	' It returns an Ellipse Object.
	'---------------------------------------------------------
	Debug.Print "AddEllipse"
	shapes1.AddText(1,1,"AddEllipse").Font.Size = 40
	Set ellipse1 = shapes1.AddEllipse( _
		Left:=5, Top:=2.5, Right:=6, Bottom:=2)
	ellipse1.Fill.ForeColor = srfColorRed
	ellipse1.Fill.Pattern = "Solid"
	Set ellipsetext = shapes1.AddText(5,1.85,"Ellipse")
	ellipsetext.Font.Size = 25
	ellipsetext.Name = "Text Ellipse"

	Wait 2
	shapes1("Text").Delete

  '-----------------------------------------------------------
	'The AddLine method adds a new line with two vertices using
	' page coordinates.  It returns a Polyline Object.
  '-----------------------------------------------------------
	Debug.Print "AddLine"
	shapes1.AddText(1,1,"AddLine").Font.Size = 40
	Set line1 = shapes1.AddLine(4, 2.25, _
		4.5,2.5)
	line1.EndArrow = srfASFilled
	Set linetext1 = shapes1.AddText(4,2.1,"Line")
	linetext1.Font.Size = 25
	linetext1.Name = "Text Line"

	'The Add

	Wait 2
	shapes1("Text").Delete

	'----------------------------------------------------
	'The AddPolygon method adds a new polygon shape using
	' page coordinates. It returns a Polygon object.
	'----------------------------------------------------
	Debug.Print "AddPolygon"
	shapes1.AddText(1,1,"AddPolygon").Font.Size = 40
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
	poly1.Fill.ForeColor = srfColorGreen
	poly1.Fill.Pattern = "Solid"
	Set polygontext = shapes1.AddText(2.6,1.7,"Polygon")
	polygontext.Font.Size = 25
	polygontext.Name = "Text Polygon"

	Wait 2
	shapes1("Text").Delete

	'-------------------------------------------------
	'The AddPolyline method adds a new polyline shape
	' using page units.  It returns a Polyline object.
	'-------------------------------------------------
	Debug.Print "AddPolyline"
	shapes1.AddText(1,1,"AddPolyline").Font.Size = 40
	Dim polylinecoords (0 To 7) As Double
	polylinecoordsarray = Array( _
		1.6, 2.1, _
		2, 2.6, _
		2.1, 2.2, _
		2.4, 2.6) 'Variant Array
	For i = 0 To 7
		polylinecoords(i) = polylinecoordsarray(i)
	Next i
	shapes1.AddPolyLine(polylinecoords)
	Set polylinetext = shapes1.AddText(1.25, 2, "Polyline")
	polylinetext.Font.Size = 25
	polylinetext.Name = "Text Polyline"

	Wait 2
	shapes1("Text").Delete 'Delete AddPolyline text.

	'--------------------------------------------------
	'The AddRectangle method adds a new rectangle shape
	' using page units.  It returns a Rectangle Object.
	'--------------------------------------------------
	Debug.Print "AddRectangle"
	shapes1.AddText(1,1,"AddRectangle").Font.Size = 40
	Set rectangle1 = shapes1.AddRectangle(Left:=0.5, Top:=2.5, _
		Right:=1, Bottom:=1.75)
	rectangle1.Fill.ForeColor = srfColorYellow
	rectangle1.Fill.Pattern = "Solid"
	Set rectangletext = shapes1.AddText(0.25, 1.5, "Rectangle")
	rectangletext.Name = "Text Rectangle"
	rectangletext.Font.Size = 25

	Wait 2
	shapes1("Text").Delete

	'---------------------------------------------------
	'The AddSymbol method adds a new symbol shape using
	' page units.  It returns a Symbol Object.
	'---------------------------------------------------
	Debug.Print "AddSymbol"
	shapes1.AddText(1,1,"AddSymbol").Font.Size = 40
	Set symbol1 = shapes1.AddSymbol(5.1, 1.4)
	symbol1.Marker.Size = 0.5
	symbol1.Marker.Set = "GSI Default Symbols"
	symbol1.Marker.Index = 100
	symbol1.Left = 5.1
	symbol1.Top = 1.35
	Set symboltext = shapes1.AddText(5,0.75, "Symbol")
	symboltext.Font.Size = 25
	symboltext.Name = "Text Symbol"

	Wait 2
	shapes1("Text").Delete

	'---------------------------------------------------
	'The AddText method adds a new text shape using page
	' units.  It returns a Text object.
	'---------------------------------------------------
	Debug.Print "AddText"
	shapes1.AddText(1,1,"AddText").Font.Size = 40

	Wait 2
	shapes1("Text").Delete

	'---------------------------------------------------
	'The AddVariogram method adds a new Variogram plot.
	' It returns a Variogram Object.
	'---------------------------------------------------
	Set plotdoc2 = surf.Documents.Add
	Set shapes2 = plotdoc2.Shapes
	Debug.Print "AddVariogram"
	With plotdoc2.PageSetup
		.Orientation = srfLandscape
		.Height = 8.5
		.Width = 11
	End With
	plotdoc2.Windows(1).Zoom(srfZoomPage)

	Set text1 = shapes2.AddText(1,1,"AddVariogram")
	text1.Font.Size = 35
	Set vario1 = Shapes2.AddVariogram(path1+"demogrid.dat")
	Wait 1
	text1.Text = "Add New Variogram Model Components"
	Wait 1
	'Add new variogram model components.
	Dim variocomponents(1 To 2) As Object
	Set variocomponents(1) = surf.NewVarioComponent(srfVarNugget,10,0)
	Set variocomponents(2) = surf.NewVarioComponent(srfVarGaussian,250,1.5)
	vario1.Model = variocomponents

	Wait 2
	shapes2("Text").Delete

	'----------------------------------------------------
	'The BlockSelect method selects all shapes within the
	' specified rectangle.
	'----------------------------------------------------
	Debug.Print "BlockSelect"
	plotdoc1.Activate
	shapes1.AddText(1,1,"BlockSelect").Font.Size =40
	shapes1.BlockSelect(Left:=0, Right:=2.62, _
		top:=2.75, bottom:=1.00)

	Wait 2
	shapes1("Text").Delete

	'--------------------------------------------------
	'The InvertSelection method selects all deselected
	' objects and deselects all selectd objects.
	'--------------------------------------------------
	Debug.Print "InvertSelection"
	shapes1.AddText(1,1,"InvertSelection").Font.Size = 40
	shapes1.InvertSelection

	Wait 2
	shapes1("Text").Delete

	'-------------------------------------------------
	'The Item method returns an individual item from a
	' collection.  It is the default method.
	'-------------------------------------------------
	Debug.Print "Item Method"
	shapes1.AddText(1,1,"Item Method").Font.Size = 40
	plotdoc1.Selection.DeselectAll
	'The following statements are equivalent.
	shapes1.Item("Text").Select
	shapes1("Text").Select
	shapes1(shapes1.Count).Select

	Wait 2
	shapes1("Text").Delete

	'------------------------------------------------------
	'The Paste method pastes the Clipboard contents to the
	' center of the page.  It returns an object.
	'------------------------------------------------------
	Debug.Print "Paste method"
	plotdoc1.Activate
	shapes1.AddText(1,1,"Paste method").Font.Size = 40
	'The Copy method is used by the Selection Collection.
	shapes1.BlockSelect(Left:=0, Right:=2.62, _
		top:=2.75, bottom:=1.00)
	plotdoc1.Selection.Copy
	plotdoc2.Activate
	shapes2(1).Delete
	shapes2.AddText(1,1,"Paste method").Font.Size = 40
	Set selectioncoll2 =Shapes2.Paste(Format:=srfPasteBest)
	For i = 1 To selectioncoll2.Count
		Debug.Print "  ";selectioncoll2(i)
	Next

	Wait 2
	shapes1("Text").Delete
	plotdoc1.Activate

	'---------------------------------------------------
	'The SelectAll method selects all the shapes in the
	' shapes collection.
	'---------------------------------------------------
	Debug.Print "SelectAll"
	shapes1.AddText(1,1,"SelectAll").Font.Size = 40
	shapes1.SelectAll
	For Each shp In plotdoc1.Selection
		Debug.Print "  ";shp
	Next

	Wait 2
	shapes1("Text").Delete

End Sub
