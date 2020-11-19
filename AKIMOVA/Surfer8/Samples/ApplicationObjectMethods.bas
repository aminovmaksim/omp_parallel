'ApplicationObjectMethods.bas demonstrates the methods
' of the Surfer Application Object.
' See ApplicationObjectProperties.bas for the properties of the
' Surfer Application Object.
' TB - 17 Oct 01.
Sub Main
	Debug.Print "----- ";Time;" -----"

	'Get existing Surfer instance, or create a new one If none exists.
	On Error Resume Next 'Turn off error reporting.
	Set surf = GetObject(,"Surfer.Application")
	If Err.Number<>0 Then
		Set surf = CreateObject("Surfer.Application")
	End If
	On Error GoTo 0 'Turn on error reporting.
	If surf.Windows.Count=0 Then surf.Documents.Add(srfDocPlot)

	surf.Visible = True
	surf.WindowState = srfWindowStateNormal
	surf.Width = 600
	surf.Height = 400
	surf.Windows(1).Zoom(srfZoomPage)

	Debug.Print "--------------------------------------------------"
	Debug.Print "Surfer ";surf.Version;" Application Object Methods"
 	Debug.Print "--------------------------------------------------"

	path1 = surf.Path +"\samples\"
	Debug.Print surf.Documents(1).Type = srfDocPlot
	If surf.Documents(1).Type = srfDocPlot Then
		Set plotdoc1 = surf.Documents(1)
	Else
		Set plotdoc1 = surf.Documents.Add
	End If
	Set plotwin1 = plotdoc1.Windows(1)
	plotwin1.Activate
	plotwin1.Zoom(srfZoomPage)

	Set shapes1 = plotdoc1.Shapes
	AppActivate "Surfer "
	plotdoc1.PageSetup.Orientation = srfLandscape

	'-----------------------------------------------
	' Use Grid Blank to blank a grid file using a BLN file.
	'-----------------------------------------------
	Debug.Print "GridBlank"
	surf.GridBlank (InGrid := path1+"demogrid.grd", _
		BlankFile := path1+"demorect.bln", _
		OutGrid := path1+"demoblanked.grd", _
		OutFmt := srfGridFmtS7)
		'srfGridFmtBinary,srfGridFmtAscii,srfGridFmtS7,srfGridFmtXYZ
	If shapes1.Count > 0 Then
		For Each shp In shapes1
			shp.Delete
		Next
	End If
	shapes1.AddText(1,1,"GridBlank").Font.Size = 40
	Set mapframe1 = shapes1.AddContourMap(path1+"demoblanked.grd")
	Set contour1 = mapframe1.Overlays("Contours")
	'-----------------------------------------------
	'Calculate a new grid with slopes in degrees using
	' GridCalculus and helens2.grd.
	'-----------------------------------------------
	Debug.Print "GridCalculus"
	surf.GridCalculus(InGrid := path1+"helens2.grd", _
		Operation := srfGCSlope, _
		OutGrid := path1+"helens2slope.grd", _
		OutFmt := srfGridFmtS7)
		' Operations: Right-click | Quick Watch to display numeric value.
		' srfGCFirstDeriv,srfGCSecondDeriv,srfGCCurvature,srfGCSlope,srfGCAspect,
		' srfGCProfCurv,srfGCPlanCurv,srfGCTanCurv,srfGCGradient,srfGCLaplacian,
		' srfGCBiharmonic,srfGCVolume,srfGCCorrelogram,srfGCPeriodogram

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridCalculus Slope").Font.Size = 40
	Set mapframe1 = shapes1.AddImageMap(path1 +"helens2slope.grd")
	Set imagemap1 = mapframe1.Overlays("Image Map")

	'-----------------------------------------------
	'Use GridConvert to convert a grid to an XYZ data file.
	'-----------------------------------------------
	Debug.Print "GridConvert"
	surf.GridConvert (InGrid := path1+"\demogrid.grd", _
		OutGrid := path1+"\demogrid2.dat", _
		OutFmt := srfGridFmtXYZ)

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridConvert").Font.Size = 40
	Set mapframe1 = shapes1.AddClassedPostMap(path1+"demogrid2.dat")
	Set classedpost1 = mapframe1.Overlays("Classed Post")

	'-----------------------------------------------
	'GridExtract extracts a subset of a grid.
	'-----------------------------------------------
	Debug.Print "GridExtract"
	surf.GridExtract (InGrid := path1+"demogrid.grd", _
		r1:=1, r2:=26, c1:=1, c2:=25, _
		OutGrid := path1+"DemoExtract.grd", _
		OutFmt := srfGridFmtS7)

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridExtract").Font.Size = 40
	Set mapframe1 = shapes1.AddContourMap(path1+"demoextract.grd")
	Set contour1 = mapframe1.Overlays("Contours")

	'------------------------------------------------------
	'GridFilter in Surfer 8 (GridMatrixSmooth in Surfer 7).
	'------------------------------------------------------
	Debug.Print "GridFilter"
	If Left(surf.Version, 1) = "8" Then
		surf.GridFilter(path1+"demogrid.grd", _
			srfFilterEmbSouthwest, _
			outgrid:=path1+"DemoFilter.grd")
		'_Application.GridFilter(_InGrid, _Filter, EdgeOp:=_, BlankOp:=_, NumPasses:=_, EdgeFill:=_, BlankFill:=_, NumRow:=_, NumCol:=_, Param1:=_, Param2:=_, UserFilter:=_, OutGrid:=_, OutFmt:=_)

		Wait 1
		For Each shp In shapes1
			shp.Delete
		Next
		shapes1.AddText(1,1,"GridFilter Emboss Southwest").Font.Size=40
		shapes1.AddContourMap(path1+"demofilter.grd")

	End If

	'-----------------------------------------------
	'Use GridFunction to create a new grid from a function.
	'-----------------------------------------------
	Debug.Print "GridFunction"
	surf.GridFunction( _
		Function:="z=(pow(x,2)+pow(y,2))*(sin(8*atan2(x,y)))", _
		xMin:=-25, xMax:=25, xInc:=1, _
		yMin:=-25, yMax:=25, yInc:=1, _
		OutGrid := path1 + "GridFunction.grd", _
		OutFmt := srfGridFmtS7)

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridFunction").Font.Size = 40
	Set mapframe1 = shapes1.AddWireframe(path1+"GridFunction.grd")
	Set wireframe1 = mapframe1.Overlays("Wireframe")

	'-----------------------------------------------
	'Use GridData to grid two columns from a data file
	'	and GridMath to compare the two grid files.
	'-----------------------------------------------
	Debug.Print "GridMath"
	surf.GridData(DataFile := path1+"sample3.dat", _
		xCol:=1, yCol:=2, zCol:=3, _
		Algorithm := srfKriging, _
		ShowReport := False, _
		OutGrid := path1+"Sample3a.grd", _
		OutFmt := srfGridFmtS7)
	surf.GridData(DataFile := path1+"sample3.dat", _
		xCol:=1, yCol:=2, zCol:=4, _
		Algorithm := srfKriging, _
		ShowReport := False, _
		OutGrid := path1+"Sample3b.grd", _
		OutFmt := srfGridFmtS7)
	surf.GridMath(Function := "C = A - B", _
		InGridA := path1+"Sample3a.grd", _
		InGridB := path1+"Sample3b.grd", _
		OutGridC := path1+"Sample3a-b.grd", _
		OutFmt := srfGridFmtS7)

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridMath").Font.Size = 40
		Set mapframe1 = shapes1.AddContourMap(path1+"sample3a.grd")
	With mapframe1
		.xLength = .xLength/2
		.yLength = .yLength/2
		.Top = 8
		.Left = 1
		.Axes("Bottom Axis").Title = "Sample3a.grd"
		.Axes("Bottom Axis").TitleFont.Size = 20
	End With

	Set mapframe2 = shapes1.AddContourMap(path1+"sample3b.grd")
	With mapframe2
		.xLength = .xLength/2
		.yLength = .yLength/2
		.Top = 8
		.Left = 5
		.Axes("Bottom Axis").Title = "Sample3b.grd"
		.Axes("Bottom Axis").TitleFont.Size = 20
	End With

	Set mapframe3 = shapes1.AddContourMap(path1+"sample3a-b.grd")
	With mapframe3
		.xLength = .xLength/2
		.yLength = .yLength/2
		.Top = 4
		.Left = 5
		.Axes("Bottom Axis").Title = "Sample3a-b.grd"
		.Axes("Bottom Axis").TitleFont.Size = 20
	End With

	'----------------------------------------------------
	'GridMosaic combines two or more adjacent grid files.
	'----------------------------------------------------
	Debug.Print "Grid Mosaic"

	'-----------------------------------------------
	'GridResiduals calculates the difference between the grid
	' and data file.
	'-----------------------------------------------
	Debug.Print "GridResiduals"
	surf.GridData(path1+"Demogrid.dat", _
		algorithm:=srfRegression, _
		RegrMaxXOrder:=2, RegrMaxYOrder:=2, _
		RegrMaxTotalOrder:=4, _
		ShowReport:=False, _
		outgrid:=path1+"DemoRegression.grd")
	surf.GridResiduals(InGrid := path1+"DemoRegression.grd", _
		DataFile := path1+"demogrid.dat", _
		xCol:=1, yCol:=2, zCol:=3, ResidCol:=4)
	Set wksdoc1 = surf.Documents("Demogrid.dat")
	wksdoc1.Close(srfSaveChangesYes)
	surf.GridData(path1+"demogrid.dat",zcol:=4, _
		algorithm:=srfKriging, _
		ShowReport:=False, _
		outgrid:=path1+"DemoResiduals.grd")

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridResiduals").Font.Size=40
	Set mapframe1 = shapes1.AddClassedPostMap(path1+"demogrid.dat", _
		zcol:=4)
	Set classedpost1 = mapframe1.Overlays("Classed Post")
	classedpost1.zCol = 4 'required in s7, fixed in s8.
	Set mapframe2 = shapes1.AddContourMap(path1+"demoresiduals.grd")
	mapframe1.Select
	mapframe2.Select
	Set mapframe3 = plotdoc1.Selection.OverlayMaps
	mapframe3.Overlays("Classed Post").SetZOrder(srfZOToFront)

	'-----------------------------------------------
	'GridSlice calculates the XYZ values and accumulated
	' distance along a cross-section profile.
	'-----------------------------------------------
	Debug.Print "GridSlice"
	surf.GridSlice (InGrid := path1+"demogrid.grd", _
		BlankFile := path1+"DemoSlice.bln", _
		OutDataFile := path1+"Demoslice.dat", _
		OutsideVal:=-8888, BlankVal:=-9999)

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridSlice").Font.Size = 40
	Set mapframe1 = shapes1.AddContourMap(path1+"demogrid.grd")
	Set mapframe2 = shapes1.AddBaseMap(path1+"demoslice.bln")
	mapframe1.Select
	mapframe2.Select
	plotdoc1.Selection.OverlayMaps
	With mapframe1
		.xLength = 0.67 * .xLength
		.yLength = 0.67 * .yLength
		.Top = 8
		.Left = 3
	End With

	Set mapframe2 = shapes1.AddPostMap(path1+"demoslice.dat", _
		xcol:=4, ycol:=3)
	With mapframe2
		.xLength = 4
		.yLength = 2
		.Top = 4
		.Left = 3
	End With

	'-----------------------------------------------
	'GridSplineSmooth interpolates new grid nodes
	' with a cubic spline.
	'-----------------------------------------------
	Debug.Print "GridSplineSmooth"
	surf.GridSplineSmooth (InGrid := path1+"demogrid.grd", _
		nRow:=2, nCol:=4, Method:=srfSplineInsert, _
		OutGrid := path1+"DemoSplineSmooth.grd", _
		OutFmt := srfGridFmtS7)

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridSplineSmooth").Font.Size = 40
	Set mapframe1 = shapes1.AddContourMap(path1+"demogrid.grd")
	With mapframe1
		.xLength = 0.67 * .xLength
		.yLength = 0.67 * .yLength
		.Top = 8
		.Left = 1
		.Axes("Bottom Axis").Title = "Demogrid.grd"
		.Axes("Bottom Axis").TitleFont.Size = 20
	End With

	Set mapframe2 = shapes1.AddContourMap(path1+"DemoSplineSmooth.grd")
	With mapframe2
		.xLength = 0.67 * .xLength
		.yLength = 0.67 * .yLength
		.Top = 8
		.Left = 6
		.Axes("Bottom Axis").Title = "DemoSplineSmooth.grd"
		.Axes("Bottom Axis").TitleFont.Size = 20
	End With

	'-----------------------------------------------
	'GridTransform performs XY transforms on a grid
	' including Scale, Offset, Mirror, and Rotate.
	'-----------------------------------------------
	Debug.Print "GridTransform"
	Surf.GridTransform (InGrid := path1+"demogrid.grd", _
		Operation := srfGridTransOffset, _
		xOffset:=3, yOffset:=-5, _
		OutGrid:=path1+"demotransform1.grd", _
		OutFmt:=srfGridFmtS7)
	Surf.GridTransform(InGrid := path1+"demogrid.grd", _
		Operation:=srfGridTransMirrorY, _
		OutGrid:=path1+"demotransform2.grd", _
		OutFmt:=srfGridFmtS7)

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1,"GridTransform").Font.Size = 40
	Set mapframe1 = shapes1.AddContourMap(path1+"demogrid.grd")
	With mapframe1
		.xLength = .xLength/2
		.yLength = .yLength/2
		.Top = 8
		.Left = 1
		.Axes("Bottom Axis").Title = "Demogrid.grd"
		.Axes("Bottom Axis").TitleFont.Size = 20
	End With

	Set mapframe2 = shapes1.AddContourMap(path1+"DemoTransform1.grd")
	With mapframe2
		.xLength = .xLength/2
		.yLength = .yLength/2
		.Top = 8
		.Left = 5
		.Axes("Bottom Axis").Title = "DemoTransform1.grd Offset"
		.Axes("Bottom Axis").TitleFont.Size = 20
	End With

	Set mapframe3 = shapes1.AddContourMap(path1+"DemoTransform2.grd")
	With mapframe3
		.xLength = .xLength/2
		.yLength = .yLength/2
		.Top = 4
		.Left = 5
		.Axes("Bottom Axis").Title = "DemoTransform2.grd MirrorY"
		.Axes("Bottom Axis").TitleFont.Size = 20
	End With

	'-----------------------------------------------
	'GridVolume calculates volumes and areas between
	' two surfaces.
	'-----------------------------------------------
	Debug.Print "GridVolume"
	Dim results() As Double 'Array of volume results.
	'Dim resultnames() As String
	AppActivate "ApplicationObjectMethods"
	surf.GridVolume (Upper := path1+"demogrid.grd", _
		Lower:=50, pResults:=results, ShowReport:=False)
	resultnames = Array( _
		"srfGVTrapVol", "srfGVSimpVol", "srfGVSimp38Vol", "srfGVPosVol", "srfGVNegVol", _
		"srfGVPosPlanarArea", "srfGVNegPlanarArea", "srfGVPosArea", "srfGVNegArea", "srfGVBlankedArea", _
 		"srfGVNumParams")
	For i = 0 To srfGVNumParams-1 Step 1
		Debug.Print resultnames(i);" ";results(i)
	Next
	Debug.Print "srfGVNumParams ";srfGVNumParams
	'srfGVTrapVol,srfGVSimpVol,srfGVSimp38Vol,srfGVPosVol,srfGVNegVol
	'srfGVPosPlanarArea,srfGVNegPlanarArea,srfGVPosArea,srfGVNegArea,srfGVBlankedArea
 	'srfGVNumParams

	Wait 5
	AppActivate "Surfer"
	For Each shp In shapes1
		shp.Delete
	Next
	shapes1.AddText(1,1, _
		"GridVolume - see Scripter Immediate Window").Font.Size = 35

	'----------------------------------------------------------
	'GridData creates a GRD file from an XYZ DAT file.
	'NewVarioComponent illustrates how to add and modify
	' variogram components (models) for kriging.  They
	' are required when gridding with kriging using anisotropy.
	'----------------------------------------------------------
  Dim LinearComponent(1 To 1) As Object
  Set LinearComponent(1) = _
    surf.NewVarioComponent( _
    VarioType:=srfVarLinear, _
    Param1:=1, _
    Param2:=1, _
    AnisotropyRatio:=5, _
    AnisotropyAngle:=45) '<-- specify anisotropy.

	'Pass an Array of one Or more VarioComponents With KrigVariogram:= .

  surf.GridData(DataFile := path1+"demogrid.dat", _
    Algorithm := srfKriging, _
    ShowReport := False, _
    OutGrid := path1+"DemoAnisot.grd", _
    KrigVariogram := LinearComponent, _
    SearchEnable := True)

	Wait 1
	For Each shp In shapes1
		shp.Delete
	Next shp
	shapes1.AddText(1,1,"GridData with Anisotropy using NewVarioComponent").Font.Size = 30
	shapes1.AddContourMap(path1+"DemoAnisot.grd")
End Sub





	
