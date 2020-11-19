'VariogramObject.bas demonstrates the properties and methods
' of the VariogramObject.
Sub Main
	Debug.Print "----- VariogramObject.bas - ";Time;" -----"

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
	If plotdoc1.Type <> srfDocPlot Then _
		Set plotdoc1 = surf.Documents.Add

	'Caption changes when add new doc.
	surf.Caption = "Surfer "+surf.Version
	AppActivate "Surfer "+surf.Version

	With plotdoc1.PageSetup
		.Orientation = srfLandscape
		.Height = 8.5
		.Width = 11
	End With
	surf.Windows(1).Zoom(srfZoomPage)

	Set plotwin1 = surf.Windows(1)
	Set shapes1 = plotdoc1.Shapes

	path1 = surf.Path+"\samples\"

	shapes1.SelectAll
	plotdoc1.Selection.Delete

	'=======================
	'Create a new variogram.
	'=======================
	Debug.Print "AddVariogram"
	shapes1.AddText(1,0.75,"AddVariogram").Font.Size = 40
	Set vario1 = shapes1.AddVariogram(path1+"demogrid.dat")
	Wait 1

	'==========================
	'VariogramObject Properties
	'==========================
	'----------------------------------------------
	'The Axes Property returns the Axes Collection.
	'----------------------------------------------
	Debug.Print " Variogram Axes"
	shapes1("Text").Text = "Variogram Axes"
	For Each axs In vario1.Axes
		axs.AxisLine.ForeColor = srfColorBabyBlue
		Wait 1
	Next axs

	'----------------------------------------------
	'The EstimatorType Property returns or sets the
	' variogram estimatin method.  It returns a
	' srfVarioEstimator enumeration value.
	'-----------------------------------------------
	For i = srfVarioVariogram To srfVarioAutocorrelation
		vario1.EstimatorType = i
		'See varioestimatorname() function at end of script.
		Debug.Print " EstimatorType:";i;":";varioestimatorname(i)
		shapes1("Text").Text = "Estimatortype:"+Str(i)+":"+ _
			Str(varioestimatorname(i))
		Wait 1
	Next i

	vario1.EstimatorType = srfVarioVariogram
	'See varioestimatorname() function at end of script.
	Debug.Print " EstimatorType:"; "srfVarioVariogram"
	shapes1("Text").Text = "Estimatortype: srfVarioVariogram"
	Wait 1

	'-----------------------------------------------------------
	'The ExperimentalLine Property returns the properties of the
	' line connecting the points in the experimental variogram.
	' It returns a LineFormat Object.
	'-----------------------------------------------------------
	Debug.Print "ExperimentalLine"
	shapes1("Text").Text = "ExperimentalLine Color"
	vario1.ExperimentalLine.ForeColor = srfColorRed
	Wait 1

	'-----------------------------------------------------------------
	'The LagDirection Property returns or sets the variogram lag
	' direction angle in degrees.  It returns a double.
	' 0 degrees = +X direction, 90 degrees = +Y direction.
	' It is only meaningful with a LagTolerance less than 90 degrees.
	' Refer to the Experimental Tab of the Variogram Properties
	' help topic for more information.
	'-----------------------------------------------------------------
	Debug.Print "LagDirection:";vario1.LagDirection
	shapes1("Text").Text = "LagDirection:"+Str(vario1.LagDirection)
	Wait 1
	vario1.LagDirection = 60
	Debug.Print "LagDirection:";vario1.LagDirection
	shapes1("Text").Text = "LagDirection:"+Str(vario1.LagDirection)
	Wait 1

	'---------------------------------------------------------------
	'The LagTolerance Property returns or sets the variogram lag
	' tolerance in degrees.  It returns a double.
	'---------------------------------------------------------------
	Debug.Print "LagTolerance:"; vario1.LagTolerance
	shapes1("Text").Text = "LagTolerance:"+Str(vario1.LagTolerance)
	Wait 1
	vario1.LagTolerance = 30
	Debug.Print "LagTolerance:";vario1.LagTolerance
	shapes1("Text").Text = "LagTolerance:"+Str(vario1.LagTolerance)
	Wait 1

	'---------------------------------------------------------------
	'The LagWidth Property returns or sets the variogram lag
	' width in XY data units.  It returns a double.
	'---------------------------------------------------------------
	Debug.Print "LagWidth:"; vario1.LagWidth
	shapes1("Text").Text = "LagWidth:"+Str(vario1.LagWidth)
	Wait 1
	vario1.LagWidth = 0.1
	Debug.Print "LagWidth:";vario1.LagWidth
	shapes1("Text").Text = "LagWidth:"+Str(vario1.LagWidth)
	Wait 1

	'---------------------------------------------------------------
	'The MaxLagDistance Property returns or sets the variogram
	' maximum lag distance in XY data units.  It returns a double.
	'---------------------------------------------------------------
	Debug.Print "MaxLagDistance:"; vario1.MaxLagDistance
	shapes1("Text").Text = "MaxLagDistance:"+Str(vario1.MaxLagDistance)
	Wait 1
	vario1.MaxLagDistance = 4.0
	Debug.Print "MaxLagDistance:";vario1.MaxLagDistance
	shapes1("Text").Text = "MaxLagDistance:"+Str(vario1.MaxLagDistance)
	Wait 1

	'----------------------------------------------------------
	'The Model Property returns or sets the array of variogram
	' component objects.  It returns a variant array.
	'----------------------------------------------------------
	Debug.Print "Variogram Model"
	shapes1("Text").Text = "Variogram Model"
	Dim variocomponents() As Object
	variocomponents = vario1.Model
	'List current model components.
	For i = LBound(variocomponents) To UBound(variocomponents)
		Debug.Print " Component";i;" Type:"; variocomponents(i).Type; " "; _
			variocomponentname(variocomponents(i).Type) 'See functions below.
		Debug.Print "  Scale or Nugget Error Variance:"; _
			variocomponents(i).Param1
		Debug.Print "  Slope, Length, or Nugget Micro Variance:"; _
			variocomponents(i).Param2
		Debug.Print "  Anisotropy Ratio & Angle:"; _
			variocomponents(i).AnisotropyRatio; _
			variocomponents(i).AnisotropyAngle
	Next i

	'Change model components.
	Debug.Print "Change Variogram Model Components"
	shapes1("Text").Text = "Change Variogram Model Components"

	Set variocomponents(0) = surf.NewVarioComponent(srfVarNugget,10,0)
	Set variocomponents(1) = surf.NewVarioComponent(srfVarGaussian,250,1.5)
	vario1.Model = variocomponents

	Wait 1

	'---------------------------------------------------------
	'The ModelLine Property returns the properties of the line
	' in the variogram model.  It returns a LineFormat Object.
	'---------------------------------------------------------
	Debug.Print "ModelLine"
	shapes1("Text").Text = "ModelLine"
	vario1.ModelLine.ForeColor = srfColorGreen
	Wait 1

	'-----------------------------------------------------------
	'The NumLags Property returns or sets the number of lags in
	' the experimental variogram.  It returns an integer.
	'-----------------------------------------------------------
	Debug.Print "NumLags:"; vario1.NumLags
	shapes1("Text").Text = "NumLags"+Str(vario1.NumLags)
	Wait 1
	vario1.NumLags = 30
	Debug.Print "NumLags:"; vario1.NumLags
	shapes1("Text").Text = "NumLags"+Str(vario1.NumLags)
	Wait 1

	'----------------------------------------------------------
	'The ShowPairs Property returns or sets the display of the
	' number of pairs used for each point in the experimental
	' variogram.  It returns a Boolean.
	'----------------------------------------------------------
	Debug.Print "ShowPairs"
	shapes1("Text").Text = "Show Number of Pairs?"+vario1.ShowPairs
	Wait 1
	vario1.ShowPairs = True
	shapes1("Text").Text = "Show Number of Pairs?"+vario1.ShowPairs
	Wait 1

	'----------------------------------------------------------
	'The PairsFont Property returns the font properties for the
	' number of pairs displayed by each point in the
	' experimental variogram.  It returns a FontFormat Object.
	'----------------------------------------------------------
	Debug.Print "PairsFont";vario1.PairsFont.Face
	shapes1("Text").Text = "PairsFont:"+ vario1.PairsFont.Face
	Wait 1
	vario1.PairsFont.Face = "Times New Roman"
	Debug.Print "PairsFont";vario1.PairsFont.Face
	shapes1("Text").Text = "PairsFont:"+ vario1.PairsFont.Face

	'-----------------------------------------------------------------
	'The ShowExperimental Property returns or sets the display of the
	' line in the experimental variogram.  It returns a Boolean.
	'-----------------------------------------------------------------
	Debug.Print "ShowExperimental? ";vario1.ShowExperimental
	shapes1("Text").Text = "ShowExperimental? "+ vario1.ShowExperimental
	Wait 1
	vario1.ShowExperimental = False 'Turn off line.
	Debug.Print "ShowExperimental? ";vario1.ShowExperimental
	shapes1("Text").Text = "ShowExperimental? "+ vario1.ShowExperimental
	Wait 1
	vario1.ShowExperimental = True 'Turn line on.

	'---------------------------------------------------------------
	'The ShowModel Property returns or sets the display of the model
	' line in the variogram.  It returns a Boolean.
	'---------------------------------------------------------------
	Debug.Print "ShowModel? ";vario1.ShowModel
	shapes1("Text").Text = "ShowModel? "+ vario1.ShowModel
	Wait 1
	vario1.ShowModel = False 'Turn off line.
	Debug.Print "ShowModel? ";vario1.ShowModel
	shapes1("Text").Text = "ShowModel? "+ vario1.ShowModel
	Wait 1
	vario1.ShowModel = True 'Turn line on.

	'--------------------------------------------------------------
	'The TitleFont returns the font properties of the variogram
	' subtitle, which displays the angle and tolerance.  It
	' returns a FontFormat Object.
	'--------------------------------------------------------------
	Debug.Print "TitleFont Face: ";vario1.TitleFont.Face
	shapes1("Text").Text = "TitleFont Face: "+vario1.TitleFont.Face
	Wait 1
	vario1.TitleFont.Size = 25
	Debug.Print "TitleFont Size: ";vario1.TitleFont.Size
	shapes1("Text").Text = "TitleFont Size: "+vario1.TitleFont.Size
	Wait 1

	'--------------------------------------------------------------
	'The SubTitleFont returns the font properties of the variogram
	' subtitle, which displays the angle and tolerance.  It
	' returns a FontFormat Object.
	'--------------------------------------------------------------
	Debug.Print "SubTitleFont Face: ";vario1.SubTitleFont.Face
	shapes1("Text").Text = "SubTitleFont Face: "+vario1.SubTitleFont.Face
	Wait 1
	vario1.SubTitleFont.Size = 25
	Debug.Print "SubTitleFont Size: ";vario1.SubTitleFont.Size
	shapes1("Text").Text = "SubTitleFont Size: "+vario1.SubTitleFont.Size
	Wait 1

	'----------------------------------------------------------------
	'The ShowSubTitle Property returns or sets the display of the
	' variogram subtitle, which displays the direction and tolerance.
	' It returns a Boolean.
	'----------------------------------------------------------------
	Debug.Print "ShowSubtitle? ";vario1.ShowSubTitle
	shapes1("Text").Text = "ShowSubtitle? "+ vario1.ShowSubTitle
	Wait 1

	vario1.ShowSubTitle = False 'Turn off subtitle.
	Debug.Print "ShowSubtitle? ";vario1.ShowSubTitle
	shapes1("Text").Text = "ShowSubtitle? "+ vario1.ShowSubTitle
	Wait 1
	vario1.ShowSubTitle = True 'Turn subtitle on.

	'--------------------------------------------------------------
	'The ShowSymbols Property returns or sets the display of the
	' symbols in the experimental variogram.  It returns a Boolean.
	' If ShowPairs is True, they remain visible even when
	' ShowSymbols is False.
	'--------------------------------------------------------------
	Debug.Print "ShowSymbols? ";vario1.ShowSymbols
	shapes1("Text").Text = "ShowSymbols? "+vario1.ShowSymbols
	Wait 1
	vario1.ShowSymbols = False  'Turn symbols off
	Debug.Print "ShowSymbols? ";vario1.ShowSymbols
	shapes1("Text").Text = "ShowSymbols? "+vario1.ShowSymbols
	Wait 1
	vario1.ShowSymbols = True  'Turn symbols on.

	'-------------------------------------------------------------
	'The Symbol Property returns a MarkerFormat Object containing
	' the properties of the symbols on the experiemental
	' variogram.
	'-------------------------------------------------------------
	Debug.Print "Symbol Index: ";vario1.Symbol.Index
	shapes1("Text").Text = "Symbol Index: "+vario1.Symbol.Index
	Wait 1
	vario1.Symbol.Index = 0
	Debug.Print "Symbol Index: ";vario1.Symbol.Index
	shapes1("Text").Text = "Symbol Index: "+vario1.Symbol.Index
	Wait 1

	'-------------------------------------------------------------
	'The ShowVariance Property returns or sets the display of the
	' variance line, a dashed horizontal line by default.  It
	' returns a Boolean.
	'-------------------------------------------------------------
	Debug.Print "ShowVariance? ";vario1.ShowVariance
	shapes1("Text").Text = "ShowSymbols? "+vario1.ShowVariance
	Wait 1
	vario1.ShowVariance = True  'Turn variance line on.
	Debug.Print "ShowVariance? ";vario1.ShowVariance
	shapes1("Text").Text = "ShowVariance? "+vario1.ShowVariance
	Wait 1

	'-----------------------------------------------------------
	'The VarianceLine Property returns a LineFormat Object with
	' the properties of the variance line.
	'-----------------------------------------------------------
	Debug.Print "VarianceLine Style: ";vario1.VarianceLine.Style
	shapes1("Text").Text = "VarianceLine Style: " + _
		vario1.VarianceLine.Style
	Wait 1
	vario1.VarianceLine.Style = "Solid"
	Debug.Print "VarianceLine Style: ";vario1.VarianceLine.Style
	shapes1("Text").Text = "VarianceLine Style: " + _
		vario1.VarianceLine.Style
	Wait 1

	'--------------------------------------------------------------
	'The VerticalScale Property returns or sets the vertical scale
	' of the variogram.  It returns a double.  Experimental tab of
	' Variogram Properties dialog box.
	'--------------------------------------------------------------
	Debug.Print "VerticalScale:"; vario1.VerticalScale
	shapes1("Text").Text = "VerticalScale:"+Str(vario1.VerticalScale)
	Wait 1
	vario1.VerticalScale = vario1.VerticalScale*2
	Debug.Print "VerticalScale:"; vario1.VerticalScale
	shapes1("Text").Text = "VerticalScale:"+Str(vario1.VerticalScale)
	Wait 1
	vario1.VerticalScale = vario1.VerticalScale/2

	'========================
	'VariogramObject Methods
	'========================
	'The AutoFit Method fits the variogram parameters
	' to the current model.
	'------------------------------------------------
	Debug.Print "AutoFit"
	shapes1("Text").Text = "AutoFit"
	variocomponents = vario1.Model
	'List current model components.
	For i = LBound(variocomponents) To UBound(variocomponents)
		Debug.Print " Component";i;" Type:"; variocomponents(i).Type; " "; _
			variocomponentname(variocomponents(i).Type) 'See functions below.
		Debug.Print "  Scale or Nugget Error Variance:"; _
			variocomponents(i).Param1
		Debug.Print "  Slope, Length, or Nugget Micro Variance:"; _
			variocomponents(i).Param2
		Debug.Print "  Anisotropy Ratio & Angle:"; _
			variocomponents(i).AnisotropyRatio; _
			variocomponents(i).AnisotropyAngle
	Next i
	Wait 1
	vario1.AutoFit
	variocomponents = vario1.Model
	Debug.Print "AutoFit values"
	For i = LBound(variocomponents) To UBound(variocomponents)
		Debug.Print " Component";i;" Type:"; variocomponents(i).Type; " "; _
			variocomponentname(variocomponents(i).Type) 'See functions below.
		Debug.Print "  Scale or Nugget Error Variance:"; _
			variocomponents(i).Param1
		Debug.Print "  Slope, Length, or Nugget Micro Variance:"; _
			variocomponents(i).Param2
		Debug.Print "  Anisotropy Ratio & Angle:"; _
			variocomponents(i).AnisotropyRatio; _
			variocomponents(i).AnisotropyAngle
	Next i

	'-----------------------------------------------------------------
	'The Export Method exports the experimental variogram data points
	' to an ASCII DAT file.
	' The X coordinates (spearation distance) are in column A,
	' the Y coordinates (variogram g(h) are in column B,
	' the number of pairs are in column C.
	' This method returns a Boolean.
	'-----------------------------------------------------------------
	Debug.Print "Export to vario.dat"
	shapes1("Text").Text = "Export to vario.dat"
	vario1.Export(path1+"vario.dat")

End Sub
Function varioestimatorname(x)
	ven = Array("0", _
		"srfVarioVariogram", _
		"srfVarioStdVariogram", _
		"srfVarioAutocovariance", _
		"srfVarioAutocorrelation")
	varioestimatorname = ven(x)
End Function
Function variocomponentname(x)
	vcn = Array("0", _
		"srfVarExponential", _
		"srfVarGaussian", _
		"srfVarLinear", _
		"srfVarLogarithmic", _
		"srfVarNugget", _
		"srfVarPower", _
		"srfVarQuadratic", _
		"srfVarRationalQuadratic", _
		"srfVarSpherical", _
		"srfVarWave")
	variocomponentname = vcn(x)
End Function
