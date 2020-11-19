'PlotDocumentObject.bas demonstrates the properties and
' methods of the Plot Document Object.
Sub Main
	Debug.Print "----- PlotDocumentObject.bas - ";Time;" -----"

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
 	Debug.Print "PlotDocumentObject Properties and Methods"

	Set plotdoc1 = surf.Documents(1)
  Set plotwin1 = surf.Windows(1)
	Set shapes1 = plotdoc1.Shapes

	path1 = surf.Path+"\samples\"
	file1 = path1+"demogrid"

	AppActivate "PlotDocumentObject"

	'===================================================================
	'Plot Document Properties
	'-------------------------------------------------------------------
	'DefaultFill returns a FillFormat object.
	'-------------------------------------------------------------------
	'Color names in Color dialog and attrib.ini.
	Debug.Print "Default Fill Color = ";plotdoc1.DefaultFill.ForeColor; _
		". (srfColorBlack = 0)";
	plotdoc1.DefaultFill.ForeColor = srfColorRed 'Current session only.
	Debug.Print ". New Color = ";plotdoc1.DefaultFill.ForeColor; _
		". (srfColorRed = 255)"

	'Set fill pattern to Solid or Fill Color will not display.
	Debug.Print "Default Fill Pattern = ";plotdoc1.DefaultFill.Pattern;
	plotdoc1.DefaultFill.Pattern = "Solid" 'Current session only.
	Debug.Print ". New Pattern = ";plotdoc1.DefaultFill.Pattern

	'-------------------------------------------------------------------
	'DefaultFont returns a FontFormat object.
	'-------------------------------------------------------------------
	Debug.Print "Default Font Properties"
	'"With" fills in "plotdoc1.DefaultFont" before the "."
	With plotdoc1.DefaultFont
		Debug.Print "  Bold = "; .Bold
		'Color names in Color dialog and attrib.ini.
		Debug.Print "  Color = "; .Color
		Debug.Print "  Face = "; .Face
		Debug.Print "  Horizontal Alignment = ";.HAlign;" (srfTALeft = 1) see SrfHTextAlign"
		Debug.Print "  Italic = "; .Italic
		Debug.Print "  Size = "; .Size
		Debug.Print "  StrikeThrough = "; .StrikeThrough
		Debug.Print "  Underline = "; .Underline
		Debug.Print "  Vertical Alignment = "; .VAlign; " (srfTATop = 1) see SrfVTextAlign"
	End With

	'-------------------------------------------------------------------
	'DefaultLine returns a LineFormat object.
	'-------------------------------------------------------------------
	Debug.Print "Default Line Properties"
	With plotdoc1.DefaultLine
		'Color names in Color dialog and attrib.ini.
		Debug.Print "  ForeColor = "; .ForeColor
		'Style names in Line Properties dialog and attrib.ini.
		Debug.Print "  Style = "; .Style;
		'Width units are page units (surf.PageUnits, srfPageUnits enums.
		Debug.Print "  Width = "; .Width
	End With

	'-------------------------------------------------------------------
	'DefaultSymbol returns a MarkerFormat object.
	'-------------------------------------------------------------------
	Debug.Print "Default Symbol Properties"
	With plotdoc1.DefaultSymbol
		'Color names in Color dialog and attrib.ini.
		Debug.Print "  Color = "; .Color
		'Index in Symbol Properties dialog and symbols.srf.
		Debug.Print "  Index = "; .Index
		'Set is any TrueType Font.
		Debug.Print "  Set = "; .Set
		'Size is in surf.PageUnits, srfPageUnits enums.
		Debug.Print "  Size = "; .Size
	End With

	'-------------------------------------------------------------------
	'PageSetup property returns a PageSetup object.
	'-------------------------------------------------------------------
	Debug.Print "Page Setup Properties"
	With plotdoc1.PageSetup
		Debug.Print "  BottomMargin = "; .BottomMargin
		Debug.Print "  Height = "; .Height
		Debug.Print "  LeftMargin = "; .LeftMargin
		Debug.Print "  Orientation = "; .Orientation; ", Portrait = 1"
		Debug.Print "  RightMargin = "; .RightMargin
		Debug.Print "  TopMargin = "; .TopMargin
		Debug.Print "  Width = "; .Width
	End With

	'-------------------------------------------------------------------
	'Selection returns the currently selected objects
	' in a Selection Collection.  See SelectionCollection.bas for more.
	'-------------------------------------------------------------------
	Debug.Print " Number of Selected Items = ";plotdoc1.Selection.Count

	'-------------------------------------------------------------------
	'Shapes returns the currently selected objects
	' in a Shapes Collection.  See ShapesCollection.bas for more.
	'-------------------------------------------------------------------
	Debug.Print " Number of Shapes = ";plotdoc1.Shapes.Count

	'-------------------------------------------------------------------
	'ShowObjectManager turns the ObjectManager on or off.
	' It returns a Boolean.
	'-------------------------------------------------------------------
	Debug.Print " ShowObjectManager = ";plotdoc1.ShowObjectManager
	'Display Object Manager.
	plotdoc1.ShowObjectManager = True

	'===================================================================
	'Plot Document Methods
	'-------------------------------------------------------------------
	'The Export method exports selected or all objects in the plot
	' document to a file.
	'-------------------------------------------------------------------
	AppActivate "Surfer"
	Set plotdoc2 = surf.Documents.Open(path1+"sample6.srf")
	plotdoc2.Export(path1+"sample6.tif", _
		Options:="Width=717, Height=949")  'Options string in double-quotes.
	Wait 3
	plotdoc2.Windows("sample6.srf").Close

	'-------------------------------------------------------------------
	'The Import method imports a file as a graphic (no map coordinates).
	' It returns a Shape Object.
	'-------------------------------------------------------------------
	plotdoc1.Import(path1+"sample6.tif", _
		Options:="Defaults=1")
	shapes1.AddText(1,1,"Import TIF Image").Font.Size = 40

	'-------------------------------------------------------------------
	'The PrintOut method
	'-------------------------------------------------------------------
	plotdoc1.PrintOut(SelectionOnly:=True)



End Sub
