Attribute VB_Name = "Module2"
'  Script2.bas demonstrates a rudimentary template script.  A data
'  file is specified at the start of the script.  A contour map and
'  wireframe map are then generated.  This can be applied to any data
'  file by changing the value of the variable "InFile" appropriately.

Sub Main

	'Declares SurferApp, Doc, Plotwindow, ContourMapFrame, ContourMap,
	'WireframeMapFrame and Wireframe as objects
	Dim SurferApp As Object
	Dim Doc As Object
	Dim Plotwindow As Object
	Dim ContourMapFrame As Object
	Dim ContourMap As Object
	Dim WireframeMapFrame As Object
	Dim Wireframe As Object
	
	'Declares InFile and OutFile as strings
	Dim InFile As String
	Dim OutFile As String
	
	'Declares retValue as a Boolean
	Dim retValue As Boolean
	
	'Creates an instance of the Surfer Application object and assigns it to the
	'variable named "SurferApp"
	Set SurferApp = CreateObject("Surfer.Application")
	
	'Assigns the name of the data file to be used and its location to the
	'variable named "InFile"
	InFile = SurferApp.Path + "\Samples\" + "demogrid.dat"
	
	'Assigns the name of the grid file to be used and its location to the
	'variable named "OutFile"
	OutFile = SurferApp.Path + "\Samples\" + "Sample"
	
	'Makes Surfer visible
	SurferApp.Visible = True
	
	'Grids the specified data file using the Kriging algorithm and assigns
	'the return value to the variable named "retValue"
	retValue = SurferApp.GridData(DataFile:=InFile, Algorithm:=srfKriging, _
		ShowReport:=False, OutGrid:=OutFile)
	
	'Creates a plot document in Surfer and assigns it to the variable named "Doc"
	Set Doc = SurferApp.Documents.Add(srfDocPlot)
	
	'Assigns the new plot window to the variable named "PlotWindow"
	Set Plotwindow = Doc.Windows(1)
	
	'Disables Plotwindow's AutoRedraw to speed up the script
	Plotwindow.AutoRedraw = False
	
	'Creates a contour map and assigns the map coordinate system to the variable
	'named "ContourMapFrame"
	Set ContourMapFrame = Doc.Shapes.AddContourMap(OutFile)

	'Assigns the contour map attributes to the variable named "ContourMap"
	Set ContourMap = ContourMapFrame.Overlays(1)

	'Displays the fill of the contour map
	ContourMap.FillContours = True
        
	'Positions the contour map to the lower half of the page
	ContourMapFrame.Left = 0.75
	ContourMapFrame.Top = 5.5
	
	'Sets the size of the contour map
	ContourMapFrame.Width = 7
	ContourMapFrame.Height = 5
	
	'Creates a wireframe map and assigns the map coordinate system to the
	'variable named "WireframeMapFrame"
	Set WireframeMapFrame = Doc.Shapes.AddWireframe(OutFile)

	'Assigns the wireframe map attributes to the variable named "Wireframe"
	Set Wireframe = WireframeMapFrame.Overlays(1)

	'Changes the XY mesh color to blue
	Wireframe.xLine.ForeColor = RGB(0, 0, 255)
	Wireframe.yLine.ForeColor = RGB(0, 0, 255)
        
	'Positions the wireframe map to the upper half of the page
	WireframeMapFrame.Left = 0.75
	WireframeMapFrame.Top = 10.75
	
	'Sets the size of the wireframe map
	WireframeMapFrame.Width = 7
	WireframeMapFrame.Height = 5
	
	'Enables AutoRedraw to update the plot window
	Plotwindow.AutoRedraw = True
	
	'Uncomment to send the page to the default printer to be printed
	'Doc.PrintOut

End Sub
