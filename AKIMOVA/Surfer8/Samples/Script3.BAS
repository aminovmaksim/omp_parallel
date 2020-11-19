'  Script3.bas creates several maps from a single data file.  The data
'  file is organized with columns corresponding to X, Y, Z1, Z2, etc.
'  A contour map is created for each set of data: X,Y,Z1, X,Y,Z2, etc.
'  The first row of the data file contains column headings.  This is
'  extracted and used to title the corresponding map.

Sub Main
	
	'Declares SurferApp, Wks, Doc, Plotwindow, Title, Map and
	'MapTitle as objects
	Dim SurferApp As Object
	Dim Wks As Object
	Dim Doc As Object
	Dim Plotwindow As Object
	Dim Map As Object
	Dim MapTitle As Object
	
	'Declares File as a string
	Dim File As String
	Dim Title As String

	'Declares retValue as a Boolean
	Dim retValue As Boolean
	
	'Declares MapNumber as an integer
	Dim MapNumber As Integer
	
	'Creates an instance of the Surfer Application object and assigns it to the
	'variable named "SurferApp"
	Set SurferApp = CreateObject("Surfer.Application")
	
	'Makes Surfer visible
	SurferApp.Visible = True
	
	'Assigns the name of the file to be used and its location to the
	'variable named "File"
	File = SurferApp.Path + "\Samples\Sample3"
	
	'Opens Sample3.dat in a new worksheet document and assigns it to
	'the variable named "Wks"
	Set Wks = SurferApp.Documents.Open(File+".dat")
	
	'Creates a plot document and assigns it to the variable named "Doc"
	Set Doc = SurferApp.Documents.Add()
	
	'Assigns the plot window to the variable named "Plotwindow"
	Set Plotwindow = Doc.Windows(1)
	
	'Disables Plotwindow's AutoRedraw to speed up the script
	Plotwindow.AutoRedraw = False
	
	'Assigns MapNumber to the value of zero and increments the MapNumber
	'For loop six times
	For MapNumber = 0 To 5 Step 1
		
		'Forms the map title by concatentating the column
		'titles in the first row of the worksheet
		Title = Wks.Cells(1,1) + _
			Wks.Cells(1,2) + _
			Wks.Cells (1,MapNumber + 3)	
		
		'Grids the specified data file using the Kriging algorithm and
		'assigns the return value to the variable named "retValue"
		retValue = SurferApp.GridData(DataFile:=File + ".dat", xCol:=1, _
			yCol:=2, zCol:=MapNumber +3, Algorithm:=srfKriging, _
			ShowReport:=False, OutGrid:=File + ".grd")
		
		'Creates a contour map and assigns the map coordinate system to
		'the variable named "Map"
		Set Map = Doc.Shapes.AddContourMap(File)
		
		'Sets the size of the map
		Map.Width = 3
		Map.Height = 3
		
		'Positions the map on the page
		map.Left = 0.5 + (MapNumber Mod 2) * (Map.Height + 1)
		map.Top = 10.5 - Int(MapNumber / 2) * (Map.Width + .5)
		
		'Positions the map title used below the map and assigns it to
		'the variable named "MapTitle"
		Set MapTitle = Doc.Shapes.AddText(Map.Left + Map.Width/2 - .5, _
			Map.Top - 3.1,"Plot of " + Title)

	'Returns to the beginning of the For loop
	Next
	
	'Enables AutoRedraw to update the plot window
	Plotwindow.AutoRedraw = True

	'Uncomment to send the page to the default printer to be printed
	'Doc.PrintOut

End Sub
