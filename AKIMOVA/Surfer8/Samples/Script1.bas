Attribute VB_Name = "Module2"
'  Script1.bas creates several maps from a single data file.  A contour map is
'  created from each of the nine gridding methods.  Each map is then labeled with
'  its corresponding gridding method.

Sub Main
        
	'Declares SurferApp, Doc, Plotwindow, Map, MapTitle1 and MapTitle2 as objects
	Dim SurferApp As Object
	Dim Doc As Object
	Dim Plotwindow As Object
	Dim Map As Object
	Dim MapTitle1 As Object
	Dim MapTitle2 As Object
	
	'Declares MethodLabel as a variant
	Dim MethodLabel As Variant
	
	'Declares retValue as a Boolean
	Dim retValue As Boolean
	
	'Declares Data, Grid and Path as strings
	Dim Data As String
	Dim Grid As String
	Dim Path As String
	
	'Declares Method as an integer
	Dim Method As Integer
	
	'Assigns the name of the data file to be used to the variable named "Data"
	Data = "Demogrid.dat"
	
	'Assigns the name of the grid file to be used to the variable named "Grid"
	Grid = "Demogrid.grd"
        
	'Assigns the array of all nine gridding algorithms to the variable named
	'"MethodLabel"
	MethodLabel = Array("Inverse Distance", "Kriging", "Minimum Curvature", _
		"Modified Shepard's Method", "Natural Neighbor", "Nearest Neighbor", _
		"Polynomial Regression", "Radial Basis Functions", _
		"Triangulation with Linear Interpolation")

	'Creates an instance of the Surfer Application object and assigns it to the
	'variable named "SurferApp"
	Set SurferApp = CreateObject("Surfer.Application")

	'Assigns the location of the data and grid files to the variable named "Path"
	Path = SurferApp.Path + "\samples\"

	'Makes Surfer visible
	SurferApp.Visible = True

    'Creates a plot document in Surfer and assigns it to the variable named "Doc"
	Set Doc = SurferApp.Documents.Add
	
	'Assigns the plot window to the variable named "Plotwindow"
	Set Plotwindow = Doc.Windows(1)
	
	'Disables Plotwindow's AutoRedraw to speed up the script
	Plotwindow.AutoRedraw = False
	
	'Increments the Method For loop nine times
	For Method = 1 To 9 Step 1
    
		'Grids the specified data file using the current gridding method
		'specified in the loop
		retValue = SurferApp.GridData(DataFile:=Path + Data, Algorithm:=Method, _
			ShowReport:=False, OutGrid:=Path + Grid)
                
		'Creates a contour map from the gridded data and assigns it to the variable
		'named "Map"
		Set Map = Doc.Shapes.AddContourMap(Path + Grid)

		'Sets the size of the map
		Map.Width = 2
		Map.Height = 2
		
		'Positions the map on the page
		Map.Left = 0.5 + ((Method - 1) Mod 3) * (Map.Height + .5)
		Map.Top = 10.5 - Int((Method - 1) / 3) * (Map.Width + .75)
		
		'Positions the first line of the map title below the map and assigns it to
		'the variable named "MapTitle1"
		Set MapTitle1 = Doc.Shapes.AddText(Map.Left + Map.Width/2 - 1, _
			Map.Top - 2.07, Data + " gridded using ")
		
		'Positions the second line of the map title below the map and assigns it to
		'the variable named "MapTitle2"
		Set MapTitle2 = Doc.Shapes.AddText(Map.Left + Map.Width/2 - 1, _
			Map.Top - 2.22, MethodLabel(Method - 1))

	'Returns to the beginning of the For loop
	Next
	
	'Enables AutoRedraw to update the plot window
	Plotwindow.AutoRedraw = True
	
	'Uncomment to send the page to the default printer to be printed
	'Doc.PrintOut

End Sub

