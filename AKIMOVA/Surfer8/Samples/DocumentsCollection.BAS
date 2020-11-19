'DocumentsCollection.bas illustrates the properties and
' methods of the Surfer Documents Collection. - TB - 25 Oct 01
Sub Main
	Debug.Print "----- DocumentsCollecton.bas - ";Time;" -----"

	'Get existing Surfer instance, or create a new one If none exists.
	On Error Resume Next 'Turn off error reporting.
	Set surf = GetObject(,"Surfer.Application")
	If Err.Number<>0 Or surf.Documents.Count = 0 Then
		Set surf = CreateObject("Surfer.Application")
		surf.Documents.Add(srfDocPlot)
	End If
	On Error GoTo 0 'Turn on error reporting.

	Debug.Print "Surfer ";surf.Version
	surf.Visible = True
	surf.WindowState = srfWindowStateNormal
	surf.Width = 600
	surf.Height = 400
	surf.Windows(1).Zoom(srfZoomPage)

	Set plotdoc1 = surf.Documents(1)
  Set plotwin1 = surf.Windows(1)

	path1 = surf.Path+"\samples\"
	file1 = path1+"demogrid"

	Set shapes1 = plotdoc1.Shapes

	'-------------------------------------------------------
	' Documents Collection Properties.
	'-------------------------------------------------------
	Debug.Print "-- Documents Collection Properties --"

	Set docs1 = surf.Documents
	Debug.Print "docs1.Application = ";docs1.Application
	Debug.Print "docs1.Count = ";docs1.Count
	Debug.Print "docs1.Parent = ";docs1.Parent

	'-------------------------------------------------------
	' Documents Collection Methods.
	'-------------------------------------------------------
	Debug.Print "-- Documents Collection Methods --"
	docs1.Add (srfDocPlot) 'Add a new plot document.
	Debug.Print "Add"

	AppActivate "Surfer"
	'Close all docs. Ask to save if changes have been made.
	Debug.Print "CloseAll"
	docs1.CloseAll(SaveChanges:=srfSaveChangesAsk)

	'Use Item to refernce a particular document.
	Debug.Print "Item"
	docs1.Add (srfDocPlot) 'Add a new plot document.
	docs1.Add (srfDocWks)  'Add a new worksheet document.
	Set plotdoc1 = docs1.Item(1)
	'Item is the default method, so the above statement can be written
	' with shorter syntax.
	Set plotdoc1 = docs1(1)
	'The document name can be used as the Item index.
	' Set plotdoc1 = docs1("Plot1")
	surf.Windows.Arrange(srfTileVert)
	Wait 3
	docs1.CloseAll(srfSaveChangesAsk)

	'Open a file in a document.
	Debug.Print "Open"
	Set plotdoc2 = docs1.Open(path1+"sample2.srf") 'A Plot document.
	Set wksdoc1 = docs1.Open(path1+"demogrid.dat") 'A Worksheet document.
	Set griddoc1 = docs1.Open(path1+"demogrid.grd") 'A Grid document.
	surf.Windows.Arrange(srfTileVert)
	Wait 3

	'Prompt to save changes to all documents, and leave them open.
	Debug.Print "SaveAll"
	docs1.SaveAll(True)
	'Close all docs and prompt to save changes.
	docs1.CloseAll(srfSaveChangesAsk)

End Sub
