// ----------------------------------------------------------------------------
// Script4.js
//
// The is a simple JScript example to create a shaded relief map.
//
// To run this from the command line:
// cscript Script4.js
//
// This requires Windows Scripting Host, available free as part of
// Microsoft Internet Explorer 5.0, or from Microsoft's web site at:
// http://msdn.microsoft.com/scripting
//
// Note: As of 7-30-99, JScript does not support passing SAFEARRAY's to 
//       ActiveX objects.  This makes it unsuitable for several of the 
//       Surfer Automation methods and properties.  Microsoft has indicated 
//       in the newsgroups that this functionality is on the list for 
//       the next version of JScript (6.0).
//
// ----------------------------------------------------------------------------

// Create an instance of the Surfer Application object and assign it to the
// variable named "SurferObj"
var SurferObj = WScript.CreateObject("Surfer.Application");

// Makes the Surfer window visible
SurferObj.Visible = true;

// Creates a plot document in Surfer and assigns it to the variable named "DocObj"
var DocObj = SurferObj.Documents.Add();

// Assigns the location of the grid file to the variable named "Path"
var GridFile = SurferObj.Path + "\\samples\\Helens2.grd";
	
// Create the shaded relief map from Helens2.grd
DocObj.Shapes.AddReliefMap(GridFile);
