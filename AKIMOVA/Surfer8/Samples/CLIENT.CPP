//-----------------------------------------------------------------------------
// 
// Client.cpp    Copyright (C) Golden Software, Inc. 1999-2000
//
// You have a royalty-free right to use, modify, reproduce and distribute
// the Sample Application Files in any way you find useful, provided that 
// you agree that Golden Software has no warranty, obligations or liability
// for any Sample Application Files.
//
// This example demonstrates how to automate Surfer from C++.  A new instance
// of Surfer is started, and a shaded relief map is created.  The shaded relief
// map colors are then "animated" from black to red.
//
// This example uses the Microsoft-specific extension "#import" to create
// header files from the type library information contained within Surfer.exe.
// The generated header files contain several smart-pointer classes that
// greatly simplify this entire process.  Make sure the #import directive
// below points to the correct location of Surfer.exe.
//
// This example was compiled with Microsoft Visual C++ 6.0 using the
// following command line:
//
// cl -GX Client.cpp
// 
//-----------------------------------------------------------------------------
#include <iostream>
#include <windows.h>
#import "..\surfer.exe" no_namespace

using namespace std;

void AnimateColorMap(IColorMapPtr& pCM);  // Forward declaration

//-----------------------------------------------------------------------------
void main()
{
   // Initialize the system DLL's
   HRESULT hr = CoInitialize(NULL);
   if(FAILED(hr))
   {
      cout << "CoInitialize Failed: " << hr << "\n\n";
      exit(1);
   }

   try
   {
      // Create an instance of the Application object
      IApplicationPtr pApp(__uuidof(Application));

      // Make the Application object visible
      pApp->Visible = VARIANT_TRUE;

      // Get the Documents collection
      IDocumentsPtr pDocs = pApp->Documents;

      // Add a new document to the Documents collection.
      IPlotDocumentPtr pDoc = pDocs->Add(srfDocPlot);

      // Get the Shapes collection from the document
      IShapesPtr pShapes = pDoc->Shapes;

      // Create a shaded relief map from the helens2.grd file.  This is 
      // in a subdirectory off the application startup directory.
      IMapFramePtr pMF = pShapes->AddReliefMap(pApp->Path + "\\Samples\\Helens2.grd");

      // Get the shaded relief overlay from the map frame
      IReliefMapPtr pRM = pMF->Overlays->Item(1L);

      // Get the ColorMap used by the relief map
      IColorMapPtr pCM = pRM->ColorMap;

      // Animate the colors of the shaded relief map from the default black 
      // to red.
      AnimateColorMap(pCM);
   }
   catch(_com_error e)
   {
      MessageBeep(0);

      // Display a generic COM error message
      cout << e.ErrorMessage() << "\n";

      // Display any application supplied information
      if(e.Description().length())
         cout << (const char *)e.Description() << "\n";
   }

   CoUninitialize();
}

//-----------------------------------------------------------------------------
// This "animates" the base color of the shaded relief map from black (0) to 
// pure red (255).  The map using the colormap will be redrawn after each call 
// to SetNodes.

void AnimateColorMap(IColorMapPtr& pCM)
{
   // Get the current node positions and colors from the color map
   SAFEARRAY *pPositions = pCM->NodePositions;
   SAFEARRAY *pColors    = pCM->NodeColors;

   // Loop to set the base color of the shaded relief map

   for(int iRed=0; iRed<=255; iRed+=51)
   {
      // Arrays returned by Surfer are always zero-based
      long Index=0;

      // Set the color of node 0 (the base color)
      COLORREF Color = RGB(iRed,0,0);
      SafeArrayPutElement(pColors,&Index,&Color);

      // Update the colormap and the map using it
      pCM->SetNodes(pPositions,pColors);

      // Wait a second so the new map can be seen
      Sleep(1000);
   }

   SafeArrayDestroy(pPositions);
   SafeArrayDestroy(pColors);
}
