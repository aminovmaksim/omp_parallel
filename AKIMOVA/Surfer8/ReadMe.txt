-------------------------------------------------------------------------------
Surfer(R) Version 8.0 ReadMe File

Copyright (c) 2002 Golden Software, Inc.
All rights reserved.
-------------------------------------------------------------------------------


System Requirements
===================
- Windows 98, Me, 2000, XP or higher.
- Windows 95 and NT are not specifically disallowed, but are no longer 
  supported.
- 800 X 600 or higher monitor resolution with support for at least 256 colors
- 20 MB of free hard disk space in the Windows drive.
- 13 MB of free hard disk space.
- At least 8 MB RAM above the Windows requirement for simple data sets.
  At least 128 MB of RAM above the Windows requirement for large bitmaps
  and surface maps recommended.


New Features in Surfer 8.0
==========================
- Color rendered surface maps.
- Overlay raster and vector maps on surface maps.
- Grid Mosaic to combine adjacent grids into a single grid file.
- Cross-validation to assess the quality of a gridding method.
- New Filters and smoothing methods, including user-defined filters.
- Map Delaunay Triangles used by Triangulation and Natural Neighbor 
  gridding methods.
- Use up to 1 billion data points in the worksheet and when gridding, 
  subject to memory and Windows version requirements.
- Retain faults in blanked grids.
- Display bitmap and metafile information.
- Generate data and grid statistics reports.
- Rotate and tilt raster maps (bitmap base maps, image maps, shaded relief
  maps).
- SRF files are compressed with new bitmap compression routines.
- New gridding methods: Moving Average, Local Polynomial.
- Data Metrics to calculate spatial data statistics.
- New variogram models: Cubic, Pentaspherical.
- Specify Z scale factor in volume calculations when Z units differ from
  XY units.
- Faster loading for large data files by postponing sorting and duplicate 
  checking until gridding begins.
- New import formats: Enhanced Metafile EMF, Golden Software Interchange GSI,
  ESRI Arc/INFO Export E00.
- New export formats: Enhanced Metafile EMF, Golden Software Interchange GSI,
  MapInfo Interchange Format MIF, Golden Software Boundary GSB.
- ANOVA statistics in gridding and variogram reports.
- Trackball tool for easier rotating, tilting, and perspective changes.
- Pan tool for dragging the view without using the scroll bars.
- Zoom Realtime tool.
- Elimination of the 32 inch page size in Windows 98, Me.
- Check for Update to download program updates.
- Modeless object properties dialog box stays open when you select 
  different objects.


Installing Surfer
=================
Golden Software does not recommend installing Surfer 8 over any previous 
version of Surfer (e.g., Surfer 7) or in the same directory as an older
version.  Surfer 8 can co-exist with older versions as long as they are 
in different directories.

Prior to installing Surfer 8, uninstall any previous versions of Surfer 8
using the Control Panel | Add/Remove Programs menu.  Select Surfer 8 from 
the list under the Install/Uninstall tab.  Click the Add/Remove button to 
uninstall Surfer.  Versions prior to Surfer 7 will not appear in the Add/Remove
Programs list.

To install Surfer 8 under Windows 2000 and XP, you need to have permissions 
as a local Administrator for that machine.


Network Installation
====================
An administrative install may be performed in order to install files to a 
network server.  Once installed on the network server, individual workstation
installations can be performed.  The server software must support long 
filenames.  

   *** ATTENTION: You may use Surfer on a networked system provided that the 
   *** number of Surfer users on the network at one time does not exceed the 
   *** number of licensed copies of Surfer.

To install Surfer on the server:
1. Log on to the file server with administrator rights.  
2. Click Start | Run.
3. Enter the path to Setup.exe followed by /a.  For example, 
   r:\Setup.exe /a
4. When Setup asks for a destination folder, choose one on the file server 
   (e.g. C:\SurferServer).  This should be a new empty directory.  Setup copies 
   all the Surfer files plus the setup program and its associated files to the 
   server drive. 
5. We recommend flagging the server folder contents as read-only.

NOTE: If you wish to run Surfer on a Windows server itself, you will need to 
      run setup again without the /a switch and install Surfer normally to a
      new directory.

Next, install Surfer at each of the workstations:
1. Log on to the file server from each workstation that will run Surfer.
2. Start Windows and run the copy of Setup.exe located in the server folder.  
   DO NOT RUN SETUP FROM THE CD-ROM!
3. When the setup asks for a destination folder, choose one on the local hard 
   drive.

You will need to enter the Surfer serial number the first time Surfer is run 
on each workstation.

Surfer must be installed directly at each workstation.  It cannot be installed 
from a remote computer over the network because the installation involves 
editing the computer's registry.  Certain parts of the registry cannot be 
edited from a remote computer.


Installation Troubleshooting 
============================
If you are experiencing errors with the installation, install Surfer and
create an installation log:

1. Go to Start | Run.

2. For a local installation, type 
   "[drive letter]:\[path]\setup.exe" /V"/L* [drive letter]:\[path]\S8.log"
  
   For a network installation, type 
   "[drive letter]:\[path]\setup.exe" /a /V"/L* [drive letter]:\[path]\S8.log"

   Replace the [drive letter] and [path] with a drive letter and path 
   on your computer.  The path to the log file cannot contain spaces.

3. Click OK.  After the installation process terminates, the log file is 
   created in the path you specify.  Please send the log file to technical 
   support along with the installation message text.


Sample Files
============
Several sample files are included in the Samples directory.  The samples
include interesting maps, data files, and some scripts that can be used 
to automate Surfer.

Script1.bas     Script to create a contour map for each gridding method
Script2.bas	    Script to create a contour map and wireframe map
Script3.bas     Script to create contour maps from multiple Z data columns
Client.cpp      Example C++ program to Automate Surfer
Script4.js      Creates a relief map using Windows Scripting Host and JScript 
Script5.vbs     Creates a relief map using Windows Scripting Host and VBScript 
Blusteel.clr    Color Spectrum File - black, blue, gray
Bywaves.clr     Color Spectrum File - alternating blue, yellow
Landsea.clr     Color Spectrum File - black, blue, cyan, brown, yellow
Rainbow.clr     Color Spectrum File - violet, blue, green, yellow, orange, red
Rainbow2.clr    Color Spectrum File - black, blue, cyan, yellow, red, white
RedHot.clr      Color Spectrum File - red, yellow
Demogrid.dat    Sample data file for DEMOGRID.GRD
Sample3.dat     Sample data used in scripts
Tutorws.dat     Used in Lesson 2 and 5 of the Tutorial
Tutorws2.dat    Used in Lesson 1 of the Tutorial
Vario1.dat      Sample data file for creating variograms
Demogrid.grd    Sample Grid file
Helens2.grd     Sample Grid file from Mount St. Helens DEM
Ca.gsb          Golden Software boundary file - California Counties
Nv.gsb          Golden Software boundary file - Nevada Counties
Sample1.srf     Wireframe map - Colorado Front Range
Sample2.srf     Stacked Wireframe and Contour map - Morrison Colorado
Sample3.srf     Filled Contour map - Sagebrush Hills Idaho
Sample4.srf     Filled Contour map - Grand Canyon Arizona
Sample5.srf     Image map - Colorado Counties
Sample6.srf     Contour/Shaded Relief map - Mt. St. Helens
Sample7.srf     Vector/Contour map of the earth
Sample8.srf     Classed Post map - Antarctic temperatures
Symbols.srf     Post map of symbols and numeric values
Tutorial.srf    Used in Lesson 5 of the Tutorial


Help Script Examples
--------------------
ApplicationObjectProperties.bas
ApplicationObjectMethods.bas
DocumentsCollection.bas
PlotDocumentObject.bas
PlotWindow.bas
RulerObject.bas
ShapeObject.bas
ShapesCollection.bas
VariogramObject.bas
Windows.bas


Surfer.ini Settings
===================
The following advanced settings can be specified under the [Settings] section 
of SURFER.INI.  In general, these settings should not be changed unless there
is a very good reason to do so:

PlotSize=240         ;Plot size in inches (10 to 2000)

DitherScaleColors=1  ;Set to a non-zero value to enable dithering of the
                     ;color scale colors.  This allows the color scale to
                     ;better match the line color used in wireframe plots
                     ;when printing on printers that dither lines.

ErrFlags=147         ;This controls error message display.  The value is the 
                     ;sum of the desired options:
                     ;   1  Display Errors
                     ;   2  Display Warnings
                     ;  16  Display errors in a dialog box
                     ; 128  Beep on errors
                     ; 256  Display Module, File, and Line info

EditReturn=0         ;Set to a non-zero value to allow the edit control in
                     ;the Text Properties dialog accept the Enter key to
                     ;begin a new line.  A value of 0 causes the Enter key
                     ;to shut down the dialog.

MaxReportedDups=100  ;Maximum number of duplicate data to report in the
                     ;gridding and variogram report.

MaxReportedExcl=100  ;Maximum number of excluded data to report in the
                     ;gridding and variogram report.

ScrollPercent=10     ;Percentage of the grid editor window size to scroll 
                     ;when the arrow buttons are pressed.

ScrollPageSize=90    ;Percentage of the plot window size to scroll when the
                     ;page up/down portion of the scroll bar is clicked.

ScrollLineSize=10    ;Percentage of the plot window size to scroll when the
                     ;arrow buttons are pressed.

ZoomMargin=100       ;The margin around the window after Zoom Page.
                     ;Measured in mils (thousandths of an inch)


Known Issues and Limitations
============================
- Gridding is much slower when faults or breaklines are specified.

- Complex faults (many line segments) are much slower than simple faults.

- Most PostScript printer device drivers are NOT 100% Microsoft Windows 
  compliant when it comes to imaging complex polygons.  Under rare circumstances 
  (mostly in filled contour maps) this can result in stray lines appearing 
  when they shouldn't.  We have worked around this in our software, but the 
  modification makes PostScript printing 3 to 5 times slower.  If you want 
  to speed up PostScript printing, but at the risk of an occasional stray line 
  when printing complex filled contour maps, you may do so by placing a 
  semi-colon character (";") in front of the following line in SURFER.INI:

     DecomposePrinterPolygons=2

- HP LaserJet 4 problems: 
  When printing on a LaserJet 4 (or 4M non-PostScript), TrueType text blocks 
  can be wider than requested or character spacing can be incorrect.  This has 
  been fixed in version 31.V1.35 of the LaserJet 4 driver.  You can get this 
  driver from Hewlett Packard at www.hp.com. 

  If you are unable to get the new driver, using the LaserJet III driver can 
  resolve problem.

- HP LaserJet Series II and III problems:
  Single characters in the axis labels, axis titles, and text blocks may not 
  print with some drivers.  In the Control Panel | Printers menu, select the
  printer, and choose Printer | Properties.  Turn on the option to "Print 
  TrueType as Graphics" to solve the problem.

- HP DeskJet problems:
  The last character in text strings may not print.  In the Surfer.ini file 
  remove the semi-colon from the beginning of the line listed below:

     [GSIMAGE]
     ;Windows Picture:Printer:Text=0x8000

  Save the changes, then restart Surfer.

- Video Problems:
  On some video cards (ATI & Trident), the last character in text blocks is not 
  displayed; however, printing does not seem to be affected.

  Solution 1: Remove the semi-colon from the beginning of the line listed below
              in the Surfer.ini file:
      
              [GSIMAGE]
              ;Windows Picture:Window:Text=0x8000

              Save the changes, then restart Surfer.

  Solution 2: Change video driver to use "plain" VGA.

- Additional Surfer.Ini settings:
  There are several additional settings related to polygon decomposition that 
  can be customized within the Surfer.ini file.  These are fairly obscure and 
  should not be needed for normal operation.  However, if you are experiencing 
  problems when outputting large polygons to a display device or metafile, you
  may want to take a look at the comments within Surfer.ini.

- Toolbar positions are not saved when Surfer is maximized:
  The position of the toolbars is based on the NON-maximized window size.  If
  you find the toolbar positions are not being saved when you restart Surfer 
  with a maximized window, follow these steps:

  1. Start Surfer, UN-maximize the main window
  2. Resize the main Surfer window so it almost fills the entire screen
  3. Position the toolbars as desired
  4. Re-Maximize the main window again
  5. The toolbar positions should now be saved correctly


Program Troubleshooting and Error Logging
=========================================
Surfer has an error logging facility that may help us resolve problems.
Run Surfer from the Start | Run menu with the following command line:

   d:\path\surfer.exe /L

where d:\path is the drive and path for Surfer 8.  By default, Surfer 8
is installed in the "c:\Program Files\Golden Software\Surfer8\" directory.  
Include the double-quotes around the path and program name.  Attempt to 
duplicate the problem, and document the steps prior to the occurrence of 
the problem.  Surfer writes information about the state of the program and 
your system to the file Surfer.msg to the Surfer directory.  Zip this file 
and e-mail it to technical support (surfersupport@goldensoftware.com) with 
a description of the problem and the steps necessary to duplicate it.


Resources: FAQ, Support Forums
==============================
The Golden Software web site contains resources to help answer your questions.

- The Surfer FAQ consists of frequently asked questions about the program:
    www.goldensoftware.com/faq/surfer-faq.shtml

- The Surfer Support Forums are open message boards for the discussion of 
  the program and how to perform various tasks:
    www.goldensoftware.ws/forum/


Uninstalling Surfer
====================
To uninstall Surfer, go to Add/Remove programs in the Control Panel.  Select 
Surfer 8 from the list under the Install / Uninstall tab.  Click the Add/Remove 
button to uninstall Surfer.


Contacting Golden Software
==========================
Please report Surfer problems to Golden Software at:

   Golden Software, Inc.
   809 14th Street
   Golden, CO 80401-1866 USA

   phone: 303-279-1021
   fax: 303-279-0909
   e-mail: surfersupport@goldensoftware.com
   web: http://www.goldensoftware.com

Golden Software business hours are 8:00AM to 5:00PM Mountain Time 
Monday through Friday (excluding U.S. holidays).

Check the Golden Software web page for FAQ's, Support Forums and 
for additional information about our products.

### end of Readme.txt ###
