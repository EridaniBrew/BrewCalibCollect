//tabs=4
/*jsl:option explicit*/
//------------------------------------------------------------------------------
//
// Script:       SystemSwap.js 
// Author:       Robert Brewington
//
//  Script to (optionally) Save all the current settings for a "system", 
//  and (optionally) restore the settings from a previous system.
//
//  A "system" is intended to be a particular setup of ACP/Maxim/FocusMax for a 
//  given telescope, imager, and focuser.
//
//  For example, suppose system 1 is "Brew Simulators"
//     ACP set up to use simulators, including simulated star fields. A variety of preferences
//        have been set up for the simulators - shorter exposure times, etc.
//     Maxim - using the simulated camera and guider, with associated filters.
//     Focuser - using the system file BrewSimulator.ini. 
//	      This contains the Profile data from running VCurves on the simulator,
//        as well as the various settings for exposure times, use 75% of CCD, etc.
//
//  System 2 is "Sky90SBigIntGuide", which would be my standard production system.
//     ACP has all my usual values for production runs
//     Maxim is set to use the SBig ST2000XM with CFW10 filter wheel,using internal guiding
//     FocusMax is set with the appropriate VCurves and settings using the LazyFocus, stored in 
//        the file Sky90SBigLZFocus.ini
//
//  Now, I can swap in either system and be set up to run, without having to run through an
//		error-prone checklist of settings.
//
//	Information is saved in the directory structure
//		Prefs.WebRoot/SystemSwap/
//			log files from each run
//			Brew Simulators/
//				ACP Settings/
//					ACPregistry.reg					Contains the registry settings for ACP
//					FilterInfo.txt
//					Settings.txt					Various settings I need to set specifically
//					Active.clb						Current pointing model
//
//				Maxim Settings/
//					Settings.txt					Settings I need to set individually
//					Settings/						Copy of the folder Prefs.WebRoot-parent/MaxIm DL 5/Settings
//						
//				FMx Settings/
//					Settings.txt					Settings I need to set individually
//					BrewSimulator.ini				Copy of the FMx ini file
//
//			Sky90SBigIntGuide/
//				ACP Settings/						Same info as above	
//				Maxim Settings/
//				FMx Settings/
//		
//	NOTE:
//		Other users need to change the variable docConfigDir below to point to the
//		folder where the configuration information Maxim and FocusMax are found
//		
// Version:     
// 1.1.0   - Add System Description file, so I can use FileDialog to select
//           systems for saving/loading. Pain to type in the file names all the time.
//         - System file can contain attributes. First one - the version of
//           BrewSystemSwap that saved the last system file. This will allow logic to
//           update systems saved under older script. Maybe.
//         - Save RotatorConfig.txt (if there is one). 
//         - Add SystemSwap/000LastLoaded file, which contains the name of the last system loaded.
//           Maybe I can display the current system name somewhere (ACP header would be nice!)

var SCRIPTVERSION = "1.1.0";

// Requires:     ACP 6.0 or later! I don't know what happens in ACP 5 or earlier
var ACP_VERSION = "6.0.0";

// This is where we find the Maxim and FocusMax individual configuration files, etc
// something like C:/Users/brew/Documents
// You need to change this to whatever your folder is
var docConfigDir = "c:\\Users\\brew\\Documents";					

//
// Constants
//
// These are various ACP registry keys for different Windows systems
var ACPKEYWin764 = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Denny\\ACP";	
var ACPKEYXP = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Denny\\ACP";
var ACPkey;										// Holds the ACP key that seems to work	
var ACPREGFILE = "ACPRegistry.reg";				// Saves the ACP registry values

// ACP Files needing to be saved/restored (relative to ACPPath)
var ACP_File_List = ["FilterInfo.txt", 
					"AutoFlatConfig.txt", 
					"RotatorConfig.txt",
					"FlipConfig.txt"];

var FOCUSMAX = "FocusMax.exe";					// Processes for starting/stopping
var MAXIM = "MaxIm_DL.exe";

var LASTLOADED = "000LastLoaded";					// SystemSwap/LastLoaded contains the system name last loaded
//
// Enhance String with a trim() method
//
String.prototype.trim = function()
{
    return this.replace(/(^\s*)|(\s*$)/g, "");
};

//
// Control variables
//
var FSO, SUP;
var restoreRootDir;			// Put folder SystemSwap in something like /users/Public/Documents/ACP Web Data/Doc Root

// Things for Verifying the Registry file
var REG_FIRST_LINE = "Windows Registry Editor";
var HKEY_LINE = "[HKEY_LOCAL";


// ==========
// MAIN ENTRY
// ==========

function main()
{
    //
    // Control variables, scoped to main()
    // Some things have already been set up by ACP: 
    // Console, Util, ACP, Prefs
    //
    var buf, button;       
    var saveCurrent, restoreIn;                                                         // Used all over the place
    
    FSO = new ActiveXObject("Scripting.FileSystemObject");
    
    // Check the ACP Version. For now, I need to be ACP 6.0.0 or more
    try {
		var acpVer = GetACPVersion();
		} catch(e) {
		Console.PrintLine("Error accessing the ACP Registry");
		Console.PrintLine("  " + (e.description ? e.description : e));
	    return;
		}
    var acpVerNum = CvtVersionToNum(acpVer);
    var reqVerNum = CvtVersionToNum(ACP_VERSION);
    if (reqVerNum > acpVerNum)
		{
		buf = "This script requires ACP Version " + ACP_VERSION + " or later. \n\nYou have ACP Version " + acpVer;
		Console.PrintLine (buf);
		Console.Alert (consIconOK, buf);
		FSO = null;
		return;
		}
	
    // for both localUser and WebUser, use WebRoot to locate SystemSwap
	// use double parent to locate docConfigDir
	var targDir = Prefs.WebRoot;    
	buf = MakeLogFilePathName(targDir);    
	Console.LogFile = buf; 
    Console.Logging = true;                                                 // === BEGIN LOGGING ===
    swapRootDir = targDir + "\\SystemSwap";
	
	
    Console.PrintLine ("BrewSystemSwap Version " + SCRIPTVERSION);
    Console.PrintLine ("Logging results to log file " + Console.LogFile);
    Console.PrintLine ("Looking for config information for MaxIm, FMx in " + docConfigDir);
    Console.PrintLine("Location of system Save directory is " + swapRootDir);
		
    CheckSystemFiles();		// verify system directories have corresponding files 
    

    // OK, what do we want to do? Ask User
    var saveName;       // serves as the target system name for saving to; used to construct the folder name
    var restorePath;		// The source path name of the directory containing the values to be installed
    var restoreName;		// just the name of the restore system
	var savePath;
    
    saveCurrent = Console.AskYesNo("Do you want to save the current system setup?");
	if (saveCurrent)
		{
		FileDialog.DefaultExt = ".txt";                             // Use the file browser
        FileDialog.DialogTitle = "Save System";
        FileDialog.Filter = "System Files (*.txt)|*.txt|All files (*.*)|*.*";
        FileDialog.FilterIndex = 1;
        FileDialog.Flags = fdHideReadOnly ;   // Does not need to exist; hide read only
        FileDialog.InitialDirectory = swapRootDir;
        if(FileDialog.ShowOpen()) 
			{  // saveName is like BrewSim; savePath is c:/xxx/yyy/BrewSim 
			saveName = FSO.GetBaseName(FileDialog.FileName);
			var dotPos = FileDialog.FileName.lastIndexOf(".txt");
			if (dotPos == -1)
				{
				savePath = FileDialog.FileName;                              // SwapFolder contains the system elements
				}
			else
				{
				savePath = FileDialog.FileName.substring(0,dotPos);
				}
			saveCurrent = true;
			}
		else
			{
			saveCurrent = false;
			}
		}
	if (saveCurrent)
	    {
	    Console.PrintLine ("");
	    Console.PrintLine ("");
	    Console.PrintLine ("******************************************");
	    Console.PrintLine ("Saving current system as " + savePath);
	    // so do it
	    try {
	        SaveCurrentSystem (savePath);
	        } catch(e){
	        Console.PrintLine ("System Save Failed. ");
	        Console.PrintLine("  " + (e.description ? e.description : e));
	        SUP = null;
	        return;
	        }
	    }
	else
	    {
	    Console.PrintLine ("User elected not to save current system");
	    }
	    
	    
	restoreIn = false;
    if (Console.AskYesNo("Do you want to restore a previously stored system setup?"))
	    {
	    FileDialog.DefaultExt = "*.txt";                             // Use the file browser
        FileDialog.DialogTitle = "Select a System to Restore";
        FileDialog.Filter = "System Files (*.txt)|*.txt|All files (*.*)|*.*";
        FileDialog.FilterIndex = 1;
        FileDialog.Flags = fdHideReadOnly + fdFileMustExist;   // 4096 + 4   Must exist and hide read only
        FileDialog.InitialDirectory = swapRootDir;
        if(FileDialog.ShowOpen()) 
			{
			restoreName = FSO.GetBaseName(FileDialog.FileName);
			var dotPos = FileDialog.FileName.lastIndexOf(".txt");
			if (dotPos == -1)
				{
				restorePath = FileDialog.FileName;                              // SwapFolder contains the system elements
				}
			else
				{
				restorePath = FileDialog.FileName.substring(0,dotPos);
				}
			restoreIn = true;
			}
		}
			
	if (restoreIn)
	    { // OK, do it
	    Console.PrintLine ("");
	    Console.PrintLine ("");
	    Console.PrintLine ("******************************************");
	    Console.PrintLine ("Swapping in system " + restorePath);
	    // do it here  restorePath has the src folder
	    try {
			RestoreNewSystem (restorePath);
			CreateLastLoaded(restoreName);
			Console.Alert (consIconOK, "          NOTE!\n" +
					"The ACP program will now be stopped. \n" +
					"Restarting it will pick up the new settings."); 
			if(!StopProcess("acp.exe")){
				Console.PrintLine("*** Failed to stop ACP");
				}
			} catch(e) {
	        Console.PrintLine ("System Restore Failed. ");
	        Console.PrintLine("  " + (e.description ? e.description : e));
	        SUP = null;
	        return;
			}
		}
	else
	    {
	    Console.PrintLine ("User elected not to restore a system.");
	    }
	
	Console.PrintLine ("============ Script completed ==============");
	Console.Logging = false;
}


// ------------------
// MakeLogFilePathName() - Construct a log file path/name 
// ------------------
//
//
function MakeLogFilePathName(root)
{
    var bdir, lfn;
    
    
    //
    // Construct/use folder path and create the folder if needed
    //
    bdir = root + "\\" + "SystemSwap"; // Substitute for directory path
    CreateFolder(bdir);                                                     // Make sure folder exists
    
    //
    // Construct the file name safely 
    //
    lfn = "SystemSwap-" + Util.FormatVar(new Date().getVarDate(), "yyyymmdd-HhNnSs") + ".log";
    
    
    
    return bdir + "\\" + lfn;                                               // Return final path/name
}


// --------------
// CreateFolder() - Create a folder with recursion
// --------------
//
function CreateFolder(f)
{
    if(FSO.FolderExists(f)) return;                                         // Already exists, no-op

    var p = FSO.GetParentFolderName(f);                                     // Get parent name (may not exist)
    if(!FSO.FolderExists(p))                                                // If doesn't exist
        CreateFolder(p);                                                    // Call ourselves recursively to create it
    FSO.CreateFolder(f);                                                    // Now create the given folder
}

// ----------------
// SafeDeleteFile() - Delete a file whether it exists or not
// ----------------
//
function SafeDeleteFile(f)
{
    try {
        FSO.DeleteFile(f);
    } catch(e) {  }
}
function SafeDeleteFolder(f)
{
    try {
        FSO.DeleteFolder(f);
    } catch(e) {  }
}

// --------------
// CheckFolderExists(path) - Verify that the specified folder exists
// --------------
//
function CheckFolderExists (path)
{
	return FSO.FolderExists(path);
}

// --------------
// CheckFileExists(path) - Verify that the specified file exists
// --------------
//
function CheckFileExists (path)
{
	return FSO.FileExists(path);
}

// --------------
// SaveCurrentSystem(savePath) - Save the current settings to the specified directory
// --------------
//
function SaveCurrentSystem(savePath)
{
	// create the needed directories
	CreateFolder (savePath);
	SafeDeleteFolder (savePath + "\\Maxim Settings");    // clearing out any previous files
	SafeDeleteFolder (savePath + "\\ACP Settings");    // clearing out any previous files
	SafeDeleteFolder (savePath + "\\FMx Settings");    // clearing out any previous files
	CreateFolder (savePath + "\\ACP Settings");
	CreateFolder (savePath + "\\FMx Settings");
	CreateFolder (savePath + "\\Maxim Settings");
	CreateFolder (savePath + "\\Maxim Settings\\Settings");
	
	SaveACPSettings(savePath + "\\ACP Settings");
	SaveFMxSettings(savePath + "\\FMx Settings");
	SaveMaximSettings(savePath + "\\Maxim Settings");
	CreateSystemFile (savePath + ".txt");
}

// --------------
// RestoreNewSystem(restorePath) - Restore a new setup from the specified directory
// --------------
//
function RestoreNewSystem(restorePath)
{
	// Verify that the directories and files are there
	if (! CheckFolderExists (restorePath))
		{
		throw "Cannot find folder " + restorePath ;
		}
	if (! CheckFolderExists (restorePath + "\\ACP Settings"))
		{
		throw "Cannot find folder " + restorePath + "\\ACP Settings";
		}
	if (! CheckFolderExists (restorePath + "\\FMx Settings"))
		{
		throw "Cannot find folder " + restorePath + "\\FMx Settings";
		}
	if (! CheckFolderExists (restorePath + "\\Maxim Settings"))
		{
		throw "Cannot find folder " + restorePath + "\\Maxim Settings";
		}
	if (! CheckFolderExists (restorePath + "\\Maxim Settings\\Settings"))
		{
		throw "Cannot find folder " + restorePath + "\\Maxim Settings\\Settings";
		}
	
	DisconnectACP();
	// Stop FocusMax and Maxim. 
	if(!StopProcess(FOCUSMAX))
        Console.PrintLine("*** Failed to stop FocusMax");
    Util.WaitForMilliseconds(1000);
    if(!StopProcess(MAXIM))
        Console.PrintLine("*** Failed to stop MaxIm");
    Util.WaitForMilliseconds(1000);
    
    RestoreACPSettings(restorePath + "\\ACP Settings");
	
	
    RestoreFMxSettings(restorePath + "\\FMx Settings");
	
	RestoreMaximSettings(restorePath + "\\Maxim Settings");
	
}

// --------------
// SaveACPSettings(savePath) - Save the current settings to the specified directory 
//   savePath is like Mysystem/ACP Settings
// --------------
//
function SaveACPSettings(savePath)
{
	Console.PrintLine("");
	Console.PrintLine ("Saving ACP Settings");
	for (var i in ACP_File_List)
		{
		var fileName = ACP_File_List[i];
		if (FSO.FileExists(ACPApp.Path + "\\" + fileName))
			{  // Copy FilterInfo.txt
			FSO.CopyFile(ACPApp.Path + "\\" + fileName, savePath + "\\" + fileName);
			Console.PrintLine("   " + fileName + " saved");
			}
		}
	
	// Copy Active.clb file 
	// need to delete previously saved Active.clb 
	SafeDeleteFile(savePath + "\\Active.clb");
	var srcFile = Util.PointingModelDir + "\\Active.clb";
	if (FSO.FileExists(srcFile))
		{  // Copy Active.clb
		FSO.CopyFile(srcFile, savePath + "\\Active.clb");
		Console.PrintLine("   Active.clb saved");
		}
		
	// Create Settings file
	/************************* Don't need this for now
	var fileOK = false;
	try {
		var settingFile = FSO.CreateTextFile(savePath + "\\Settings.txt", true);  // Make new Settings file
		fileOK = true;
	} catch(e) {
        Console.PrintLine("**Failed to create ACP Settings file: " + 
            (e.description ? e.description : e));
    }	
    
    // Write ACP settings
	if (fileOK)
		{
		settingFile.WriteLine("Setting 1: 123");                                 
	
		settingFile.Close();
		}
	****************/
    
	// Create ACP Registry file
	ExportRegistryFile (savePath + "\\" + ACPREGFILE);
	VerifyRegistryFile (savePath + "\\" + ACPREGFILE);
	
}

// --------------
// SaveFMxSettings(savePath) - Save the current settings to the specified directory (Mysystem/FMx Settings)
//    save Ascom Profile settings
//    save the .ini file
// --------------
//
function SaveFMxSettings(savePath)
{
	Console.PrintLine("");
	Console.PrintLine ("Saving FocusMax Settings");
	// Create Settings file
	//var FMx = new ActiveXObject("FocusMax.FocusControl");
	var fileOK = false;
	try {
		var settingFile = FSO.CreateTextFile(savePath + "\\Settings.txt", true);  // Make new Settings file
		fileOK = true;
	} catch(e) {
        throw ("**Failed to create FMx Settings file: \n" + 
            (e.description ? e.description : e));
    }	
    
    // Write FMx settings
    // loop through the settings in the Ascom profile for Focusmax
    // skip the first (blank) key
    if (! fileOK) return;

	var pr;
	try {
		pr = new ActiveXObject("DriverHelper.Profile");
	} catch (e) {
	    throw "SaveFMx: Could not create DriverHelper.Profile " + (e.description ? e.description : e);
    }
	pr.DeviceType = "Focuser";

	var subkeys = pr.Values('FocusMax.Focuser');
	var enu = new Enumerator(subkeys);                                  
	for(; !enu.atEnd(); enu.moveNext())
		{
		var key =  enu.item();
		if (key != "")
			{
			var value = pr.GetValue('FocusMax.Focuser',key);
			settingFile.WriteLine (key + ": " + value);
			}
		}

	settingFile.Close();
	Console.PrintLine("   FocuxMax settings saved");

	// Copy FMx .ini file    
	var fileName = pr.GetValue('FocusMax.Focuser',"System Name");
	var srcPath = pr.GetValue('FocusMax.Focuser',"System Path") + "\\" + fileName;
	var tarPath = savePath + "\\" + fileName;
	if (FSO.FileExists(srcPath))
		{
		try {
			FSO.CopyFile(srcPath, tarPath);	
			Console.PrintLine("   FocusMax system file saved");
		} catch(e){
			throw ("Could not copy FocusMax file " + srcPath + (e.description ? e.description : e));
			}
		}
	pr = null;
	enu = null;
}

// --------------
// SaveMaximSettings(savePath) - Save the current settings to the specified directory (Mysystem/Maxim Settings)
// --------------
//
function SaveMaximSettings(savePath)
{
	Console.PrintLine("");
	Console.PrintLine ("Saving Maxim Settings");
	/*
	// Create Settings file
	var MaxIm = new ActiveXObject("MaxIm.Application");
	// var Camera = MaxIm.CCDCamera;
	var fileOK = false;
	try {
		var settingFile = FSO.CreateTextFile(savePath + "\\Settings.txt", true);  // Make new Settings file
		fileOK = true;
	} catch(e) {
        Console.PrintLine("**Failed to create MaxIm Settings file: " + 
            (e.description ? e.description : e));
    }	
    
    // Write Maxim settings
	if (fileOK)
		{
		// MaxIm's version is a single but can have roundoff - 5.0599994 instead of 5.06
		settingFile.WriteLine("MaxIm Version: " + Util.FormatVar(MaxIm.Version, "0.00"));                                 
		
		settingFile.Close();
		}	
	//MaxIm =null;
	*/
	
	// Save the various .txt files from docConfigDir/MaxIm DL 5/Settings
    var src = docConfigDir + "\\MaxIm DL 5\\Settings";
    var tar = savePath + "\\Settings";
		{
		try {
		    FSO.CopyFolder(src , tar, true);
			} catch(e){
			throw "Could not copy Maxim folder " + src +
				(e.description ? e.description : e);
			}
		}
    
    Console.PrintLine("   Maxim Settings files saved");
	
}

// --------------
// RestoreACPSettings(savePath) - restore the settings from the specified directory 
//   savePath is like Mysystem/ACP Settings
// --------------
//
function RestoreACPSettings(savePath)
{
	// Copy savePath/FilterInfo.txt to ACPPath + "\\FilterInfo.txt"
	// need to delete current ACPPath/FilterInfo.txt
	Console.PrintLine("");
	Console.PrintLine("Restoring ACP Settings");	
	for (var i in ACP_File_List)
		{
		var fileName = ACP_File_List[i];
		SafeDeleteFile(ACPApp.Path + "\\" + fileName);		// delete existing FilterInfo
		if (FSO.FileExists(savePath + "\\" + fileName))
			{  // Copy FilterInfo.txt
			FSO.CopyFile(savePath + "\\" + fileName, ACPApp.Path + "\\" + fileName);
			Console.PrintLine("   " + fileName + " restored");
			}
		}
	
	// restore Active.clb
	// If no points are in the model, we do not have a file Active.clb
	var tarFile = Util.PointingModelDir + "\\Active.clb";
	var srcFile = savePath + "\\Active.clb";
	SafeDeleteFile(tarFile);		// delete existing model
	if (FSO.FileExists(srcFile))
		{  // Copy source to target
		FSO.CopyFile(srcFile, tarFile);
		Console.PrintLine("   Pointing Model restored");
		}
	else
		{
		Console.PrintLine("   No saved Pointing Model.\n   Pointing Model Cleared");
		}
	
	
	// restore ACP registry values
	if (VerifyRegistryFile (savePath + "\\" + ACPREGFILE))
	    {
		ImportRegistryFile (savePath + "\\" + ACPREGFILE);
		}
}

// --------------
// RestoreMaximSettings(savePath) - Restore the settings to the specified directory (Mysystem/Maxim Settings)
// --------------
//
function RestoreMaximSettings(savePath)
{
	Console.PrintLine("");
	Console.PrintLine("Restoring Maxim Settings");	
	
		// Save the various .txt files from docConfigDir/MaxIm DL 5/Settings
    var tar = docConfigDir + "\\MaxIm DL 5\\Settings";
    var src = savePath + "\\Settings";
		{
		try {
		    FSO.CopyFolder(src , tar, true);
			} catch(e){
			throw "Could not copy to Maxim folder " + tar +
				(e.description ? e.description : e);
			}
		}
 
    // Restore the various .txt files from docConfigDir/MaxIm DL 5/Settings
	/*****
	var srcDir = savePath + "\\Settings\\";
	var tarDir = docConfigDir + "\\MaxIm DL 5\\Settings\\";
	var fld = FSO.GetFolder(srcDir);
    var e = new Enumerator(fld.Files);                                  // Enumerate immediate parent folder's files
    for(; !e.atEnd(); e.moveNext())
    {
        var fn = e.item().Name;                                         // These are file names only
        if(fn.search(/.txt$/i) >= 0)                                    // If this is a txt file
			{
            Console.PrintLine ("      Copying " + fn);
            FSO.CopyFile(srcDir + fn, tarDir + fn);     // Copy to new folder
            }
    }
    *******/
    Console.PrintLine("   Maxim settings files restored");
	
}

//
// Stop a process by executable name. WMI magic.
//
var WMI = null;                                                         // Avoid creating lots of these
function StopProcess(exeName)
{
    var x = null;
    Console.PrintLine("   Stopping " + exeName);
    try {
        // Magic WMI incantations!
        if(!WMI) WMI = new ActiveXObject("WbemScripting.SWbemLocator").ConnectServer(".","root/CIMV2");
    	x = new Enumerator(WMI.ExecQuery("Select * From Win32_Process Where Name=\"" + exeName + "\"")).item();
        if(!x) {                                                        // May be 'undefined' as well as null
            Console.PrintLine("   (" + exeName + " not running)");
        	return true;                                                // This is a success, it's stopped
        }
        x.Terminate();
        Console.PrintLine("   OK");
        return true;
    } catch(ex) {
        Console.PrintLine("*** WMI: " + (ex.message ? ex.message : ex));
        return false;
    }
}

//
// Start program by executable path, only if not already running.
// Waits 15 sec after starting for prog to initialize.
//
function StartProgram(exePath, windowState)
{
    try {
        var f = FSO.GetFile(exePath);
        if(IsProcessRunning(f.Name))                                    // Proc name is full file name
            return true;                                                // Already running
        if(IsProcessRunning(f.ShortName))                               // COM servers can have proc name of file short name
            return true;
        Console.PrintLine("Starting " + f.Name);
        Util.ShellExec(exePath, "", windowState);
        Util.WaitForMilliseconds(5000);                                // Wait for prog to initialize (was 15 sec)
        Console.PrintLine("OK");
        return true;
    } catch(ex) {
        Console.PrintLine("Exec: " + ex.message);
        return false;
    }
}

//
// Test if program is running (by exe name)
//
function IsProcessRunning(exeName)
{
    try {
        // Magic WMI incantations!
        if(!WMI) WMI = new ActiveXObject("WbemScripting.SWbemLocator").ConnectServer(".","root/CIMV2");
    	var x = new Enumerator(WMI.ExecQuery("Select * From Win32_Process Where Name=\"" + exeName + "\"")).item().Name;
    	return true;
    } catch(e) {
    	return false;
    }
}

// --------------
// DisconnectACP() - Disconnect Telescope and Camera in ACP
// --------------
//
function DisconnectACP()
{
    if(Telescope.Connected) {
        Console.PrintLine("   Parking scope"); 
        Telescope.Park();                                               // Park the scope if possible, and/or close/park dome
        if(Telescope.Connected)
            Telescope.Connected = false;                                // Disconnect it from ACP
        Console.PrintLine("   Telescope Disconnected");
    }
    
    if(Util.CameraConnected) 
		{
        Camera.LinkEnabled = false;
        Util.CameraConnected = false;                                   // Disconnect it from ACP
        Console.PrintLine("   Imager disconnected.");
        }
}

// --------------
// RestoreFMxSettings(savePath) - Restore the settings from the specified directory (Mysystem/FMx Settings)
// --------------
//
function RestoreFMxSettings(restorePath)
{
	// Open Settings file
	var FMx;
	
    Console.PrintLine("");
	Console.PrintLine("Restoring FocusMax Settings");	
	
	// Read FMx settings
    var settingsColl = ParseSettingsFile(restorePath + "\\Settings.txt");
    // Restore FMx .ini file    
    var fileName = settingsColl["System Name"];
	var tarPath = settingsColl["System Path"]  + "\\" + fileName;
	var srcPath = restorePath + "\\" + fileName;
	if (FSO.FileExists(srcPath))
		{
		try {
			FSO.CopyFile(srcPath, tarPath,true);
			Console.PrintLine("   FocusMax system file restored");
		} catch(e){
			throw("ERROR: Could not copy FocusMax file " + srcPath + " to " + tarPath + "\n" + (e.description ? e.description : e));
			}
		}
	else
		{
		throw ("ERROR: could not find source FMx System File " + srcPath);
		}
		
	
	var pr;
	try {
		pr = new ActiveXObject("DriverHelper.Profile");
	} catch (e) {
	    throw "Restore FMx: Could not create DriverHelper.Profile " + (e.description ? e.description : e);
    }
	pr.DeviceType = "Focuser";

	for(var key in settingsColl)
		{
		pr.WriteValue('FocusMax.Focuser',key, settingsColl[key]);
		}

	Console.PrintLine("   FocusMax Profile Settings restored");
	settingsColl=  null;
	pr = null;
}

//
// GetACPVersion
//		Read the Version from the Registry
//
function GetACPVersion()
{
	var wsh;
	var buf;
	var key;

	buf = "0.0.0";
	wsh = new ActiveXObject ("WScript.Shell");
	
	//
	// We need to try various registry keys depending on the Windows OS
	ACPkey = "";
	try 
		{  // first, try XP
		buf = wsh.RegRead (ACPKEYWin764 + "\\Ver");
		ACPkey = ACPKEYWin764;
		Console.PrintLine("Appears to be Windows 7 64 bit");
	} catch(e) {
	    try {
		    buf = wsh.RegRead (ACPKEYXP + "\\Ver");
			ACPkey = ACPKEYXP;
			Console.PrintLine("Appears to be Windows XP");
	    } catch(e){
        throw("GetACPVersion: There was a problem reading the registry:\n" +
					"  " + e.description);
        }
    }
	wsh = null;
	
	return buf;
}

function CvtVersionToNum(ver)
{
	var pieces = ver.split(".");
	var num = pieces[0] * 10000 + 100 * pieces[1] + pieces[2];
	return num;
}

//
// ParseSettingsFile
//    Read the Settings file and break it into pieces. 
//    skip lines starting with ;
//    Lines should be in the form name:value
//        Example:     Focal length   : 400
//    Returned array of objects can be addressed like
//    x = myCollection['Focal length'] 
//    to give x the value "400"
function ParseSettingsFile(path)
{
	var settingFile;
    var myColl = new Array();
    var key, data; 

	try {
		settingFile = FSO.OpenTextFile (path, 1,false);  // Open Settings file
	} catch(e) {
        throw("**Failed to open Settings file " + path + ": " + 
            (e.description ? e.description : e));
    }	
    
    // loop through the file
    var buf;
    var key,dataStr, colPos;
    while (! settingFile.AtEndOfStream)
		{
		buf = settingFile.ReadLine();
		// check for comment -- ; in first pos
		var first = buf.substring(0,1);
		if (first != ";")
			{
			colPos = buf.indexOf(":");
			if (colPos > 0)
				{
				key = buf.slice(0,colPos).trim();
				dataStr = buf.slice(colPos+1).trim();
				myColl[key] = dataStr;
				}
			}
		}
	settingFile.Close();
	
	return myColl;
}

//
// ExportRegistryFile (filepath) 
// runs a command shell to run regedit to export the subroot of the ACP registry
//
function ExportRegistryFile(filePath)
{
	// Regedit to export the ACP values
	try {
		var obShell = new ActiveXObject("Shell.Application");
	} catch(e) {
        throw("**Failed to create Shell.Application to run regedit " + 
            (e.description ? e.description : e));
        return;
  	} 
	
	var buf;
	buf = "/e \"" + filePath + "\" \"" + ACPkey +  "\" ";
	
	try
		{
		obShell.ShellExecute("regedit.exe", buf, "", "open", "1");
	} catch(e) {
        obShell = null;
        throw("**ImportRegistryFile: Failed to execute regedit " + buf + " \n" +
            (e.description ? e.description : e));
    }
	
	// Need to wait while this runs?
	Util.WaitForMilliseconds(3000);
	Console.PrintLine ("   ACP Registry Export complete");
	
	obShell = null;
}

//
// ImportRegistryFile (filepath)
// Runs a command shell to run regedit to export the subroot of the ACP registry
//
function ImportRegistryFile(filePath)
{
	try {
		var obShell = new ActiveXObject("Shell.Application");
	} catch(e) {
        throw("**Failed to create Shell.Application to run regedit " + 
            (e.description ? e.description : e));
        return;
  	} 
	
	// /s and /S cause regedit not to prompt for whether we really want to do this
	var buf = "/s /S \"" + filePath + "\" ";
	try
		{
		obShell.ShellExecute("regedit.exe", buf, "", "open", "1");
	} catch(e) {
        obShell = null;
        throw("**ImportRegistryFile: Failed to execute regedit " + 
            (e.description ? e.description : e));
    } 
	
	Util.WaitForMilliseconds(3000);
	Console.PrintLine ("   ACP Registry Import complete");
	
	obShell = null;
}

//
// VerifyRegistryFile (filePath)
// Check the saved ACP registry file to make sure it looks like it worked
// returns true if OK, false indicating an error.
//
function VerifyRegistryFile (filePath)
{
	var regFile;
    var buf;
    var lineNum = 0;
    var hkeyLines = 0;				// number of lines starting with HKEY_LINE
    var e;
	
	// Check that file exists
	if (! FSO.FileExists(filePath))
		{
		throw("**VerifyRegistryFile: Registry file does not exist " + filePath);
		}
		
	try {
		// NOTE - regedit writes the exported file as Unicode, not Ascii
		regFile = FSO.OpenTextFile(filePath, 1, false, -1);  // ForReading is 1   Unicode is -1
	} catch(e) {
        throw("**VerifyRegistryFile: Failed to open Registry file " + filePath + ": " + 
            (e.description ? e.description : e));
    }	
    
    // loop through the file counting the number of [HKEY lines
    // Check first line for Windows Registry Editor
    while (regFile.AtEndOfStream == false)
		{
		try {
			buf = regFile.ReadLine();
			} catch(e) {
			throw("**VerifyRegistryFile: ReadLine failed " + 
				(e.description ? e.description : e));
			}
		// check first line
		if (lineNum == 0)
			{
			if (buf.indexOf(REG_FIRST_LINE) != 0)
				{
				regFile.Close();
				throw ("VerifyRegistryFile failed. First line starts with " + 
						buf.substring(0,22) + 
						"Should be " + REG_FIRST_LINE);
				}
			}
		
		// check for HKEY line
		if (buf.indexOf(HKEY_LINE) == 0)
			{
			hkeyLines++;
			}
		lineNum++;
		}
	
	// Did we get a bunch of HKEY lines?
	if (hkeyLines < 20)
		{
		regFile.Close();
		throw ("***VerifyRegistryFile failed. Expecting at least 20 HKey lines. Only found " + hkeyLines);
		}
	regFile.Close();
	
	Console.PrintLine ("   Registry file verified");
	return true;
}

//
// DisplayCurrentSystems
// Since I don't seem to have a FileDialog which selects a folder,
// we hack it up here.
// Display a list of the folders within SystemSwap
//  returns the number of system folders found
function DisplayCurrentSystems(systemPath)
{
	var fld = FSO.GetFolder(systemPath);
	var folderCount = fld.SubFolders.Count;
    if (folderCount > 0)
		{
		Console.PrintLine("");
		Console.PrintLine("");
		Console.PrintLine("**** Current Saved Systems: ****");
		var e = new Enumerator(fld.SubFolders);                                  // Only enumerates subFolders
		for(; !e.atEnd(); e.moveNext())
			{
			    Console.PrintLine ("  " + e.item().Name);
			}
		}
	return folderCount;
}

//
// CheckSystemFiles
// Check for necessary upgrades to saved systems
// For now: verify that existing systems have corresponding system files
// If not, upgrade them:
//      a) add system file
//      b) copy existing RotatorConfig.txt from current system to the saved system
function CheckSystemFiles()
{
	var fld = FSO.GetFolder(swapRootDir);
    var folderCount = fld.SubFolders.Count;
    if (folderCount > 0)
		{
		Console.PrintLine("");
		Console.PrintLine("");
		Console.PrintLine("**** Checking Saved Systems: ****");
		var e = new Enumerator(fld.SubFolders);  
		var converted = 0;                                // Only enumerates subFolders
		for(; !e.atEnd(); e.moveNext())
			{
			var sysFilePath = swapRootDir + "\\" + e.item().Name + ".txt";
			if (!CheckFileExists(sysFilePath))
				{
				converted++;
				Console.PrintLine("   " + e.item().Name + " needs conversion");
				CreateSystemFile (sysFilePath);
				
				// Copy RotatorConfig.txt file if found
				var Rotsrc = ACPApp.Path + "\\RotatorConfig.txt";
				var Rottar = swapRootDir + "\\" + e.item().Name + "\\ACP Settings\\RotatorConfig.txt";
				if (FSO.FileExists(Rotsrc))
					{  // Copy FilterInfo.txt
					FSO.CopyFile(Rotsrc, Rottar);
					Console.PrintLine("   RotatorConfig.txt saved");
					}
				Console.PrintLine("   Done.");
				}
			}
		if (converted == 0)
		   {
		   Console.PrintLine("All systems OK.");
		   }
		}
}

//
// CreateSystemFile(systemName)
//    systemName is the desired path name of the system, including the .txt
function CreateSystemFile(systemName)
{
	try {
		var sysFile = FSO.CreateTextFile(systemName, true);  // Make new system file
		} catch(e) {
			throw "**Failed to create system file: " + systemName + " : " +
			(e.description ? e.description : e);
		}	
	sysFile.WriteLine("BrewSystemSwap Version: " + SCRIPTVERSION);                                 

	sysFile.Close();
}

//
// CreateLastLoaded(sysName)
//    Create LastLoaded file, containing the name of the system just restored
function CreateLastLoaded (sysName)
{
	try {
		var sysFile = FSO.CreateTextFile(swapRootDir + "\\" + LASTLOADED, true);  // Make new system file
		} catch(e) {
			throw "**Failed to create file: " + LASTLOADED + " : " +
			(e.description ? e.description : e);
		}	
	sysFile.WriteLine(sysName);                                 

	sysFile.Close();
}