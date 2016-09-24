//tabs=4
/*jsl:option explicit*/
//------------------------------------------------------------------------------
//
// Script:       BrewCalibCollect.js 
// Author:       Robert Brewington
//
//  ACP Script to collect whatever darks, bias, and flats are in a directory and create the appropriate Master files for
//  MaximDL in the Calibration process.
//
//  Overall process:
//    1. Scan the folder, build the object structure describing the files. Both Masters and Subs are in the structure.
//
//    2. Create the Master files for the discovered subs.
//
//    3. Remove the subs and old masters being replaced. Save the newly created masters. For now, removing means moving the files into a 
//       subfolder "Removed". When things seems to work well, will delete the files instead.
//
//    4. Run the Maxim routine to replace the existing masters in Maxim. This is the equivalent of the  Auto-Generate (Clear Old) button in Maxim.
//       We has to do step 2 because Maxim does not provide the routine equivalent to the Replace with Masters button.
//
//  Log file lists information about the run.
//
//	NOTE:
//		Other users need to change the variables below to point to the
//		folder where the calibration files are found.
//      You would also need to verify / change the logic which determines the parameters of
//      each file. My naming conventions may be different.
//		
// Version:     
// 1.0.0   - Initial version

var SCRIPTVERSION = "1.0.0";

// This is where we find the Flat, Bias, Dark, and previous Master files to be added to the Calibration sets.
// You need to change this to whatever your folder is
var calibDir = "C:\\Users\\erida\\Dropbox\\BrewSky\\Programs\\BrewCalibCollect\\STF8300";

//
// Constants
//

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
    var masterDir = "";
    var indivDir = "";
    
    FSO = new ActiveXObject("Scripting.FileSystemObject");
    
    SUP = new ActiveXObject("ACP.AcquireSupport");
    SUP.Initialize();


	
	// use double parent to locate docConfigDir
    var targDir = indivCalibDir;   // Prefs.WebRoot;
	buf = MakeLogFilePathName(targDir);    
	Console.LogFile = buf; 
    Console.Logging = true;                                                 // === BEGIN LOGGING ===
    
    Console.PrintLine("");
    Console.PrintLine("");
    Console.PrintLine("BrewCalibCollect Version " + SCRIPTVERSION);
    Console.PrintLine ("Logging results to log file " + Console.LogFile);
    	
    
    console.printline("Using calibration directory " + calibDir);
    
    // traverse the files in the Calib directory
    //var indivFileColl, masterFileColl;
    //indivFileColl = TraverseIndivFiles(indivCalibDir + "/" + indivDir);
    //if (indivFileColl.count === 0) {
    //    Console.PrintLine("=== No individual files to process. Exiting script.")
    //    return;
    //}
    //masterFileColl = TraverseMasterFiles(masterCalibDir + "/" + masterDir);
    
    // Match Masters needing deletion
    //if (masterFileColl.count > 0) {
    //    // need to delete masters
    //} 

    // Do the Maxim collection
    SetCalibrationInMaxim();

    // Delete the individual files
    DeleteIndivFiles(indivCalibDir + "/" + indivDir);

    
	Console.PrintLine ("============ Script completed ==============");
	Console.Logging = false;
	SUP.Terminate();

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
    bdir = root + "\\" + "BrewCalibCollectLog"; // Substitute for directory path
    CreateFolder(bdir);                         // Make sure folder exists
    
    //
    // Construct the file name safely 
    //
    lfn = "BrewCalibCollect-" + Util.FormatVar(new Date().getVarDate(), "yyyymmdd-HhNnSs") + ".log";
    
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








function DeleteIndivFiles(indivDirPath) {
    // Traverse the Individual Files directory looking for Darks/Bias, and Flats
    // When one is found, delete the file
    var filename;
    
    // Loop through the files
    Console.PrintLine(" ");
    Console.PrintLine("============= Deleting Files " + indivDirPath);
    if (CheckFolderExists(indivDirPath)) {
        var folder = FSO.GetFolder(indivDirPath);
        var fc = new Enumerator(folder.files);
        for (; !fc.atEnd() ; fc.moveNext()) {
            filename = fc.item().name;
            if ((filename.substr(0, 4).toUpperCase() === "DARK") || 
                (filename.substr(0, 3).toUpperCase() === "SKY")){
                //Console.PrintLine("Delete IndFile " + filename );
            }
        }
    } else {
        Console.PrintLine("Individual File Folder not found {" + indivDirPath + "}");
    }

    return ;
}

function SetCalibrationInMaxim() {
    // get Maxim Application object
    Console.PrintLine("Create Maxim Application");
    var MaximApp = new ActiveXObject("Maxim.Application");

    var filePaths = indivCalibDir        // + ";" + masterCalibDir;
    Console.PrintLine("Starting CreateCalibrationGroups with " + filePaths);
    try {
        var SIGMACLIPCOMBINE = 3;
        var AUTOOPTIMIZE = 2;
        var numGroups = MaximApp.CreateCalibrationGroups(filePaths, SIGMACLIPCOMBINE, AUTOOPTIMIZE, false);
    }
    catch (ex){
        console.printline("CreateCalibrationGroups failed " + ex.message)
    }
    Console.PrintLine("Calibration groups complete with " + numGroups + " created");
    return;
}


/************
function TraverseIndivFiles(indivDirPath) {
    // Traverse the Individual Files directory looking for Darks/Bias, and Flats
    // When one is found, add it to a list of files meeting its key criteria
    // In the end, I should have a structure like
    //    key DARK,1x1,-10,300,750600
    //    File List
    //          "Dark-10-001-300-1-.fit"
    //          "Dark-10-002-300-1-.fit"
    //            ...
    //    key DARK,1x1,-10,BIAS,750600
    //    File List
    //          "Dark-10-001-Bias-1-.fit"
    //
    var indFileList = new myCollection();
    var key;
    var filename;
    var filesize;

    // Loop through the files
    Console.PrintLine(" " );
    Console.PrintLine("============= Traversing path " + indivDirPath);
    if (CheckFolderExists(indivDirPath)) {
        var folder = FSO.GetFolder(indivDirPath);
        var fc = new Enumerator(folder.files);
        for (; !fc.atEnd() ; fc.moveNext()) {
            filename = fc.item().name;
            filesize = fc.item().size
            if (filename.substr(0, 4).toUpperCase() === "DARK") {
                key = MakeDarkKey(filename, filesize);
                Console.PrintLine("    IndFile " + filename + " size " + filesize);
                Console.PrintLine("    ==>Dark key is " + key);
                indFileList.Add(key, filename);
            }
            else if (filename.substr(0, 3).toUpperCase() === "SKY") {
                key = MakeFlatKey(filename, filesize);
                Console.PrintLine("    IndFile " + filename + " size " + filesize);
                Console.PrintLine("    ==>Flat key is " + key);
                indFileList.Add(key, filename);
            }
            else {
                //Console.PrintLine("    Ignoring file " + filename);
            }
        }
    } else {
        Console.PrintLine("Individual File Folder not found {" + indivDirPath + "}");
        return undefined;
    }
    
    return (indFileList);
}

function TraverseMasterFiles(masterDirPath) {
    // Traverse the Master Files directory looking for Master files
    // When one is found, add it to a list of files meeting its key criteria
    // In the end, I should have a structure like
    //    key MASTER,1x1,-10,300,750600
    //    File List
    //          "Master_Bias 1_3352x2532_Bin1x1_Temp-20C_ExpTime0ms.fit "
    //    key MASTER,2x2,-10,BIAS,750600
    //    File List
    //          "Master_Bias 2_1676x1266_Bin2x2_Temp-10C_ExpTime0ms.fit"
    //
    var indFileList = new myCollection();
    var key;
    var filename;
    var filesize;

    // Loop through the files
    Console.PrintLine(" " );
    Console.PrintLine("========== Traversing Master path " + masterDirPath);
    if (CheckFolderExists(masterDirPath)) {
        var folder = FSO.GetFolder(masterDirPath);
        Console.PrintLine("  set enumerator ");
        var fc = new Enumerator(folder.files);
        Console.PrintLine("  start loop ");
        for (; !fc.atEnd() ; fc.moveNext()) {
            filename = fc.item().name;
            filesize = fc.item().size
            if (filename.substr(0, 6).toUpperCase() === "MASTER") {
                key = MakeMasterKey(filename, filesize);
                Console.PrintLine("    MasterFile " + filename + " size " + filesize );
                Console.PrintLine("    ==>Master key is " + key);
                //indFileList.Add(key, filename);
                }
            else {
                //Console.PrintLine("    Ignoring file " + filename);
            }
        }
    } else {
        Console.PrintLine("Individual File Folder not found {" + indivDirPath + "}");
        return undefined;
    }

    return (indFileList);
}

function MakeDarkKey(filename,filesize)
    // Dark-10-002-300-1-.fit is my dark format
    // key needs to be DARK,-10,1x1,300   which allows split to be used
{
    var type_file, bin, temp, exptime;

    var file_pieces;
    file_pieces = filename.split("-");
    type_file = file_pieces[0];
    type_file = type_file.toUpperCase();
    temp = "-" + file_pieces[1];

    var s = file_pieces[3];   // exposure, or Bias 
    if (s.toUpperCase() === "BIAS") {
        exptime = "BIAS";
    }
    else {
        // should be an exposure time in seconds
        exptime = s;
    }
    var bins = file_pieces[4];
    bin = bins + "x" + bins;

    var key = type_file + "," + bin + "," + temp + "," + exptime + "," + filesize;       // allows use of .split(",")
    return key;
}

function MakeMasterKey(filename, filesize) {
    // Master_Bias 1_324x243_Bin2x2_ExpTime0ms.fit  is my Master Bias format
    // key needs to be DARK,-10,1x1,300   which allows split to be used
    // Master_Dark 10_324x243_Bin2x2_ExpTime3s.fit is Master Dark
    // Master_Flat Blue 10_Blue_3352x2532_Bin1x1_Temp-15C_ExpTime100ms.fit ia Master Flat
    return "MASTER";
}

function MakeFlatKey(filename, filesize) {
    // Sky-Blue-bin1-001.fts  is my Flat format  size 18340
    // key needs to be FLAT,BLUE,1x1,18340
    var pieces;

    var  bin, filter;

    pieces = filename.split("-");
    filter = pieces[1].toUpperCase();

    var s = pieces[2];   // bin1, bin2, ...
    s = s.substr(3,1);
    bin = s + "x" + s;

    var key = "FLAT" + "," + filter + "," + bin + "," + filesize;       // allows use of .split(",")
    return key;
}



var myCollection = function()
    // Implementation of my collection object
    // It make a structure like
    // myCollection:
    //   count;    number of file Classes involved
    //   classCollection:   classObject

    //    key DARK,1x1,-10,300,750600
    //    File List
    //          "Dark-10-001-300-1-.fit"
    //          "Dark-10-002-300-1-.fit"
    //            ...
    //    key DARK,1x1,-10,BIAS,750600
    //    File List
    //          "Dark-10-001-Bias-1-.fit"
{
    this.count = 0;
    this.collection = {};     

    this.Add = function (key, filename)
    {
        var ret = 0;     // returns number of files in the list

        var fileArr = this.collection[key];
        if (fileArr === undefined) {
            // key does not exist
            var newFileArr = [];
            newFileArr[0] = filename;
            this.collection[key] = newFileArr;
            this.count++;
            ret = newFileArr.length;
        } else {
            // key exists; append filename to fileArr
            fileArr[fileArr.length] = filename;
            ret = fileArr.length;
        }
        
        return ret;
    }

    this.Remove = function (key) {
        if (this.collection[key] == undefined)
            return undefined;
        delete this.collection[key]
        return --this.count
    }

    this.Item = function (key) {
        // returns undefined if key does not exist, returns fileArr if key is found
        return this.collection[key];            
    }
}
*************/