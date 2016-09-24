Imports System.IO

Module Main
    '------------------------------------------------------------------------------
    '
    ' Script:       BrewCalibCollect.js 
    ' Author:       Robert Brewington
    '
    '  ACP Script to collect whatever darks, bias, And flats are in a directory And create the appropriate Master files for
    '  MaximDL in the Calibration process.
    '
    '  Overall process:
    '    1. Scan the folder, build the object structure describing the files. Both Masters And Subs are in the structure.
    '
    '    2. Create the Master files for the discovered subs.
    '
    '    3. Remove the subs And old masters being replaced. Save the newly created masters. For now, removing means moving the files into a 
    '       subfolder "Removed". When things seems to work well, will delete the files instead.
    '
    '    4. Run the Maxim routine to replace the existing masters in Maxim. This Is the equivalent of the  Auto-Generate (Clear Old) button in Maxim.
    '       We has to do step 2 because Maxim does Not provide the routine equivalent to the Replace with Masters button.
    '
    '  Log file lists information about the run.
    '
    '  Run string:
    '      BrewCalibCollect.exe -opt calib_folder_path
    '      See WriteUsage for options
    '
    '	NOTE:
    '		Other users need to change the variables below to point to the
    '		folder where the calibration files are found.
    '      You would also need to verify / change the logic which determines the parameters of
    '      each file. My naming conventions may be different.
    '		
    ' Version:     
    ' 1.0.0   - Initial version

    ' This Is where we find the Flat, Bias, Dark, And previous Master files to be added to the Calibration sets.
    ' You need to change this to whatever your folder is
    Private calibDir As String = ""

    Dim FSO
    Dim logObjWriter As StreamWriter

    Sub Main(ByVal sArgs() As String)

        ' Check run time parameters.
        ' Need the target calibration directory
        ' Optional Average technique
        ' BrewCalibCollect.exe /a:Avg "C:\Users\erida\Dropbox\BrewSky\Programs\BrewCalibCollect\STF8300"
        If (sArgs.Length = 0) Then
            Console.WriteLine("Missing input parameters")
            WriteUsage()
            Exit Sub
        End If

        Dim i As Integer
        Dim averageMethod As MaxIm.CombineType = 0
        For i = 0 To sArgs.Length - 2
            ' look for options
            If (UCase(sArgs(i)) = "-AVG") Then
                averageMethod = MaxIm.CombineType.mxAvgCombine
            ElseIf (UCase(sArgs(i)) = "-MEDIAN") Then
                averageMethod = MaxIm.CombineType.mxMedianCombine
            ElseIf (UCase(sArgs(i)) = "-SIGMA") Then
                averageMethod = MaxIm.CombineType.mxSigmaClipCombine
            ElseIf (UCase(sArgs(i)) = "-SDMASK") Then
                averageMethod = MaxIm.CombineType.mcStdDevMaskCombine
            Else
                Console.WriteLine("Invalid option " & sArgs(i))
                WriteUsage()
                Exit Sub
            End If
        Next i

        ' Last parameter is Calibration folder
        calibDir = sArgs(sArgs.Length - 1)
        ' Check valid folder
        If (Not System.IO.Directory.Exists(calibDir)) Then
            Console.WriteLine("Invalid calibration folder " & calibDir)
            WriteUsage()
            Exit Sub
        End If

        'FSO = CreateObject("Scripting.FileSystemObject")
        Dim CalibLibrary As CCalibLibrary = New CCalibLibrary

        ChDir(calibDir)

        ' Create Log File
        Dim logFileName As String = "BrewCalibCollectLog.txt"
        logObjWriter = New System.IO.StreamWriter(logFileName, False)
        LogMsg(" ")
        LogMsg("==========================================")
        LogMsg("Starting calibration collection at " & Now().ToString())
        LogMsg("Calibration folder is " & calibDir)
        LogMsg("Average method is " & averageMethod)


        'Traverse the calibDir folder looking for files
        CalibLibrary.PopulateGroups(calibDir)
        CalibLibrary.LogGroups()

        'CalibLibrary.CreateMasters(averageMethod)


        ' Delete the individual files
        CalibLibrary.RemoveMasterFiles()
        Dim numFrames As Integer = CalibLibrary.CountCombineFrames()

        ' Close subs in Maxim
        Dim MaximApp As MaxIm.Application = New MaxIm.Application()
        MaximApp.CloseAll()
        MaximApp = Nothing

        ' Do the Maxim SetCalibration
        SetCalibrationInMaxim(numFrames)

        CalibLibrary.RemoveSubs()

        LogMsg("============ Script completed ==============")
        logObjWriter.close()
    End Sub

    Public Sub LogMsg(s As String)
        logObjWriter.WriteLine(s)
        logObjWriter.Flush()
    End Sub

    Public Sub WriteUsage()
        Dim s As String
        Console.WriteLine("Usage: BrewCalibCollect.exe -opt calib_folder_path")
        Console.WriteLine(" ")
        Console.WriteLine("   opt    is averaging method")
        Console.WriteLine("          avg - average pixels")
        Console.WriteLine("          median - median value of pixels")
        Console.WriteLine("          sigma - average after eliminating pixels beyond 3 sigma")
        Console.WriteLine("          sdmask - special Maxim case of sigma rejection")
        Console.WriteLine(" ")
        Console.WriteLine("   calib_folder_path is the folder containing Darks, Flats, Bias, and Master files")
        Console.WriteLine(" ")
        Console.WriteLine("Press enter to continue...")
        s = Console.ReadLine()

    End Sub

    Private Sub ObsoleteSetCalibrationInMaxim()
        Console.WriteLine("Maxim SetCalibration of masters")
        Dim MaximApp As MaxIm.Application = New MaxIm.Application()
        Dim numGroups As Integer = 0

        Dim filePaths As String = CurDir()
        Try
            numGroups = MaximApp.CreateCalibrationGroups(filePaths, MaxIm.CombineType.mcStdDevMaskCombine, MaxIm.DarkFrameScalingType.dfsAutoOptimize, False)

        Catch ex As SystemException
            Console.WriteLine("SetCalibrationInMaxim failed " + ex.ToString())
            LogMsg("SetCalibrationInMaxim failed " + ex.ToString())
        End Try

        LogMsg("Calibration groups complete with " + CStr(numGroups) + " created")
        MaximApp = Nothing
    End Sub

    Private Sub SetCalibrationInMaxim(numFrames)
        ' Run an AutiIt script to have Maxim hit the two buttons
        ' a) Autogenerate (clear old)
        '    This task can be done by scripting Maxim
        ' b) Replace with Masters
        '    This task does not appear to be scriptable:(
        Dim start_info As System.Diagnostics.ProcessStartInfo =
            New System.Diagnostics.ProcessStartInfo("C:\Users\erida\Dropbox\BrewSky\Programs\BrewCalibCollect\BrewCalibCollect\BrewCalibCollect\MaximButtons.exe",
                                                    CStr(numFrames))
        start_info.UseShellExecute = True
        start_info.CreateNoWindow = True

        ' Make the process and set its start information.
        Dim proc As New Process
        proc.StartInfo = start_info

        ' Start the process, wait for AutoIt to complete
        LogMsg("Starting AutoIt with " & numFrames & " frames")
        proc.Start()

        Dim timeout As Integer = numFrames * 3 * 1000 '1 minute in milliseconds

        If Not proc.WaitForExit(timeout) Then
            ' AutoIt did not return in time
            LogMsg("AutoIt timed out")
        End If
    End Sub
End Module
