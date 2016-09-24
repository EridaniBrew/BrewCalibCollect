Imports MaxIm

Public Class CCalibGroup
    ' Class containing properties for a single calibration group.
    ' A group is (example) Darks at -10C bin2 300seconds
    '
    ' In theory, each group should only have one Master file. If we have multiple, and we create a new Master from subs,
    ' then all of the old multiple masters will be removed.
    '
    Private mGroupKey As String                    ' the key for this group

    Dim NewMasterDoc As MaxIm.Document            ' a new master file is built here if needed

    Dim ExistingMasterFD As Collection            ' possible multiple existing masters for this group    
    '                                               collection of CFileDescriptor, Keyed by FileName (they should all have the same FileKey)

    Dim SubFD As Collection                       ' collection of CFileDescriptor, keyed by FileName

    Public Sub New(key)
        mGroupKey = key
        ExistingMasterFD = New Collection     ' collection of FD
        SubFD = New Collection                ' collection of FD
    End Sub

    Public ReadOnly Property SubCount() As Integer
        Get
            Return SubFD.Count
        End Get
    End Property

    Public Sub LogGroup()
        Dim myFd As CFileDescriptor
        LogMsg("Group " & mGroupKey)
        LogMsg("  Master Files:")
        For Each myFd In ExistingMasterFD
            LogMsg("    " & myFd.GetFileName())
        Next
        LogMsg("  Sub Files:")
        For Each myFd In SubFD
            LogMsg("    " & myFd.GetFileName())
        Next
        LogMsg("  ")
    End Sub

    Public Sub AddFile(fd As CFileDescriptor)
        '  add fd to the appropriate collection

        If (fd.IsMaster) Then
            ExistingMasterFD.Add(fd, fd.GetFileName)
        Else
            SubFD.Add(fd, fd.GetFileName)
        End If
    End Sub

    Public Sub CreateMaster(averageMethod As MaxIm.CombineType)
        ' averageMethod is AVG, SIGMA, or MINMAX
        'If (SubFD.Count > 0) Then
        '    Dim masterName As String = MakeMasterFileName(mGroupKey)
        '    LogMsg("Creating Master for " & mGroupKey & " as " & masterName)
        '    Console.WriteLine("Creating Master for " & masterName)
        '    Dim imgCollection As Collection = New Collection
        '    Dim fd As CFileDescriptor
        '    For Each fd In SubFD
        '        imgCollection.Add(fd.MaximDoc, fd.GetFileName)
        '    Next
        '    NewMasterDoc = New MaxIm.Document
        '    Dim combineWorked As Boolean = True
        '    ' Different combine methods per file type
        '    Dim combineMethod As MaxIm.CombineType = CombineType.mxAvgCombine
        '    Dim normalize As Boolean = False
        '    If (averageMethod <> 0) Then
        '        ' specified on run string
        '        combineMethod = averageMethod
        '    Else
        '        ' decide based on sub type and number of subs
        '        If ((SubFD.Count > 9) And (fd.IsDark)) Then
        '            combineMethod = MaxIm.CombineType.mcStdDevMaskCombine
        '            normalize = True
        '        ElseIf ((SubFD.Count > 2) And (fd.IsDark)) Then
        '            combineMethod = MaxIm.CombineType.mxMedianCombine
        '        ElseIf (fd.IsDark) Then
        '            combineMethod = MaxIm.CombineType.mxAvgCombine

        '        ElseIf ((SubFD.Count > 9) And (fd.IsBias)) Then
        '            combineMethod = MaxIm.CombineType.mcStdDevMaskCombine
        '            normalize = True
        '        ElseIf (fd.IsBias) Then
        '            combineMethod = MaxIm.CombineType.mxMedianCombine

        '        ElseIf ((SubFD.Count > 9) And (fd.IsFlat)) Then
        '            combineMethod = MaxIm.CombineType.mcStdDevMaskCombine
        '            normalize = True
        '        ElseIf (fd.IsFlat) Then
        '            combineMethod = MaxIm.CombineType.mxMedianCombine
        '            normalize = True
        '        End If
        '    End If
        '    Try
        '        NewMasterDoc.CombineImages(MaxIm.AlignType.mxNoAlignment, False,
        '                              combineMethod,
        '                               normalize,       ' true for flats
        '                              imgCollection)
        '        NewMasterDoc.SaveFile(CurDir() & "/new-" & masterName, MaxIm.ImageFormatType.mxFITS, False, MaxIm.PixelDataFormatType.mxFloatPF)
        '    Catch ex As SystemException
        '        LogMsg("Error creating master for group key " & mGroupKey)
        '        LogMsg("  Error: " & ex.ToString())
        '        combineWorked = False
        '    End Try
        '    ' Bias - needs to add 100 Pedestal, set Pedestal keyword
        '    ' Dark - subtract Bias (how to match?) add 100 pedestal
        '    ' Flat - 
        '    ' Add history comments as well
        '    imgCollection.Clear()
        '    NewMasterDoc.Close()

        '    If (combineWorked) Then
        '        ' Clean up files
        '        ' remove master(s)
        '        For Each fd In ExistingMasterFD
        '            LogMsg("  Removing master " & fd.GetFileName)
        '            fd.RemoveFile()
        '        Next

        '        ' remove subs
        '        For Each fd In SubFD
        '            LogMsg("  Removing sub " & fd.GetFileName)
        '            fd.RemoveFile()
        '            fd.MaximDoc.Close()
        '        Next
        '    End If
        'End If
    End Sub


    Public Sub RemoveMasterFiles()
        ' However, we might have two situations where we need to remove masters
        ' a) we could have multiple masters for unknown reasons. 
        '    Keep the newest one, remove the others (RemoveDupMasters)
        ' b) if we have 2 or more subs, remove all the masters which will be replaced
        '    by new masters later. (RemoveOldMasters)

        RemoveOldMasters()
        RemoveDupMasters()

    End Sub

    Private Sub RemoveOldMasters()
        Dim fd As CFileDescriptor

        If (SubFD.Count > 1) Then
            ' Multiple subs means we willl be generating a master
            ' So delete all the existing ones
            For i = 1 To ExistingMasterFD.Count
                fd = ExistingMasterFD.Item(i)
                LogMsg("Removed Old master " & fd.GetFileName)
                fd.RemoveFile()
            Next
            ExistingMasterFD.Clear()
        End If
    End Sub

    Private Sub RemoveDupMasters()
        Dim fd As CFileDescriptor

        If (ExistingMasterFD.Count > 1) Then
            ' find newest master
            Dim masterFD As CFileDescriptor
            masterFD = ExistingMasterFD.Item(1)
            Dim newestLastModified As Date
            Dim thisLastModified As Date
            newestLastModified = System.IO.File.GetLastWriteTime(CurDir() & "/" & masterFD.GetFileName)
            Dim datestr = newestLastModified.ToLongDateString
            Dim newestIdx As Integer = 1

            Dim i As Integer
            For i = 2 To ExistingMasterFD.Count
                fd = ExistingMasterFD.Item(i)
                thisLastModified = System.IO.File.GetLastWriteTime(CurDir() & "/" & fd.GetFileName)
                datestr = thisLastModified.ToLongDateString
                If (thisLastModified > newestLastModified) Then
                    newestIdx = i
                    newestLastModified = thisLastModified
                End If
            Next
            ' remove master(s)
            For i = 1 To ExistingMasterFD.Count
                fd = ExistingMasterFD.Item(i)
                If (i <> newestIdx) Then
                    LogMsg("Removed duplicate master " & fd.GetFileName)
                    fd.RemoveFile()
                End If
            Next i
        End If

    End Sub

    Public Sub RemoveSubs()
        ' Remove the sub files; they have been converted into masters

        Dim fd As CFileDescriptor

        For Each fd In SubFD
            LogMsg("Removed Subframe " & fd.GetFileName)
            fd.RemoveFile()
        Next
        SubFD.Clear()
    End Sub

    Private Function MakeMasterFileName(groupKey) As String
        ' group key is like DARK|1x1|300|-10|NOFILTER|3352x2532
        ' master name is
        '   Master_Bias 1_3352x2532_Bin1x1_Temp-10C_ExpTime0ms.fit
        '   Master_Dark 1_1676x1266_Bin2x2_Temp-10C_ExpTime600s.fit
        '   Master_Flat Blue 4_Blue_1676x1266_Bin2x2_ExpTime28s.fit
        Dim name As String = ""
        Dim arr() As String = groupKey.split("|")
        name = "Master_"
        If (arr(0) = "DARK") Then
            ' build dark name
            name = name & "Dark 1_" &
                arr(5) & "_" &
                "Bin" & arr(1) & "_" &
                "Temp" & arr(3) & "C_" &
                "ExpTime" & arr(2) & "s.fit"
        ElseIf (arr(0) = "BIAS") Then
            ' build Bias name
            name = name & "Bias 1_" &
                arr(5) & "_" &
                "Bin" & arr(1) & "_" &
                "Temp" & arr(3) & "C_" &
                "ExpTime0ms.fit"
        ElseIf (arr(0) = "FLAT") Then
            ' build Flat name
            name = name & "Flat " & arr(4) & " 1_" & arr(4) & "_" &
                arr(5) & "_" &
                "Bin" & arr(1) & "_" &
                "ExpTime0ms.fit"
        Else
            ' Build weird name?
            name = name & groupKey & ".fit"
        End If

        MakeMasterFileName = name
    End Function
End Class
