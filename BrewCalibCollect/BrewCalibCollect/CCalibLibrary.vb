Imports System
Imports System.IO

Public Class CCalibLibrary

    Private mCalibGroup As Collection     ' collection of CCalibGroup, keyed by FDKey (from MakeKey)

    ' Constructor
    Public Sub New()
        mCalibGroup = New Collection
    End Sub



    Private Sub PutFileInGroup(filename As String)
        Dim group As CCalibGroup
        Dim fd As CFileDescriptor

        fd = New CFileDescriptor(filename)

        ' if the group does not exist, make one
        If (mCalibGroup.Contains(fd.GroupKey)) Then
            group = mCalibGroup.Item(fd.GroupKey)
        Else
            group = New CCalibGroup(fd.GroupKey)
            mCalibGroup.Add(group, fd.GroupKey)
        End If
        group.AddFile(fd)

    End Sub

    Public Sub PopulateGroups(calibPath As String)
        ' Walk the dirctory looking for fits files. For each, assign to the appropriate CalibGroup
        ' Note that we have already done chdir to the correct directory

        Console.WriteLine("Populating groups")

        ' get the fits files in our directory
        Dim di As New IO.DirectoryInfo(calibPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.fit")
        Dim fi As IO.FileInfo

        For Each fi In aryFi
            Console.WriteLine("  " & fi.Name)
            PutFileInGroup(fi.Name)
        Next

        aryFi = di.GetFiles("*.fts")
        For Each fi In aryFi
            Console.WriteLine("  " & fi.Name)
            PutFileInGroup(fi.Name)
        Next
    End Sub

    Public Sub LogGroups()
        ' for testing - write out the group structure to Log file
        Dim group As CCalibGroup

        Console.WriteLine("Logging groups")
        For Each group In mCalibGroup

            group.LogGroup()
        Next

    End Sub

    Public Sub CreateMasters(averageMethod As String)
        ' For each group, create any needed masters
        '   averageMethod is AVG, SIGMA, or MINMAX
        Dim group As CCalibGroup

        Console.WriteLine("Creating masters")
        For Each group In mCalibGroup
            group.CreateMaster(averageMethod)
        Next
    End Sub


    Public Sub RemoveMasterFiles()
        ' remove replaced masters, subs
        Dim group As CCalibGroup

        Console.WriteLine("Removing duplicate masters")
        For Each group In mCalibGroup
            group.RemoveMasterFiles()
        Next
    End Sub

    Public Sub RemoveSubs()
        ' Remove the various sub frames that have been combined into Masters
        Dim group As CCalibGroup

        For Each group In mCalibGroup
            group.RemoveSubs()
        Next

    End Sub
    Public Function CountCombineFrames() As Integer
        Dim frameCount As Integer = 0
        Dim group As CCalibGroup

        For Each group In mCalibGroup
            If (group.SubCount > 1) Then
                frameCount = frameCount + group.SubCount
            End If
        Next
        Return frameCount
    End Function


End Class
