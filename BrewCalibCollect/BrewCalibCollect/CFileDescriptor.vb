Public Class CFileDescriptor
    Private Const DARKFILE = "Dark"
    Private Const BIASFILE = "Bias"
    Private Const FLATFILE = "Flat"

    Private mType As String             ' Dark/Bias/Flat

    Private Const SUBFILE = "Sub"
    Private Const MASTERFILE = "Master"
    Private mSub As String              ' Sub/Master

    Private mTemp As String            ' Dark, Bias frame temperature
    Private mBin As String              ' binning 1x1, 2x2, ...
    Private mFilter As String           ' Flat file filter
    Private mExposure As String

    Private mFilename As String
    Private mGroupKey As String

    'Private mMaximApp As MaxIm.Application = Nothing
    Private mMaximDoc As MaxIm.Document = Nothing       ' Maxim document containg the Fits file contents. Only loaded if needed

    ' Constructor
    Public Sub New(filename As String)
        ' Need to parse fileName/groupKey to determine these properties
        '   groupKey is like DARK|1x1|300|-10|NONE
        mFilename = filename

        'If (IsNothing(mMaximApp)) Then
        'mMaximApp = New MaxIm.Application()
        'End If

        If (InStr(filename, "Master") > 0) Then
            mSub = MASTERFILE
            mMaximDoc = Nothing
        Else
            mSub = SUBFILE
            LoadMaximDoc()
        End If

        MakeGroupKey(filename)
        If (Not IsNothing(mMaximDoc)) Then
            mMaximDoc.Close()
        End If
    End Sub

    Public ReadOnly Property GetFDType() As String
        Get
            Return mType
        End Get
    End Property

    Public Property GroupKey() As String
        Get
            Return mGroupKey
        End Get
        Set(key As String)
            mGroupKey = key
        End Set
    End Property

    'Public ReadOnly Property MaximDoc() As MaxIm.Document
    '    Get
    '        Return mMaximDoc
    '    End Get
    'End Property

    Public ReadOnly Property IsDark() As Boolean
        Get
            If mType = DARKFILE Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsFlat() As Boolean
        Get
            If mType = FLATFILE Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsBias() As Boolean
        Get
            If mType = BIASFILE Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property IsMaster() As Boolean
        Get
            If mSub = MASTERFILE Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property GetFileName() As String
        Get
            Return mFilename
        End Get
    End Property

    Public ReadOnly Property ImageDimensions() As String
        Get
            Dim myDimension As String = ""
            If (Not IsNothing(mMaximDoc)) Then
                myDimension = FitsHeader("NAXIS1") & "x" & FitsHeader("NAXIS2")
            End If
            Return myDimension
        End Get
    End Property


    Private Function FitsTemperature() As String
        ' note may be no SET_TEMP fits header
        ' Also, need to truncate string -15.0000 -> -15
        Dim temp As String = ""
        If (Not IsNothing(mMaximDoc)) Then
            temp = FitsHeader("SET-TEMP")
        End If
        FitsTemperature = temp
    End Function

    Public Sub RemoveFile()
        ' Remove the file
        ' For now, move the file to the Remove folder
        If (Not System.IO.Directory.Exists(CurDir() & "/Remove")) Then
            System.IO.Directory.CreateDirectory(CurDir() & "/Remove")
        End If
        Try
            System.IO.File.Move(CurDir() & "/" & GetFileName,
                            CurDir() & "/Remove/" & GetFileName)
        Catch ex As SystemException
            LogMsg("Remove failed for file " & GetFileName)
        End Try
    End Sub

    Public Sub LoadMaximDoc()
        mMaximDoc = New MaxIm.Document()
        mMaximDoc.OpenFile(CurDir() & "/" & mFilename)
    End Sub


    Private Function FitsHeader(fitskey As String) As String
        Dim val As String = ""
        val = mMaximDoc.GetFITSKey(fitskey)
        Return val
    End Function




    Private Function MakeSubDarkKey() As String
        '  Dark-10-002-300-1-.fit is my dark format
        '     No filter
        mType = DARKFILE
        mFilter = "NOFILTER"
        mTemp = "NOTEMP"
        Dim dimension As String
        Dim arr As String() = mFilename.Split("-")
        Dim theKey As String
        If (arr(4) = "1") Then
            mBin = "1x1"
        ElseIf (arr(4) = "2") Then
            mBin = "2x2"
        Else
            mBin = "NOBIN"
        End If
        mExposure = arr(3)
        mTemp = arr(1)
        If (mTemp <> "0") Then
            mTemp = "-" & mTemp
        End If
        dimension = ImageDimensions

        theKey = "DARK|" & mBin & "|" & mExposure & "|" & mTemp & "|NOFILTER|" & dimension
        MakeSubDarkKey = theKey
    End Function

    Private Function MakeSubFlatKey() As String
        '  Sky-Blue-bin1-001.fts is my flat format
        '     No exposure, temp is in file
        mType = FLATFILE
        mTemp = FitsTemperature()
        mExposure = "NOEXP"
        Dim dimension As String
        Dim arr As String() = mFilename.Split("-")
        Dim theKey As String
        If (arr(2) = "bin1") Then
            mBin = "1x1"
        ElseIf (arr(2) = "bin2") Then
            mBin = "2x2"
        Else
            mBin = "NOBIN"
        End If
        mFilter = UCase(arr(1))
        dimension = ImageDimensions

        theKey = "FLAT|" & mBin & "|NOEXP|" & mTemp & "|" & mFilter & "|" & dimension
        MakeSubFlatKey = theKey

    End Function

    Private Function MakeSubBiasKey() As String
        ' dark-10-014-Bias-1-.fit
        '     No filter or exposure
        mType = BIASFILE
        mFilter = "NOFILTER"
        mExposure = "NOEXP"
        Dim dimension As String
        Dim arr As String() = mFilename.Split("-")
        Dim theKey As String
        If (arr(4) = "1") Then
            mBin = "1x1"
        ElseIf (arr(4) = "2") Then
            mBin = "2x2"
        Else
            mBin = "NOBIN"
        End If
        mTemp = arr(1)
        If (mTemp <> "0") Then
            mTemp = "-" & mTemp
        End If
        dimension = ImageDimensions

        theKey = "BIAS|" & mBin & "|NOEXP|" & mTemp & "|NOFILTER|" & dimension
        MakeSubBiasKey = theKey

    End Function

    Private Function MakeMasterDarkKey() As String
        '  Master_Dark 1_1676x1266_Bin2x2_Temp-10C_ExpTime600s.fit is my dark format
        '  May not have Temperature if camera does not cool! (STi)
        '     No filter
        mType = DARKFILE
        mFilter = "NOFILTER"
        Dim dimension As String
        Dim arr As String() = mFilename.Split("_")
        Dim theKey As String
        If (arr(3) = "Bin1x1") Then
            mBin = "1x1"
        ElseIf (arr(3) = "Bin2x2") Then
            mBin = "2x2"
        Else
            mBin = "NOBIN"
        End If

        mExposure = mFilename.Substring(InStr(mFilename, "ExpTime") - 1)
        Dim posS As Integer = InStr(mExposure, "s")
        mExposure = mExposure.Substring(7, posS - 8)
        If (InStr(mFilename, "Temp") > 0) Then
            mTemp = mFilename.Substring(InStr(mFilename, "Temp") - 1)
            mTemp = mTemp.Substring(4, InStr(mTemp, "C") - 5)
        Else
            mTemp = "NOTEMP"                ' no temperature on this one
        End If
        dimension = arr(2)

        theKey = "DARK|" & mBin & "|" & mExposure & "|" & mTemp & "|NOFILTER|" & dimension
        MakeMasterDarkKey = theKey

    End Function

    Private Function MakeMasterFlatKey() As String
        ' Master_Flat Blue 2_Blue_1676x1266_Bin2x2_Temp-10C_ExpTime100ms
        '     No exposure 
        ' Apparently Master file may or may not have a temperature.
        ' Seems to be based on whether camera has temperature; maybe sometimes I forget to
        ' enable coolers?
        mType = FLATFILE
        mTemp = "NOTEMP"
        Dim arr As String() = mFilename.Split("_")
        If (InStr(mFilename, "Temp") > 0) Then
            mTemp = arr(5).Substring(4)
            mTemp = mTemp.Remove(mTemp.Length - 1)
        End If
        mExposure = "NOEXP"
        Dim dimension As String
        Dim theKey As String
        If (arr(4) = "Bin1x1") Then
            mBin = "1x1"
        ElseIf (arr(4) = "Bin2x2") Then
            mBin = "2x2"
        Else
            mBin = "NOBIN"
        End If
        mFilter = UCase(arr(2))
        dimension = arr(3)

        theKey = "FLAT|" & mBin & "|NOEXP|" & mTemp & "|" & mFilter & "|" & dimension
        MakeMasterFlatKey = theKey
    End Function

    Private Function MakeMasterBiasKey() As String
        ' Master_Bias 1_3352x2532_Bin1x1_Temp-10C_ExpTime0ms
        '     No filter or exposure        mType = DARKFILE
        mType = BIASFILE
        mExposure = "NOEXP"
        mFilter = "NOFILTER"

        Dim arr As String() = mFilename.Split("_")
        Dim theKey As String
        If (arr(3) = "Bin1x1") Then
            mBin = "1x1"
        ElseIf (arr(3) = "Bin2x2") Then
            mBin = "2x2"
        Else
            mBin = "NOBIN"
        End If
        If (InStr(mFilename, "Temp") > 0) Then
            mTemp = arr(4).Substring(4)
            mTemp = mTemp.Substring(0, InStr(mTemp, "C") - 1)
        Else
            mTemp = "NOTEMP"
        End If
        Dim dimension As String = arr(2)

        theKey = "BIAS|" & mBin & "|NOEXP|" & mTemp & "|NOFILTER|" & dimension
        MakeMasterBiasKey = theKey

    End Function


    Private Sub MakeGroupKey(filename As String)
        ' Using the filename, make a key of the form
        '    MasterSub|Type|bin|Exposure|Temperature|Filter
        ' Examples:
        '    MASTER|DARK|1x1|300|-10|NONE    No filter
        '    SUB|BIAS|1x1|0|-10|NONE         No filter or exposure
        '    SUB|FLAT|2x2|0|0|BLUE           No exposure or temperature
        '
        '     key needs to be DARK,-10,1x1,300   which allows split to be used

        Dim theGroupKey As String = ""

        If (UCase(filename.Substring(0, 4)) = "DARK") Then
            If (InStr(filename, "Bias") > 0) Then
                theGroupKey = MakeSubBiasKey()
            Else
                theGroupKey = MakeSubDarkKey()
            End If
        ElseIf (UCase(filename.Substring(0, 3)) = "SKY") Then
            theGroupKey = MakeSubFlatKey()
        ElseIf (UCase(filename.Substring(0, 5)) = "PANEL") Then
            theGroupKey = MakeSubFlatKey()
        ElseIf (UCase(filename.Substring(0, 6)) = "MASTER") Then
            ' Master, now what type?
            If (UCase(filename.Substring(7, 4)) = "DARK") Then
                theGroupKey = MakeMasterDarkKey()
            ElseIf (UCase(filename.Substring(7, 4)) = "BIAS") Then
                theGroupKey = MakeMasterBiasKey()
            ElseIf (UCase(filename.Substring(7, 4)) = "FLAT") Then
                theGroupKey = MakeMasterFlatKey()
            Else
                ' unknown master type?
                LogMsg(">>>> Unknown Master file type for file " & filename)
            End If
        Else
            LogMsg(">>>> Unknown Sub file type for file " & filename)
        End If
        'LogMsg("file " & filename & " gave key " & theKey)
        mGroupKey = theGroupKey

    End Sub
End Class
