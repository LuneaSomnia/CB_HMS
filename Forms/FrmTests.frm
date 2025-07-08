VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmTests 
   Caption         =   "Doctor Tests (Images)"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16095
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   16095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete ATT"
      Height          =   735
      Left            =   13920
      TabIndex        =   16
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ListBox LstAttachments 
      Height          =   1425
      Left            =   13920
      TabIndex        =   14
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox TxtDestPath 
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox TxtFileName 
      Height          =   285
      Left            =   7560
      TabIndex        =   12
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton CmdSaveFile 
      Caption         =   "Save Files"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13920
      TabIndex        =   10
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Type"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   13695
      Begin VB.OptionButton Option1 
         Caption         =   "All "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11760
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptPDFfile 
         Caption         =   "PDF File"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton OptWordDoc 
         Caption         =   "Word Document"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton OptImageFile 
         Caption         =   "Image File"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Full Image"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton CmdImageToBinary 
      Appearance      =   0  'Flat
      Caption         =   " Save To Database"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13920
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   4
      Top             =   8520
      Width           =   2055
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   13695
      Begin VB.PictureBox Picture1 
         Height          =   2295
         Left            =   4920
         ScaleHeight     =   2235
         ScaleWidth      =   2955
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Image ImgPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DragMode        =   1  'Automatic
         Height          =   7335
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   13455
      End
   End
   Begin MSComDlg.CommonDialog CommonDog 
      Left            =   6480
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Browse File From Location"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "FrmTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const MAX_PATH = 260

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SetCurrentDirectory Lib "kernel32" _
    Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Private Declare Function GetCurrentDirectory Lib "kernel32" _
    Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   Alias "SHGetPathFromIDListA" _
  (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

'Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, ByVal lpBuffer As String) As Long
        
Private Declare Function lstrcat Lib "kernel32" _
    Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    
Dim RsRecords As New ADODB.Recordset
Dim RsImageStore As New ADODB.Recordset
Public StrImagePath As String

Private Sub CmdBrowse_Click()
On Error GoTo Errorhandler
    CommonDog.ShowOpen
    StrImagePath = CommonDog.FileName
   'Picture1. = StrImagePath
   ImgPreview.Container = StrImagePath
        If OptImageFile.Value = True Then
            ImgPreview.Picture = LoadPicture(StrImagePath)
        Else
            With CommonDog
'                .CancelError = True
'                .Flags = cdlOFNExplorer
'                .ShowOpen
                If Not .FileName = "" Then
                    TxtFileName = .FileName
                    TxtFileName = Mid(Trim(TxtFileName), InStrRev(Trim(TxtFileName), "\") + 1)
                Else
                    TxtFileName = "Select Source File..."
                    TxtFileName = ""
                End If
            End With
        
        End If
   GlbTestSourceLocation = StrImagePath
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
    Exit Sub
    Resume
End Sub

Private Sub CmdDelete_Click()
            Resp = MsgBox("Are you sure you wish to Delete this File?", vbQuestion + vbYesNo, "Data Vault Systems")
                If Resp = vbYes Then
                    Conn.Execute "DELETE FROM  TEST_SCAN WHERE SCANID = '" & GetID_NameFromCombo(LstAttachments.Text, 2) & "'"
                End If
            LoadAttachmentList
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Uncheck()
    
End Sub
Private Sub CmdImageToBinary_Click()
On Error GoTo Errorhandler
    TxtCardNumber = StrDocCardNo
    TxtVisitNumber = StrDocVisitNumber
    
    If OptWordDoc.Value = True Then
            'GoTo SavePhysicalFiles
            MsgBox "Saving Word Doc Completed Succesfully", vbInformation, "Data Vault Systems"
        Exit Sub
    ElseIf OptPDFfile.Value = True Then
            'GoTo SavePhysicalFiles
            MsgBox "Saving PDF file Completed Succesfully", vbInformation, "Data Vault Systems"
        Exit Sub
    ElseIf OptImageFile.Value = True Then
    
    End If
    'StrImagePath = d
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM TEST_SCAN where IMAGETYPE = '" & GlbTestImageType & "'", Conn, adOpenStatic, adLockOptimistic
    
   'IF IMAGE ALREADY SAVED, THEN THE ASSUMPTION IS THAT ITS BEING REPLACED.
   If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT VISITNUMBER FROM TEST_SCAN WHERE CARDNUMBER = '" & TxtCardNumber & "' and VISITNUMBER = '" & TxtVisitNumber & "' and IMAGETYPE = '" & GlbTestImageType & "'", Conn, adOpenStatic, adLockOptimistic
        If RsRecords.EOF = False Then
            Resp = MsgBox("A Scan Document has already been saved for this Patient. Do you wish to Replace it?", vbQuestion + vbYesNo, "Data Vault Systems")
                If Resp = vbYes Then
                    GoTo Updating
                Else
                    Exit Sub
                End If
        End If
        
        RsImageStore.AddNew
Updating:
            RsImageStore!CardNumber = TxtCardNumber
            RsImageStore!VISITNUMBER = TxtVisitNumber
            RsImageStore!ImageType = GlbTestImageType
            LoadImageFromFileToDB StrImagePath, RsImageStore, "SCANIMAGE", FileLen(StrImagePath)
            'RsImageStore!PROCESSNUMBER = "001"
        RsImageStore.Update
        
        MsgBox "Conversion and Saving Completed Succesfully", vbInformation, "Data Vault Systems"
        RsRecords.Close
        'POPULATEPreScan
        'POPULATEPostScan
        'ClearText FrmConverter
    Exit Sub
Errorhandler:
   MsgBox Err.Number & " " & Err.Description
   Exit Sub
   Resume
End Sub
Private Sub SavePhysicalFiles(SourceFileLocation, Destination)
On Error GoTo Errorhandler
        TxtDestPath = Destination
    If Not Dir(Trim(SourceFileLocation)) = "" Then
        'If Not Dir(Trim(Destination), vbDirectory) = "" Then
            If Not Right(Trim(Destination), 1) = "\" Then
                TxtDestPath = Trim(Destination) & "\"
            End If
            Dim destFile As String
            destFile = TxtDestPath & Trim(TxtFileName)
            
       ' Else
        '    MsgBox "Please select destination folder.", vbExclamation, "Missing Destination Folder"
        'End If
        
        '***********************************
            FileCopy Trim(SourceFileLocation), destFile
            MsgBox "File's copied succesfully", vbInformation, "Data Vault Systems"
            
            GlbTestDestinationLocation = destFile
            'LOAD FILE USING DEFAULT APPLICATION
            Debug.Print ShellExecute(hWnd, "open", " & destFile & ", vbNullString, vbNullString, 1)
        '***************************************
        
    Else
        MsgBox "Please select source file.", vbExclamation, "Missing Source File"
    End If
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Data Vault System Ltd"
       Exit Sub
       Resume
    Resume
End Sub

Private Sub CmdSaveFile_Click()
On Error GoTo Errorhandler
   'IF IMAGE ALREADY SAVED, THEN THE ASSUMPTION IS THAT ITS BEING REPLACED.
   If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT VISITNUMBER FROM TEST_SCAN WHERE CARDNUMBER = '" & StrDocCardNo & "' and VISITNUMBER = '" & GlbImageVisitNumber & "' and IMAGETYPE = '" & GlbTestImageType & "'", Conn, adOpenStatic, adLockOptimistic
        If RsRecords.EOF = False Then
            Resp = MsgBox("A Test Document has already been saved for this Patient. Do you wish to Replace it?", vbQuestion + vbYesNo, "Data Vault Systems")
                If Resp = vbYes Then
                    Conn.Execute "DELETE FROM TEST_SCAN WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & GlbImageVisitNumber & "' AND IMAGETYPE = '" & GlbTestImageType & "'"
                    GoTo Updating
                Else
                    Resp = MsgBox("Do you wish to add another document?", vbQuestion + vbYesNo, "Data Vault Systems")
                        If Resp = vbYes Then
                            GoTo Updating
                        End If
                End If
        End If
       
Updating:
    GlbTestDestinationLocation = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'TestFilesLocation'")
    GlbTestDestinationLocation = GlbTestDestinationLocation & "\" & TxtFileName
        Conn.Execute "INSERT INTO TEST_SCAN (CARDNUMBER,VISITNUMBER,IMAGETYPE,IMAGELOCATION)" & _
                     "VALUES('" & StrDocCardNo & "','" & GlbImageVisitNumber & "','" & GlbTestImageType & "','" & GlbTestDestinationLocation & "')"
            
      
    'SAVE IMAGE
    GlbTestDestinationLocation = FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'TestFilesLocation'")
    SavePhysicalFiles GlbTestSourceLocation, GlbTestDestinationLocation
        
        'MsgBox "Saving Completed Succesfully", vbInformation, "Data Vault Systems Ltd"
        RsRecords.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Image Problem"
       Exit Sub
    Resume
End Sub

Private Sub Form_Load()
On Error GoTo Errorhandler
    centerform Me
    
    LoadTestImage
    LoadAttachmentList
    
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Data Vault System Ltd"
       Exit Sub
    Resume
End Sub
Private Sub LoadAttachmentList()
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM TEST_SCAN WHERE CARDNUMBER = '" & StrDocCardNo & "' and VISITNUMBER = '" & StrDocVisitNumber & "' and IMAGETYPE = '" & GlbTestImageType & "'", Conn, adOpenForwardOnly, adLockOptimistic
        With RsImageStore
            While .EOF = False
                LstAttachments.AddItem "ATT - " & !SCANID
                .MoveNext
            Wend
        End With
End Sub
Private Sub LoadTestImage()
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    'CONFIRM ON WHICH VISIT TO SAVE TEST IMAGE
        If FrmTreatment.ChkShowingHistory.Value = 1 Then
            RsImageStore.Open "SELECT * FROM TEST_SCAN WHERE CARDNUMBER = '" & StrDocCardNo & "' and VISITNUMBER = '" & GlbImageVisitNumber & "' and IMAGETYPE = '" & GlbTestImageType & "'", Conn, adOpenForwardOnly, adLockOptimistic
        Else
            RsImageStore.Open "SELECT * FROM TEST_SCAN WHERE CARDNUMBER = '" & StrDocCardNo & "' and VISITNUMBER = '" & StrDocVisitNumber & "' and IMAGETYPE = '" & GlbTestImageType & "'", Conn, adOpenForwardOnly, adLockOptimistic
        End If
            With RsImageStore
                If .EOF = False Then
                    'LOAD IMAGE
                        If OptImageFile.Value = True Then
                            LoadPictureFromDB RsImageStore, "SCANIMAGE", ImgPreview, D
                        End If
                        
                    GlbTestDestinationLocation = !ImageLocation
                    Debug.Print ShellExecute(hWnd, "open", GlbTestDestinationLocation, vbNullString, vbNullString, 1)
                End If
            End With
    RsImageStore.Close
End Sub
Public Function LoadPictureFromDB(ByRef rs As ADODB.Recordset, ByVal fldName As String, ByRef Image1 As Object, Optional ByVal strFileName As String)

    On Error GoTo Errorhandler
    
    'If Recordset is Empty, Then Exit
    If rs Is Nothing Then
        'GoTo procNoPicture
    End If
    
    Dim strTempFileName As String
    'strTempFileName = GetParam("IMAGETEMPFILE")
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    
    strStream.Write rs.Fields(fldName).Value

    If strFileName = "" Then
        strFileName = "c:\Temp\Temp.bmp"
    End If
    strStream.SaveToFile strFileName, adSaveCreateOverWrite
    ImgPreview.Picture = LoadPicture(strFileName)
    On Error Resume Next
    'Image1.DisplayBlankImage Image1.Width, Image1.Height, 800, 600, 2
    'Image1.Display
    If Err.Number <> 0 Then
        'Image1.DisplayBlankImage Image1.Width, Image1.Height, 800, 600, 2
        'Image1.Refresh
        ImgPreview.Picture = LoadPicture(App.Path & "\NoImage.bmp")
        LoadPictureFromDB = False
    Else
        LoadPictureFromDB = True
    End If
    On Error GoTo 0
    'Kill ("C:\Temp.tif")
    'LoadPictureFromDB = True

Exit Function
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Image Problem"
    'SystemErrorHandler Err.Number, Err.Description
       Exit Function
    Resume
End Function

Private Sub btnSelectFolder_Click()
On Error GoTo Errorhandler
'===================================
Dim lRet As Long
Dim sBuffer As String
Dim sTitle As String
Dim tBrowseInfo As BrowseInfo
Dim sCurDir As String
Dim lPidl As Long

    sTitle = "Select Destination Folder"
    
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
                   BIF_EDITBOX Or BIF_VALIDATE Or BIF_NEWDIALOGSTYLE
    End With
    
    lRet = SHBrowseForFolder(tBrowseInfo)
    
    If lRet > 0 Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lRet, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        txtDestFolder.Text = sBuffer
    End If
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Data Vault System Ltd"
       Exit Sub
    Resume
End Sub


Private Sub LstAttachments_DblClick()
    
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM TEST_SCAN WHERE CARDNUMBER = '" & StrDocCardNo & "' and VISITNUMBER = '" & GlbImageVisitNumber & "' and IMAGETYPE = '" & GlbTestImageType & "' AND SCANID = '" & GetID_NameFromCombo(LstAttachments.Text, 2) & "'", Conn, adOpenForwardOnly, adLockOptimistic
            With RsImageStore
                If .EOF = False Then
                    'LOAD IMAGE
                        If OptImageFile.Value = True Then
                            LoadPictureFromDB RsImageStore, "SCANIMAGE", ImgPreview, D
                        End If
                        
                    GlbTestDestinationLocation = !ImageLocation
                    Debug.Print ShellExecute(hWnd, "open", GlbTestDestinationLocation, vbNullString, vbNullString, 1)
                End If
            End With
    RsImageStore.Close
End Sub
