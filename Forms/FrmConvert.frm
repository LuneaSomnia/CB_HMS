VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmConverter 
   Caption         =   "Lab Result - Scan and Save"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16935
   Icon            =   "FrmConvert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   16935
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkReplace 
      Caption         =   "Replace Captured Image"
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
      Left            =   9000
      TabIndex        =   23
      Top             =   1250
      Width           =   2535
   End
   Begin VB.CommandButton CmdFull 
      Caption         =   "Full Image"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox TxtLabTest 
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   7440
      Width           =   7215
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   7560
      TabIndex        =   7
      Top             =   120
      Width           =   9255
      Begin VB.TextBox TxtVisitNumber 
         Height          =   375
         Left            =   7680
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox TxtCardNumber 
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtNames 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label4 
         Caption         =   "Card Number"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Visit Number"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Client Full Name"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   7560
      TabIndex        =   5
      Top             =   1560
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Height          =   2295
         Left            =   360
         ScaleHeight     =   2235
         ScaleWidth      =   2955
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Image ImgPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DragMode        =   1  'Automatic
         Height          =   5415
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   9015
      End
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
      Left            =   15600
      TabIndex        =   4
      Top             =   1155
      Width           =   1215
   End
   Begin VB.CommandButton CmdReport 
      Caption         =   "Print Report"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   8040
      Width           =   9255
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
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   8640
      Width           =   9255
   End
   Begin VB.CommandButton CmdImageToBinary 
      Caption         =   "Convert Image && Save To Database"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   7440
      Width           =   9255
   End
   Begin Crystal.CrystalReport CrstlRpt 
      Left            =   16080
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CommonDog 
      Left            =   15360
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab TabScan 
      Height          =   7215
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pre Scan"
      TabPicture(0)   =   "FrmConvert.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtCountOne"
      Tab(0).Control(1)=   "GridPreScan"
      Tab(0).Control(2)=   "Label5"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Post Scan"
      TabPicture(1)   =   "FrmConvert.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "GridPostScan"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TxtCountTwo"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox TxtCountTwo 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   5400
         TabIndex        =   21
         Top             =   6720
         Width           =   1695
      End
      Begin VB.TextBox TxtCountOne 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   -69600
         TabIndex        =   19
         Top             =   6720
         Width           =   1695
      End
      Begin VSFlex6DAOCtl.vsFlexGrid GridPreScan 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   10821
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483624
         ForeColor       =   16711680
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   65535
         ForeColorSel    =   16711680
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483624
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   4
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin VSFlex6DAOCtl.vsFlexGrid GridPostScan 
         Height          =   6135
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   10821
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483624
         ForeColor       =   16711680
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   65535
         ForeColorSel    =   16711680
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483624
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   4
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Number of Scanned Items"
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
         Left            =   3000
         TabIndex        =   20
         Top             =   6765
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Lab Requests Pending Result Scan"
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
         Left            =   -72600
         TabIndex        =   18
         Top             =   6765
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
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
      Left            =   12840
      TabIndex        =   3
      Top             =   1250
      Width           =   2535
   End
End
Attribute VB_Name = "FrmConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim Dataconn As New ADODB.Connection
    Dim RsImageStore As New ADODB.Recordset
    Dim RsRecords As New ADODB.Recordset
    Dim RsGrid As New ADODB.Recordset
    Public StrImagePath As String
Private Sub CmdBrowse_Click()
On Error GoTo ERRORHANDLER
    CommonDog.ShowOpen
    StrImagePath = CommonDog.FileName
   'Picture1. = StrImagePath
   ImgPreview.Container = StrImagePath
   ImgPreview.Picture = LoadPicture(StrImagePath)
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdFull_Click()
    On Error GoTo ERRORHANDLER
    FrmFullImage.ImgFull = ImgPreview.Picture
    FrmFullImage.Show
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Description
End Sub

Private Sub CmdImageToBinary_Click()
On Error GoTo ERRORHANDLER
    If TxtCardNumber = "" Then Exit Sub
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM LAB_SCAN", Dataconn, adOpenStatic, adLockOptimistic
    
   'IF IMAGE ALREADY SAVED, THEN THE ASSUMPTION IS THAT ITS BEING REPLACED.
   If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT VISITNUMBER FROM LAB_SCAN WHERE CARDNUMBER = '" & TxtCardNumber & "' and VISITNUMBER = '" & TxtVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        If RsRecords.EOF = False Then
            Resp = MsgBox("A Scan Document has already been saved for this Patient. Do you wish to Replace it?", vbQuestion + vbYesNo)
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
            LoadImageFromFileToDB StrImagePath, RsImageStore, "SCANIMAGE", FileLen(StrImagePath)
            'RsImageStore!PROCESSNUMBER = "001"
        RsImageStore.Update
        
        MsgBox "Conversion and Saving Completed Succesfully", vbInformation
        RsRecords.Close
        POPULATEPreScan
        POPULATEPostScan
        ClearText FrmConverter
    Exit Sub
ERRORHANDLER:
   MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdReport_Click()
On Error GoTo ERRORHANDLER
                        With Me.CrstlRpt
                           .SelectionFormula = ""
                            STRReportName = "PASSPORT PHOTOS ZA WATU"
                           .Connect = Dataconn.ConnectionString
                           .ReportFileName = App.Path & "\REPORTS\Image Report.rpt"
                           .WindowTitle = StrCompanyName & " - " & "DATAFLEX PASSPORT REPORT"
                           .Destination = 0
                           .WindowState = crptMaximized
                           .Action = 1
                        End With
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ERRORHANDLER
    'Set Dataconn = Nothing
    'Dataconn = Conn
    'Dataconn.Open
    POPULATEPreScan
    POPULATEPostScan
    centerform Me
    TabScan.Tab = 0
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
    Exit Sub
    Resume
End Sub

Private Sub POPULATEPreScan()
    On Error GoTo ERRORHANDLER
   Dim Rcount As Integer
   KARI = GlbSysDate
   
    GridPreScan.Clear
    GridPreScan.Rows = 1
    GridPreScan.Cols = 3
    GridPreScan.ColAlignment(1) = flexAlignCenterCenter
    GridPreScan.ColWidth(1) = 3105
    GridPreScan.ColWidth(2) = 99999
    GridPreScan.FormatString = "CARD NUMBER|   PATIENTS FULL NAME  | REQUESTING DOCTOR | VISIT NUMBER | LAB TEST REQUEST DETAILS "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT PATIENT_DETAILS.*,COMPLAINS.LABREQUEST,COMPLAINS.DOCTOR,COMPLAINS.VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND COMPLAINS.TOLABORATORY = '1'", Conn, adOpenDynamic, adLockOptimistic
        RsGrid.Open "SELECT PATIENT_DETAILS.*,COMPLAINS.LABREQUEST,COMPLAINS.DOCTOR,COMPLAINS.VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND (COMPLAINS.TOLABORATORY = '1') AND (COMPLAINS.VisitNumber NOT IN (SELECT VisitNumber From LAB_SCAN)) ", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridPreScan.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & "DR. " & !DOCTOR & vbTab & !VISITNUMBER & vbTab & !LABREQUEST
                        Rcount = Rcount + 1
                        .MoveNext
                        TxtCountOne = Rcount
                    Wend
                End With
            End If
        RsGrid.Close
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub
Private Sub POPULATEPostScan()
    On Error GoTo ERRORHANDLER
   Dim Rcount As Integer
   KARI = GlbSysDate
    GridPostScan.Clear
    GridPostScan.Rows = 1
    GridPostScan.Cols = 3
    GridPostScan.ColAlignment(1) = flexAlignCenterCenter
    GridPostScan.ColWidth(1) = 3105
    GridPostScan.ColWidth(2) = 99999
    GridPostScan.FormatString = "CARD NUMBER|   PATIENTS FULL NAME  | REQUESTING DOCTOR | VISIT NUMBER | LAB TEST REQUEST DETAILS "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT PATIENT_DETAILS.*,COMPLAINS.LABREQUEST,COMPLAINS.DOCTOR,COMPLAINS.VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND COMPLAINS.TOLABORATORY = '1'", Conn, adOpenDynamic, adLockOptimistic
        RsGrid.Open "SELECT PATIENT_DETAILS.*,COMPLAINS.LABREQUEST,COMPLAINS.DOCTOR,COMPLAINS.VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.TOLABORATORY = '1' AND COMPLAINS.VisitNumber IN (SELECT     VisitNumber FROM LAB_SCAN)", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        If Not IsNull(!LABREQUEST) Then
                            GridPostScan.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & "DR. " & !DOCTOR & vbTab & !VISITNUMBER & vbTab & Replace(!LABREQUEST, vbCrLf, ", ")
                        Else
                            GridPostScan.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & "DR. " & !DOCTOR & vbTab & !VISITNUMBER & vbTab & !LABREQUEST
                        End If
                        .MoveNext
                        Rcount = Rcount + 1
                        TxtCountTwo = Rcount
                    Wend
                End With
            End If
        RsGrid.Close
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub GridPostScan_Click()
    On Error GoTo ERRORHANDLER
    TxtLabTest = ""
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM LAB_SCAN WHERE CARDNUMBER = '" & GridPostScan.TextMatrix(GridPostScan.Row, 0) & "' AND VISITNUMBER = '" & GridPostScan.TextMatrix(GridPostScan.Row, 3) & "'", Conn, adOpenStatic, adLockOptimistic
        With RsImageStore
                LoadPictureFromDB RsImageStore, "SCANIMAGE", ImgPreview, d
        End With
    RsImageStore.Close
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Description
End Sub

Private Sub GridPreScan_dblclick()
On Error GoTo ERRORHANDLER
    TxtNames = GridPreScan.TextMatrix(GridPreScan.Row, 1) '& " " & GridPreScan.TextMatrix(GridPreScan.Row, 2) & " " & GridPreScan.TextMatrix(GridPreScan.Row, 3)
    TxtCardNumber = GridPreScan.TextMatrix(GridPreScan.Row, 0)
    TxtLabTest = GridPreScan.TextMatrix(GridPreScan.Row, 4)
    TxtVisitNumber = GridPreScan.TextMatrix(GridPreScan.Row, 3)
    ImgPreview.Picture = LoadPicture(App.Path & "\NoImage.bmp")
    Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description
End Sub


Private Sub ImgPreview_DblClick()
On Error GoTo ERRORHANDLER
    FrmFullImage.ImgFull = LoadPicture(StrImagePath)
    FrmFullImage.Show
Exit Sub
ERRORHANDLER:
    MsgBox Err.Number & " " & Err.Description

End Sub
Public Function LoadPictureFromDB(ByRef rs As ADODB.Recordset, ByVal fldName As String, ByRef Image1 As Object, Optional ByVal strFileName As String)

    On Error GoTo ERRORHANDLER
    
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
        strFileName = "c:\Temp.bmp"
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
ERRORHANDLER:
    MsgBox Err.Description, vbExclamation, "Image Problem"
    'SystemErrorHandler Err.Number, Err.Description
End Function

