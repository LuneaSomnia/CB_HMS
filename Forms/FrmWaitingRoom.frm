VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FrmWaitingRoom 
   Caption         =   "Doctors Waiting Room"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   Icon            =   "FrmWaitingRoom.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtCount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6240
      TabIndex        =   6
      Text            =   "0"
      Top             =   6960
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "List of Observed Patients     "
      TabPicture(0)   =   "FrmWaitingRoom.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Double Click On Patient to Call For Treatment"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   12015
         Begin VSFlex6DAOCtl.vsFlexGrid Grid 
            Height          =   5775
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   10186
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
            BackColor       =   65535
            ForeColor       =   16711680
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   16711680
            BackColorBkg    =   -2147483636
            BackColorAlternate=   49344
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
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   7320
      Width           =   12255
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh Waiting List"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   9720
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label Label22 
      Caption         =   "Number of Patients to be Seen"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   6960
      Width           =   2775
   End
End
Attribute VB_Name = "FrmWaitingRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset
Private Sub FillWaitingRoom()
   ' On Error GoTo ErrorHandler
   Dim lvCount As Integer
   KARI = GlbSysDate
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 10
    Grid.ColAlignment(1) = flexAlignCenterCenter
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "DOCTOR  | CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |BLOOD PRESSURE |WEIGHT | HEIGHT     |VISIT DATE  "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '0'", Conn, adOpenDynamic, adLockOptimistic
        'ADDED INUSE=0 SG 22012011
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND TODOCTORS = '1' AND INUSE = '0' AND COMPLAINS.CARDNUMBER <> '9999/99' AND DISMISSED = 'FALSE'", Conn, adOpenDynamic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        If Trim(!DOCTOR) = "Null" Then lvdoctor = "NONE" Else lvdoctor = !DOCTOR
                        Grid.AddItem lvdoctor & vbTab & !CardNumber & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER & vbTab & !BP & vbTab & !Weight & vbTab & !Height & vbTab & !VisitDate
                        .MoveNext
                        TxtCount = lvCount + 1
                    Wend
                End With
            End If
    Exit Sub
'ErrorHandler:
  '  MsgBox Err.Description

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdRefresh_Click()
    FillWaitingRoom
End Sub

Private Sub Form_Load()
    centerform Me
    FillWaitingRoom
    GlbCurrentForm = EnumDoctors
End Sub

Private Sub Grid_DblClick()
On Error GoTo ErrorHandler
    'COLLECT VALUES
    If Grid.Row = 0 Then Me.Caption = "There are no Patients on the doctor's waiting List.": Exit Sub
    StrDocCardNo = Grid.TextMatrix(Grid.Row, 1)
    StrDocVisitNumber = Grid.TextMatrix(Grid.Row, 2)
    StrDocBMI = Grid.TextMatrix(Grid.Row, 8)
    StrDocVisitDate = Grid.TextMatrix(Grid.Row, 9)
    
    'FLAG RECORD AS IN-USE. SG 22012011
    Conn.Execute "UPDATE COMPLAINS SET INUSE = '1' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
    
    FrmTreatment.Show
    Unload Me
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub


