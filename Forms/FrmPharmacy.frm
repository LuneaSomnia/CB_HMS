VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPharmacy 
   Caption         =   "Pharmacy"
   ClientHeight    =   9225
   ClientLeft      =   5670
   ClientTop       =   2970
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Garamond"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPharmacy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   9975
   Begin TabDlg.SSTab TabPharmacy 
      Height          =   3135
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "List of Patients Waiting for Drugs"
      TabPicture(0)   =   "FrmPharmacy.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "List of Patients Collecting Previously Unpaid Drugs"
      TabPicture(1)   =   "FrmPharmacy.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Malipo ya Pole Pole."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2655
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   9495
         Begin VSFlex6DAOCtl.vsFlexGrid GridCredit 
            Height          =   2295
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4048
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
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2655
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   9495
         Begin VSFlex6DAOCtl.vsFlexGrid Grid 
            Height          =   2295
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4048
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
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Post Patient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   7440
      Width           =   9735
      Begin VB.CommandButton CmdPost 
         Caption         =   "POST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptConsultation 
         Caption         =   "TO CONSULTATION"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   280
         Width           =   2055
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO OBSERVATION"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   20
         Top             =   280
         Width           =   1935
      End
      Begin VB.OptionButton OptDoctors 
         Caption         =   "TO DOCTORS"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   19
         Top             =   280
         Width           =   1455
      End
      Begin VB.OptionButton OptCashier 
         Caption         =   "TO CASHIERS"
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   6360
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport CrstlRpt 
      Left            =   4440
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   8280
      Width           =   9735
      Begin VB.CommandButton CMDPrintPrescription 
         Caption         =   "Print Prescription"
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
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh List"
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
         TabIndex        =   14
         Top             =   240
         Width           =   1695
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
         Left            =   7680
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdSubmit 
         Caption         =   "Drugs Submitted"
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
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drug Prescription"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.ListBox LstPrescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   2520
         Width           =   9495
      End
      Begin VB.TextBox TxtPrescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2880
         Width           =   9495
      End
      Begin VB.TextBox TxtDiagnosis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1560
         Width           =   9495
      End
      Begin VB.TextBox TxtDocName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   6
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox TxtPatientName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox TxtCardNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   9615
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9615
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "Doctor's Prescription"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Doctor's Diagnosis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Doctor's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Patient's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Card Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmPharmacy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrid As New ADODB.Recordset
Dim rsfill As New ADODB.Recordset
Dim RsPrescription As New ADODB.Recordset
Dim ItemSelected As Integer
Private Sub PopulatePharmacy()
    On Error GoTo ErrorHandler
   KARI = GlbSysDate
    Grid.Clear
    LstPrescription.Clear
    ClearText FrmPharmacy
    Grid.Rows = 1
    Grid.Cols = 9
    Grid.ColAlignment(1) = flexAlignCenterCenter
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER| VISIT NUMBER |  PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |BLOOD PRESSURE |WEIGHT | HEIGHT     |VISIT DATE "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '1' AND TOPHARMACY = '1' AND DRUGS <> '1'", Conn, adOpenDynamic, adLockOptimistic
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND TOPHARMACY = '1' AND DRUGSISSUED = 'False'", Conn, adOpenDynamic, adLockOptimistic

            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !cardnumber & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FIRSTNAME & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER & vbTab & !BP & vbTab & !Weight & vbTab & !Height & vbTab & !VisitDate
                        .MoveNext
                    Wend
                End With
            End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
Private Sub PopulatePartialPayments()
    On Error GoTo ErrorHandler
    Dim lvNames As String
   KARI = GlbSysDate
    
    GridCredit.Clear
    LstPrescription.Clear
    ClearText FrmPharmacy
    GridCredit.Rows = 1
    GridCredit.Cols = 3
    GridCredit.ColAlignment(1) = flexAlignCenterCenter
    GridCredit.ColWidth(1) = 3105
    GridCredit.ColWidth(2) = 3990
    GridCredit.FormatString = "CARD NUMBER  | VISIT NUMBER |  PATIENTS FULL NAME    |    "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DDMMMYYYY") & "' AND COMPLAINS.OBSERVED = '1'AND DIAGNOSED = '1' AND TOPHARMACY = '1' AND DRUGS <> '1'", Conn, adOpenDynamic, adLockOptimistic
        RsGrid.Open "SELECT DISTINCT CardNumber, VisitNumber FROM CREDITORS_NOTES WHERE SUBMITTED = '0'", Conn, adOpenDynamic, adLockOptimistic

            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        lvNames = FindRecord("PATIENT_DETAILS", "FIRSTNAME + ' ' + SECONDNAME AS NAMES", "CARDNUMBER = '" & !cardnumber & "'")
                        GridCredit.AddItem !cardnumber & vbTab & !VISITNUMBER & vbTab & lvNames
                        .MoveNext
                    Wend
                End With
            End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
   ' Resume
End Sub
Public Sub ManageProcessFlow(ActiveForm)
    Dim RsControls As New ADODB.Recordset
    RsControls.Open "SELECT * FROM PROCESSFLOW WHERE SCREENID = '" & ActiveForm & "'", Conn, adOpenStatic, adLockOptimistic
        If RsControls.EOF = False Then
            With RsControls
                If !CONSULTATION = 1 Then OptConsultation.Item(1).Enabled = True
                If !OBSERVATION = 1 Then OptObservation.Item(2).Enabled = True
                If !DOCTORS = 1 Then OptDoctors.Item(3).Enabled = True
                If !CASHIER = 1 Then OptCashier.Item(5).Enabled = True
                'If !LAB = 1 Then OptLab.Item(5).Enabled = True
            End With
        End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub CmdPost_Click()
Dim StrCardNumber As String
KARI = GlbSysDate
StrCardNumber = Grid.TextMatrix(Grid.Row, 0)
    Select Case ItemSelected
   Case 1
        'To Consultation
            DUMMY = SendPatient(EnumConsultation, StrCardNumber, KARI)
            'FrmPatients.Show
            'Unload Me
    Case 2
        'To Observation
            DUMMY = SendPatient(EnumObservation, StrCardNumber, KARI)
            'FrmObservation.Show
            'Unload Me
   Case 3
        'To Doctors
            DUMMY = SendPatient(EnumDoctors, StrCardNumber, KARI)
            'FrmWaitingRoom.Show
            'Unload Me
    Case 5
        'To Cashier
            DUMMY = SendPatient(EnumCashier, StrCardNumber, KARI)
            'FrmCashier.Show
            'Unload Me
    End Select
    PopulatePharmacy
End Sub

Private Sub CmdRefresh_Click()
    If TabPharmacy.Tab = 0 Then
        PopulatePharmacy
    Else
        PopulatePartialPayments
    End If
End Sub

Private Sub CmdSubmit_Click()
On Error GoTo ErrorHandler
    Select Case TabPharmacy.Tab
        Case 0
            If Grid.TextMatrix(Grid.Row, 0) = "" Then Exit Sub
            'UPDATE DRUGS SUBMITTED FLAG SG 22012011
            Conn.Execute "UPDATE COMPLAINS SET DRUGSISSUED = '1',DISMISSED = 'TRUE',PHARMASYST = '" & GlbCurrentUser & "' WHERE CARDNUMBER = '" & Grid.TextMatrix(Grid.Row, 0) & "' AND VISITNUMBER = '" & Grid.TextMatrix(Grid.Row, 1) & "'"
            
            MsgBox "Drugs Submission Updated Succesfully", vbInformation
            PopulatePharmacy
        Case 1
            Conn.Execute "UPDATE CREDITORS_NOTES SET SUBMITTED = '1' WHERE CARDNUMBER = '" & GridCredit.TextMatrix(GridCredit.Row, 0) & "' AND VISITNUMBER = '" & GridCredit.TextMatrix(GridCredit.Row, 1) & "'"
            MsgBox "Drugs Submission Updated Succesfully", vbInformation
            PopulatePartialPayments
    End Select
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub CMDPrintPrescription_Click()
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=Today123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\PRESCRIPTION.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & Grid.TextMatrix(Grid.Row, 0) & "'"
                   .SelectionFormula = " {PRESCRIPTION.VISITNUMBER} = '" & Grid.TextMatrix(Grid.Row, 1) & "'"
                   .Destination = 0
                   .Action = 1
                End With
End Sub

Private Sub Form_Load()
    centerform Me
    PopulatePharmacy
    ManageProcessFlow 4
    GlbCurrentForm = EnumPharmacy
End Sub

Private Sub Grid_DblClick()
    Dim lvTrueCardNumber As String
    If Grid.Row = 0 Then Exit Sub
        StrPharmCardNumber = Grid.TextMatrix(Grid.Row, 0)
        StrPharmVisitDate = Grid.TextMatrix(Grid.Row, 8)
        If rsfill.State = 1 Then rsfill.Close
                rsfill.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CardNumber = COMPLAINS.CardNumber AND COMPLAINS.CardNumber = '" & StrPharmCardNumber & "' AND COMPLAINS.VisitDate = '" & Format(StrPharmVisitDate, "DDMMMYYYY") & "' AND TOPHARMACY = '1'", Conn, adOpenStatic, adLockOptimistic
                    With rsfill
                        TxtCardNumber = Grid.TextMatrix(Grid.Row, 0)
                        TxtPatientName = !SURNAME & " " & !FIRSTNAME & " " & !SECONDNAME
                        TxtDocName = !DOCTOR
                        If TxtDiagnosis <> Null Then TxtDiagnosis = !DIAGNOSIS
                        If TxtPrescription <> Null Then TxtDiagnosis = !PRESCRIPTION
                        
                        'PRESCRIPTIONS NI MORE THAN ONE SO WACHA NI LOOP
                        LstPrescription.Clear
                            If CStr(TxtCardNumber.Text) = "0" Then ' MEANS THIS RECORD IS FROM WARDS
                                lvTrueCardNumber = FindRecord("COMPLAINS", "ADMISSIONNUMBER", "VISITNUMBER = '" & !VISITNUMBER & "'")
                                RsPrescription.Open "SELECT * FROM PRESCRIPTION WHERE CARDNUMBER = '" & lvTrueCardNumber & "' AND VISITNUMBER = '" & !VISITNUMBER & "' AND CODE <> '001'", Conn, adOpenStatic, adLockOptimistic
                            Else
                                RsPrescription.Open "SELECT * FROM PRESCRIPTION WHERE CARDNUMBER = '" & TxtCardNumber & "' AND VISITNUMBER = '" & !VISITNUMBER & "' AND CODE <> '001'", Conn, adOpenStatic, adLockOptimistic
                            End If
                                With RsPrescription
                                    While .BOF = False And .EOF = False
                                        LstPrescription.AddItem !CODE & " - " & !Description & ",       " & "Quantity : " & !Quantity
                                        .MoveNext
                                    Wend
                                End With
                            RsPrescription.Close
                        'END
                rsfill.Close
            End With
End Sub

Private Sub OptAdmission_Click(index As Integer)
    ItemSelected = OptAdmission.Item(index).index
End Sub

Private Sub GridCredit_Click()
    LstPrescription.Clear
    rsfill.Open "SELECT * FROM CREDITORS_NOTES WHERE CARDNUMBER = '" & GridCredit.TextMatrix(GridCredit.Row, 0) & "' AND VISITNUMBER = '" & GridCredit.TextMatrix(GridCredit.Row, 1) & "'", Conn, adOpenStatic, adLockOptimistic
        If rsfill.EOF = False Then
            With rsfill
                While .EOF = False
                    LstPrescription.AddItem !NOTES
                    .MoveNext
                Wend
            End With
        End If
    rsfill.Close
End Sub

Private Sub OptCashier_Click(index As Integer)
    ItemSelected = OptCashier.Item(index).index
End Sub

Private Sub OptConsultation_Click(index As Integer)
    ItemSelected = OptConsultation.Item(index).index
End Sub

Private Sub OptDoctors_Click(index As Integer)
    ItemSelected = OptDoctors.Item(index).index
End Sub

Private Sub OptObservation_Click(index As Integer)
    ItemSelected = OptObservation.Item(index).index
End Sub

Private Sub OptPharmacy_Click(index As Integer)
    ItemSelected = OptPharmacy.Item(index).index
End Sub

Private Sub TabPharmacy_Click(PreviousTab As Integer)
    PopulatePartialPayments
End Sub

