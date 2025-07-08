VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmChequesMaintenance 
   Appearance      =   0  'Flat
   Caption         =   "Cheque Maintanance"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   Icon            =   "FrmChequesMaintenance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Navigate"
      Height          =   3735
      Left            =   7800
      TabIndex        =   14
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   1600
         Width           =   1335
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   920
         Width           =   1335
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "New"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cheque Details"
      TabPicture(0)   =   "FrmChequesMaintenance.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "List Of Cheque"
      TabPicture(1)   =   "FrmChequesMaintenance.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   24
         Top             =   360
         Width           =   7335
         Begin VSFlex6DAOCtl.vsFlexGrid G 
            Height          =   2895
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5106
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
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7335
         Begin MSComCtl2.DTPicker DTDepositDate 
            Height          =   375
            Left            =   5160
            TabIndex        =   23
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55181313
            CurrentDate     =   39608
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   315
            Left            =   6720
            TabIndex        =   21
            Top             =   2760
            Width           =   375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   315
            Left            =   6720
            TabIndex        =   20
            Top             =   2280
            Width           =   375
         End
         Begin VB.TextBox TxtCardNumber 
            BackColor       =   &H8000000F&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   360
            Width           =   5055
         End
         Begin VB.TextBox TxtChequeNumber 
            Height          =   375
            Left            =   2040
            TabIndex        =   6
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox TxtAccountNumber 
            Height          =   375
            Left            =   2040
            TabIndex        =   5
            Top             =   1800
            Width           =   5055
         End
         Begin VB.TextBox TxtVisitNumber 
            BackColor       =   &H8000000F&
            Height          =   375
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   840
            Width           =   5055
         End
         Begin VB.ComboBox CmbBank 
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Top             =   2280
            Width           =   4575
         End
         Begin VB.ComboBox CmbBranch 
            Height          =   315
            Left            =   2040
            TabIndex        =   2
            Top             =   2760
            Width           =   4575
         End
         Begin VB.Label Label7 
            Caption         =   "Deposit Date"
            Height          =   255
            Left            =   4080
            TabIndex        =   22
            Top             =   1400
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Card Number"
            Height          =   255
            Left            =   720
            TabIndex        =   13
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Cheque Number"
            Height          =   255
            Left            =   600
            TabIndex        =   12
            Top             =   1392
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Account Number"
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   1848
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Visit Number"
            Height          =   255
            Left            =   840
            TabIndex        =   10
            Top             =   936
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Originatin Bank"
            Height          =   255
            Left            =   720
            TabIndex        =   9
            Top             =   2304
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Originating Branch"
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   2760
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "FrmChequesMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Private Sub AddMode()
    CmdNew.Enabled = False
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = True
    CmdExit.Caption = "Cancel"
End Sub
Private Sub EditMode()
    CmdNew.Enabled = False
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = True
    CmdExit.Caption = "Cancel"
End Sub
Private Sub ResetMode()
    CmdNew.Enabled = True
    CmdEdit.Enabled = True
    CmdDelete.Enabled = True
    CmdSave.Enabled = False
    CmdExit.Caption = "Exit"
End Sub

Private Sub CmbBank_CLICK()
    CmbBranch.Clear
    'POPULATE DROP DOWN BOXES
    Dim RsCombo As New ADODB.Recordset
        RsCombo.Open "SELECT BRCODE,BRANCHNAME FROM BRANCHES WHERE BKCODE = '" & Left(CmbBank, 2) & "'", Conn, adOpenStatic, adLockOptimistic
            With RsCombo
                While .EOF = False
                    CmbBranch.AddItem !BRCODE + " - " + !BRANCHNAME
                    .MoveNext
                Wend
            End With
        Set RsCombo = Nothing
End Sub
Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 5
    G.FormatString = "CARD NUMBER| ACCOUNT NUMBEER | CHEQUE NUMBER | VISIT NUMBER | DEPOSIT DATE"
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  CHEQUE_PAYMENTS", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !CHQNO & vbTab & !ACCOUNTNUMBER & vbTab & !cardnumber & vbTab & !VISITNUMBER & vbTab & !DEPOSITDATE
                    End With
                RsRecords.MoveNext
                Wend
            End If
        RsRecords.Close
    G.Editable = True
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
End Sub

Private Sub CmdDelete_Click()
    ResetMode
End Sub

Private Sub CmdEdit_Click()
    EditMode
End Sub

Private Sub CmdExit_Click()
    If CmdExit.Caption = "Cancel" Then
        ResetMode
    Else
        Unload Me
    End If
End Sub

Private Sub CmdNew_Click()
    AddMode
End Sub

Private Sub CmdSave_Click()
    If TxtCardNumber = "" Then MsgBox "CARD DETAILS HAVE NOT BEEN SELECTED. SELECT ON PREVIOUS SCREEN": Exit Sub
    
    Conn.Execute "INSERT INTO CHEQUE_PAYMENTS(CHQNO,ACCOUNTNUMBER,CARDNUMBER,VISITNUMBER,DEPOSITDATE)" & _
                 "VALUES('" & TxtChequeNumber & "','" & TxtAccountNumber & "','" & TxtCardNumber & "','" & TxtVisitNumber & "','" & Format(DTDepositDate, "DD MMM YYYY") & "')"
    ResetMode
    Unload Me
End Sub

Private Sub Form_Load()
    centerform Me
    TxtCardNumber = StrDocCardNo
    TxtVisitNumber = StrDocVisitNumber
    'POPULATE DROP DOWN BOXES
    Dim RsCombo As New ADODB.Recordset
        RsCombo.Open "SELECT CODE,FULLNAME FROM BANKS", Conn, adOpenStatic, adLockOptimistic
            With RsCombo
                While .EOF = False
                    CmbBank.AddItem !CODE + " - " + !fullname
                    .MoveNext
                Wend
            End With
        Set RsCombo = Nothing
    SSTab1.Tab = 0
    AddMode
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        POPULATEGRID
    End If
End Sub

Private Sub SSTab1_DblClick()
    'POPULATEGRID
End Sub
