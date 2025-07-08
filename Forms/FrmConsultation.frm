VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmConsultation 
   Caption         =   "Consultation Fee Maintenance"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   Icon            =   "FrmConsultation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Navigate"
      Height          =   5535
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton CmdAdd 
         Caption         =   "New"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton CMdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "List Of Consultation Fees per Billing Company"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   5535
      Begin VSFlex6DAOCtl.vsFlexGrid G 
         Height          =   3015
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5318
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
   Begin VB.Frame Frame2 
      Caption         =   "Consultation Fees"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox TxtConsultationFees 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox TxtDiscount 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox CboServiceProvider 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Discount"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Consultation Fee"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Service Provider"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmConsultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim BlnEditing As Boolean

Public Sub AddMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmConsultation
    ClearText FrmConsultation
End Sub
Public Sub EditMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmConsultation
End Sub
Public Sub ResetMode()
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    CMdClose.Caption = "Close"
    GreyOut FrmConsultation
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
    G.Clear: G.Rows = 1: G.Cols = 2
    G.FormatString = "SERVICE PROVIDEER| CONSULTATION FEE|DISCOUNTED AMOUNT "
    'G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM  SERVICE_PROVIDER", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !COMPANYCODE + " - " + !SERVICEPROVIDER & vbTab & !ConsultationFee & vbTab & !DISCOUNT
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

Private Sub cmdadd_Click()
    AddMode
End Sub

Private Sub CMDCLOSE_Click()
    If CMdClose.Caption = "Cancel" Then
        ResetMode
    Else
        Unload Me
    End If
End Sub

Private Sub CmdDelete_Click()
Dim Resp
    Resp = MsgBox("Are you sure you wish to delete this record?", vbInformation + vbYesNo)
        If Resp = vbYes Then
            Conn.Execute "DELETE FROM [CONSULTATION FEE] WHERE COMPANYCODE = '" & Mid(CboServiceProvider, 1, 3) & "'"
            MsgBox "Record Deleted Succesfully", vbInformation
            POPULATEGRID
            BlnEditing = False
        Else
            MsgBox "Deletion aborted", vbInformation
        End If
    ResetMode
End Sub

Private Sub CmdEdit_Click()
    EditMode
    BlnEditing = True
End Sub

Private Sub CmdSave_Click()
    Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM SERVICE_PROVIDER WHERE SERVICEPROVIDER = '" & GetID_NameFromCombo(CboServiceProvider, 2) & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            If BlnEditing = True Then
                If .BOF = False And .EOF = False Then
                    !COMPANYCODE = Mid(CboServiceProvider, 1, 3)
                    !SERVICEPROVIDER = GetID_NameFromCombo(CboServiceProvider.Text, 2)
                    !ConsultationFee = TxtConsultationFees
                    !DISCOUNT = TxtDiscount
                    .Update
                    MsgBox "Consultation Fees Edited Successfully", vbInformation
                    BlnEditing = False
                End If
            Else
                If .BOF = False And .EOF = False Then
                    MsgBox "Billing Company with code " & TxtCompanyCode & " is already Maintained", vbExclamation, "Duplication"
                Else
                    .AddNew
                    !COMPANYCODE = Mid(CboServiceProvider, 1, 3)
                    !SERVICEPROVIDER = GetID_NameFromCombo(CboServiceProvider.Text, 2)
                    !ConsultationFee = TxtConsultationFees
                    !DISCOUNT = TxtDiscount
                    .Update
                    MsgBox "Consultation Fees Successfully Maintained", vbInformation
                End If
            End If
        End With
    POPULATEGRID
    ResetMode
    BlnEditing = False
End Sub
Private Sub G_Click()
    On Error GoTo ErrorHandler
        CboServiceProvider.Text = G.TextMatrix(G.Row, 0)
        TxtConsultationFees = G.TextMatrix(G.Row, 1)
        TxtDiscount = G.TextMatrix(G.Row, 2)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()

    'POPULATE GRID
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM SERVICE_PROVIDER", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboServiceProvider.AddItem RsRecords!COMPANYCODE & " - " & RsRecords!SERVICEPROVIDER
            RsRecords.MoveNext
        Wend
    RsRecords.Close
    
    POPULATEGRID
    ResetMode
    centerform Me
End Sub


Private Sub G_DblClick()
    FrmCashier.TxtAmount = G.TextMatrix(G.Row, 1) - G.TextMatrix(G.Row, 2)
    Unload Me
End Sub
