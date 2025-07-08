VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form FrmBanksDefin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Codes Definition"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ForeColor       =   &H00C0C0C0&
   Icon            =   "FrmBanksDefin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   1200
      Width           =   1330
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   720
      Width           =   1330
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   240
      Width           =   1330
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox TxtBankAbbrev 
         BackColor       =   &H80000018&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtBankCode 
         BackColor       =   &H80000018&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtBankName 
         BackColor       =   &H80000018&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Abbreviation"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Code"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid VsBanksDetails 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3625
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
      GridColor       =   32768
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   4
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   3
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
      WordWrap        =   0   'False
      TextStyle       =   3
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin VB.CheckBox chkMybank 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "My Bank"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "FrmBanksDefin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim RsRecordTypes As New ADODB.Recordset
Dim RecordTypes() As Variant

Private Sub ChkAppoint_Click(Value As Integer)
    cboAppointBank.Enabled = Value
End Sub

Private Sub cmdadd_Click()
On Error GoTo ErrorHandler
    Select Case UCase(CmdAdd.Caption)
        Case "&ADD"
            'Enable all Controls
            CmdAdd.Caption = "&Save"
            cmdClose.Caption = "&Cancel"
            cmdDelete.Enabled = False
            For x = 0 To FrmBanksDefin.Controls.Count - 1
                If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
                    FrmBanksDefin.Controls(x).Enabled = True
                ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
                    FrmBanksDefin.Controls(x).Enabled = True
                End If
            Next x
            'Clear all the Controls
            For x = 0 To FrmBanksDefin.Controls.Count - 1
                If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
                    FrmBanksDefin.Controls(x).ListIndex = -1
                ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
                    FrmBanksDefin.Controls(x).Text = ""
                End If
            Next x
            TxtBankCode.SetFocus
        Case "&SAVE"
            'Adding to the Database and Validation
            For x = 0 To FrmBanksDefin.Controls.Count - 1
                If TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
                    If FrmBanksDefin.Controls(x).Text = "" Then MsgBox ("Before Saving The Record you Have to Input all the Necessary Fields"), vbExclamation: Exit Sub
                End If
            Next x
            If RsRecordTypes.State = 1 Then RsRecordTypes.Close
            RsRecordTypes.Open ("SELECT CODE,ABBREV,FULLNAME FROM BANKS WHERE CODE='" & TxtBankCode.Text & "'"), DataConn, adOpenKeyset, adLockOptimistic
            If RsRecordTypes.BOF = True And RsRecordTypes.EOF = True Then
                RsRecordTypes.AddNew
            Else
                If MsgBox("This Record already Exists. Do you wish to update?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
            End If
            With RsRecordTypes
                    !Code = TxtBankCode.Text
                    !Abbrev = TxtBankAbbrev.Text
                    !FullName = TxtBankName.Text
                    '!Default = chkMybank.Value
                .Update
            End With
            MsgBox ("Record Successfully Saved!"), vbInformation
            FillRecordTypes
            'Clear all the Controls
            For x = 0 To FrmBanksDefin.Controls.Count - 1
                If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
                    FrmBanksDefin.Controls(x).ListIndex = -1
                ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
                    FrmBanksDefin.Controls(x).Text = ""
                End If
            Next x
            'DISABLE all Controls
            CmdAdd.Caption = "&Add"
            cmdClose.Caption = "&Close"
            cmdDelete.Enabled = True
            For x = 0 To FrmBanksDefin.Controls.Count - 1
                If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
                    FrmBanksDefin.Controls(x).Enabled = False
                ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
                    FrmBanksDefin.Controls(x).Enabled = False
                End If
            Next x
    End Select
Exit Sub
ErrorHandler:
    SystemErrorHandler err.Number, err.Description, Me
    Select Case ErrorMessageResponse
        Case 3 'Abort
            Exit Sub
        Case 4 'Retry
            Resume
        Case 5 'Ignore
            Resume Next
        Case Else
    End Select
End Sub
Private Sub FillRecordTypes()
On Error GoTo ErrorHandler
    VsBanksDetails.Clear
    VsBanksDetails.Rows = 1
    VsBanksDetails.Cols = 3
    VsBanksDetails.FormatString = "Bank Code|Bank Abbreviation|Bank FullName"
    VsBanksDetails.ColWidth(0) = 1000
    VsBanksDetails.ColWidth(1) = 1000
    VsBanksDetails.ColWidth(2) = 7000
    VsBanksDetails.ColAlignment(0) = flexAlignCenterCenter
    If RsRecordTypes.State = 1 Then RsRecordTypes.Close
    RsRecordTypes.Open ("SELECT CODE,ABBREV,FULLNAME FROM BANKS ORDER BY CODE ASC"), DataConn, adOpenKeyset, adLockOptimistic
    If RsRecordTypes.BOF = False And RsRecordTypes.EOF = False Then
        RsRecordTypes.MoveLast: RsRecordTypes.MoveFirst
        ReDim RecordTypes(0 To RsRecordTypes.RecordCount - 1, 0 To RsRecordTypes.Fields.Count - 1)
        For x = 0 To RsRecordTypes.RecordCount - 1
            For Y = 0 To RsRecordTypes.Fields.Count - 1
                RecordTypes(x, Y) = RsRecordTypes.Fields(Y).Value
            Next Y
            RsRecordTypes.MoveNext
        Next x
        'loading into the Grid
        VsBanksDetails.Rows = UBound(RecordTypes, 1) + 2
        For x = 0 To UBound(RecordTypes, 1)
            For Y = 0 To UBound(RecordTypes, 2)
                VsBanksDetails.TextMatrix(x + 1, Y) = IIf(IsNull(RecordTypes(x, Y)), "", RecordTypes(x, Y))
            Next Y
        Next x
    End If
Exit Sub
ErrorHandler:
    SystemErrorHandler err.Number, err.Description, Me
    Select Case ErrorMessageResponse
        Case 3 'Abort
            Exit Sub
        Case 4 'Retry
            Resume
        Case 5 'Ignore
            Resume Next
        Case Else
    End Select
End Sub

Private Sub CMDCLOSE_Click()
On Error GoTo ErrorHandler
Dim rsMybank As ADODB.Recordset

    Select Case UCase(cmdClose.Caption)
        Case "&CLOSE"
            Unload Me
        Case "&CANCEL", "CANCEL"
            'Enable all Controls
            CmdAdd.Caption = "Add"
            cmdClose.Caption = "&Close"
            CmdAdd.Enabled = True
            For x = 0 To FrmBanksDefin.Controls.Count - 1
                If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
                    FrmBanksDefin.Controls(x).Enabled = False
                ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
                    FrmBanksDefin.Controls(x).Enabled = False
                End If
            Next x
            'cLEAR
            'Clear all the Controls
            For x = 0 To FrmBanksDefin.Controls.Count - 1
                If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
                    FrmBanksDefin.Controls(x).ListIndex = -1
                ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
                    FrmBanksDefin.Controls(x).Text = ""
                End If
            Next x
            cmdDelete.Enabled = False
    End Select
    Exit Sub
ErrorHandler:
    SystemErrorHandler err.Number, err.Description, Me
    Select Case ErrorMessageResponse
        Case 3 'Abort
            Exit Sub
        Case 4 'Retry
            Resume
        Case 5 'Ignore
            Resume Next
        Case Else
    End Select
End Sub

Private Sub CmdDelete_Click()
On Error GoTo ErrorHandler
    If MsgBox("Are you Sure you Want To Delete The Record with Voucher Code " & vbCr & " Number " & TxtBankCode.Text & " From the Database", vbYesNo + vbExclamation) = vbYes Then
        DataConn.Execute ("DELETE FROM BANKS WHERE CODE='" & TxtBankCode.Text & "'")
        MsgBox ("Record Successfully Deleted"), vbInformation
        cmdDelete.Enabled = False
    End If
    'Clear all the Controls
    For x = 0 To FrmBanksDefin.Controls.Count - 1
        If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
            FrmBanksDefin.Controls(x).ListIndex = -1
        ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
            FrmBanksDefin.Controls(x).Text = ""
        End If
    Next x
    'DISABLE all Controls
    CmdAdd.Caption = "&Add"
    cmdClose.Caption = "&Close"
    cmdDelete.Enabled = True
    For x = 0 To FrmBanksDefin.Controls.Count - 1
        If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
            FrmBanksDefin.Controls(x).Enabled = False
        ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
            FrmBanksDefin.Controls(x).Enabled = False
        End If
    Next x
    cmdDelete.Enabled = False
    CmdAdd.Caption = "&Add"
    cmdClose.Caption = "&Close"
    FillRecordTypes
Exit Sub
ErrorHandler:
    SystemErrorHandler err.Number, err.Description, Me
    Select Case ErrorMessageResponse
        Case 3 'Abort
            Exit Sub
        Case 4 'Retry
            Resume
        Case 5 'Ignore
            Resume Next
        Case Else
    End Select
End Sub

Private Sub Form_KeyPress(keyascii As Integer)
     On Error GoTo ErrorHandler
     keyascii = UPPER(keyascii)
    If keyascii = 13 Then SendKeys (Chr(9))
    Exit Sub
ErrorHandler:
    SystemErrorHandler err.Number, err.Description, Me
    Select Case ErrorMessageResponse
        Case 3 'Abort
            Exit Sub
        Case 4 'Retry
            Resume
        Case 5 'Ignore
            Resume Next
        Case Else
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    centerform Me
    FillRecordTypes
    For x = 0 To FrmBanksDefin.Controls.Count - 1
        If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
            FrmBanksDefin.Controls(x).Enabled = False
        ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
            FrmBanksDefin.Controls(x).Enabled = False
        End If
    Next x
    cmdDelete.Enabled = False
    
    ACTIONDONE = "Loaded " & Me.Caption
    AuditTrail g_sCurrentUser, Format(Now, "HH:MM:SS"), ACTIONDONE, Format(Date, "dd/mm/yyyy"), "MN05", audOperations
    Exit Sub
ErrorHandler:
    SystemErrorHandler err.Number, err.Description, Me
    Select Case ErrorMessageResponse
        Case 3 'Abort
            Exit Sub
        Case 4 'Retry
            Resume
        Case 5 'Ignore
            Resume Next
        Case Else
    End Select
End Sub

Private Sub TxtBankCode_KeyPress(keyascii As Integer)
    If ((keyascii) < 48 Or (keyascii) > 57) And keyascii <> 9 And keyascii <> 8 And keyascii <> 13 Then keyascii = 0: Exit Sub
End Sub

Private Sub VsBanksDetails_Click()
     On Error GoTo ErrorHandler
     If VsBanksDetails.Row = 0 Then Exit Sub
    'Assuming the User Wants to Edit the Record
    CmdAdd.Caption = "Save"
    cmdClose.Caption = "Cancel"
    cmdDelete.Enabled = True
    For x = 0 To FrmBanksDefin.Controls.Count - 1
        If TypeOf FrmBanksDefin.Controls(x) Is ComboBox Then
            FrmBanksDefin.Controls(x).Enabled = True
        ElseIf TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
            FrmBanksDefin.Controls(x).Enabled = True
        End If
    Next x
    'Clear all the Controls
    For x = 0 To FrmBanksDefin.Controls.Count - 1
        If TypeOf FrmBanksDefin.Controls(x) Is TextBox Then
            FrmBanksDefin.Controls(x).Text = ""
        End If
    Next x
    'Bank Code
    VsBanksDetails.Col = 0
    TxtBankCode.Text = VsBanksDetails.Text
    'Bank Abbreviation
    VsBanksDetails.Col = 1
    TxtBankAbbrev.Text = VsBanksDetails.Text
    'Bank Name
    VsBanksDetails.Col = 2
    TxtBankName.Text = VsBanksDetails.Text
   
    VsBanksDetails.Col = 0
    Exit Sub
ErrorHandler:
    SystemErrorHandler err.Number, err.Description, Me
    Select Case ErrorMessageResponse
        Case 3 'Abort
            Exit Sub
        Case 4 'Retry
            Resume
        Case 5 'Ignore
            Resume Next
        Case Else
    End Select
End Sub

Private Sub VsBanksDetails_KeyDown(KEYCODE As Integer, Shift As Integer)
    VsBanksDetails_Click
End Sub


