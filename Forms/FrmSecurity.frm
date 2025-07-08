VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmSecurity 
   Caption         =   "Security Module"
   ClientHeight    =   5400
   ClientLeft      =   3195
   ClientTop       =   2970
   ClientWidth     =   9825
   Icon            =   "FrmSecurity.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   9825
   Begin VB.Frame Frame2 
      Caption         =   "Navigation"
      Height          =   5295
      Left            =   8280
      TabIndex        =   13
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   4680
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "User Creation and Rights allocation"
      TabPicture(0)   =   "FrmSecurity.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "List Of Maintained Users"
      TabPicture(1)   =   "FrmSecurity.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "List Of Users"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   7815
         Begin VSFlex6DAOCtl.vsFlexGrid G 
            Height          =   4335
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   7646
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
         Caption         =   "User Details"
         Height          =   4695
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   7815
         Begin VB.CheckBox ChkLockUser 
            Caption         =   "Lock User From Login Access"
            Height          =   255
            Left            =   5040
            TabIndex        =   16
            Top             =   2099
            Width           =   2535
         End
         Begin VB.Frame Frame4 
            Caption         =   "Assign Rights"
            Height          =   2415
            Left            =   120
            TabIndex        =   17
            Top             =   2160
            Width           =   7575
            Begin VSFlex6DAOCtl.vsFlexGrid ProfileGrid 
               Height          =   2055
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   7335
               _ExtentX        =   12938
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
               Editable        =   -1  'True
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
         Begin VB.TextBox TxtPassword 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   2280
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   1560
            Width           =   4335
         End
         Begin VB.TextBox TxtFullName 
            Height          =   375
            Left            =   2280
            TabIndex        =   1
            Top             =   960
            Width           =   4335
         End
         Begin VB.TextBox TxtUsername 
            Height          =   375
            Left            =   2280
            TabIndex        =   0
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label3 
            Caption         =   "Password"
            Height          =   375
            Left            =   1080
            TabIndex        =   12
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Full Name"
            Height          =   255
            Left            =   1080
            TabIndex        =   11
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "System User Name"
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   480
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "FrmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Dim Conn As New ADODB.Connection
Dim RsRecords As New ADODB.Recordset
Dim RsSalaries As New ADODB.Recordset
Dim RsSearch As New ADODB.Recordset
Dim RSProfiles As New ADODB.Recordset
Dim PassEncrypt As New Encryption.EncryptDecrypt

Public Sub AddMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = False
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmSecurity
    SSTab1.Tab = 0
End Sub
Public Sub EditMode()
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdSave.Enabled = True
    CmdDelete.Enabled = True
    CMdClose.Caption = "Cancel"
    ReverseGreyOut FrmSecurity
End Sub
Public Sub ResetMode()
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdSave.Enabled = False
    CmdDelete.Enabled = False
    CMdClose.Caption = "Close"
    GreyOut FrmSecurity
End Sub
Public Sub POPULATEGRID()
On Error GoTo ErrorHandler
Dim KARI As String
    G.Clear: G.Rows = 1
    G.Cols = 2
    G.FormatString = "USER NAME|FULL NAME|LOCK STATUS|"
    G.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM USERS", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        G.AddItem !UserName & vbTab & !fullname & vbTab & KARI
                    End With
                RsRecords.MoveNext
                Wend
            End If
        RsRecords.Close
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
        
End Sub
Public Sub POPULATEPROFILES()
On Error GoTo ErrorHandler
    ProfileGrid.Clear: ProfileGrid.Rows = 1
    ProfileGrid.Cols = 2
    ProfileGrid.FormatString = "USER NAME|MENU TO ACCESS|USER STATUS|"
    ProfileGrid.ColDataType(2) = flexDTBoolean
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        If TxtUserName <> "" Then
            RsRecords.Open "SELECT * FROM PROFILES WHERE USERNAME = '" & TxtUserName & "' order by rights asc", Conn, adOpenStatic, adLockOptimistic
        Else
            RsRecords.Open "SELECT * FROM PROFILES order by rights asc", Conn, adOpenStatic, adLockOptimistic
        End If
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                While RsRecords.EOF = False
                    With RsRecords
                        ProfileGrid.AddItem !UserName & vbTab & !RIGHTS & vbTab & !ACCESS
                    End With
                RsRecords.MoveNext
                Wend
            End If
        RsRecords.Close
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
On Error GoTo ErrorHandler
Dim Resp
    Resp = MsgBox("Are you sure you want to Delete this record?", vbQuestion + vbYesNo)
    If Resp = vbYes Then
        Conn.Execute "DELETE FROM USERS WHERE USERNAME = '" & TxtUserName & "'"
        POPULATEGRID
        
        'DELETE FROM PROFILES AS WELL
        Conn.Execute "DELETE FROM PROFILES WHERE USERNAME = '" & TxtUserName & "'"
        MsgBox "Record Deleted Succesfully", vbInformation
        POPULATEPROFILES
        'AuditTrail StrCurrentUser, Date, Time, "Deleted user '" & TxtUserName.Text & "'"
        AuditTrail GlbCurrentUser, EnumSecurity, GlbSysDate, Time, "Deleted User - " + " " & TxtUserName & ""
    Else
        MsgBox "Deletion Proces Aborted", vbInformation
    End If
    ResetMode
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
    
End Sub

Private Sub CmdEdit_Click()
    EditMode
    'Clear password so as not to edit it when assigning rights.
    TxtPassword.Text = ""
End Sub

Private Sub CmdSave_Click()
On Error GoTo ErrorHandler
'Dim KARI, KARI2, KARI3, KARI4 As String
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM USERS WHERE USERNAME = '" & TxtUserName & "'", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
            'ENCRYPT PASSWORD FIRST
                PassEncrypt.Text = TxtPassword.Text
                Dim strLegaltext As String
                PassEncrypt.Keystring = "SOLOMON"
                PassEncrypt.DoXor
                PassEncrypt.Stretch

            '**********************
                        RsRecords!UserName = TxtUserName
                        If TxtPassword <> "" Then 'if txtpassword is blank, then password was not being reset. it is rights allocation.
                            RsRecords!Password = PassEncrypt.Text
                        End If
                            RsRecords!fullname = TxtFullName
                            RsRecords!Status = ChkLockUser.Value
                        RsRecords.Update
                        'Routine to asign rights. SG 25022011
                        If RSProfiles.State = 1 Then Set RSProfiles = Nothing
                        RSProfiles.Open "SELECT DISTINCT RIGHTS FROM PROFILES ORDER BY RIGHTS asc", Conn, adOpenStatic, adLockOptimistic
                            With RSProfiles
                                For i = 1 To ProfileGrid.Rows - 1
                                    If .EOF = True Then Exit For
                                    If ProfileGrid.TextMatrix(i, 2) = True Then
                                       Conn.Execute "UPDATE PROFILES SET ACCESS = '1' WHERE USERNAME = '" & TxtUserName & "' AND RIGHTS = '" & ProfileGrid.TextMatrix(i, 1) & "'"
                                       Else
                                       Conn.Execute "UPDATE PROFILES SET ACCESS = '0' WHERE USERNAME = '" & TxtUserName & "' AND RIGHTS = '" & ProfileGrid.TextMatrix(i, 1) & "'"
                                    End If
                                    .MoveNext
                                Next
                            End With
                        
                    MsgBox "Record Edited Successfully", vbInformation
                    'AuditTrail StrCurrentUser, Date, Time, "Assigned Rights to User '" & TxtUsername & "'"
                    POPULATEGRID
            Else
            If TxtUserName = "" Then MsgBox "Username Cannot be blank. Please enter before saving", vbInformation: Exit Sub
            If TxtUserName = "" Then MsgBox "Full Name Cannot be blank. Please enter before saving", vbInformation: Exit Sub
            If TxtUserName = "" Then MsgBox "Password Cannot be blank. Please enter before saving", vbInformation: Exit Sub
            
                With RsRecords
                    .AddNew
                        PassEncrypt.Text = TxtPassword.Text
                        PassEncrypt.Keystring = "SOLOMON"
                        PassEncrypt.DoXor
                        PassEncrypt.Stretch
                    
                        RsRecords!UserName = TxtUserName
                        RsRecords!Password = PassEncrypt.Text
                        RsRecords!fullname = TxtFullName
                        RsRecords!CHANGEPASSWORD = "1"
                        RsRecords!Status = "0"
                    .Update
                    'CREATE PROFILES FOR ADDED USER
                    Set RsRecords = Nothing
                    RsRecords.Open "SELECT DISTINCT ProfileDescription FROM PROFILEITEMS", Conn, adOpenStatic, adLockOptimistic
                        If RsRecords.BOF = False And RsRecords.EOF = False Then
                            Set RsSearch = Nothing: RsRecords.MoveLast: RsRecords.MoveFirst
                                RsSearch.Open "SELECT * FROM PROFILES", Conn, adOpenStatic, adLockOptimistic
                                    While RsRecords.EOF = False
                                        RsSearch.AddNew
                                            RsSearch!UserName = TxtUserName
                                            RsSearch!RIGHTS = RsRecords!ProfileDescription
                                           RsSearch!ACCESS = "0"
                                        RsSearch.Update
                                        RsRecords.MoveNext
                                    Wend
                                Set RsSearch = Nothing
                        End If
                    MsgBox "Record Added Successfully", vbInformation
                    POPULATEGRID
                    'AuditTrail StrCurrentUser, Date, Time, "ADDED USER '" & TxtUserName & "'"
                    AuditTrail GlbCurrentUser, EnumSecurity, GlbSysDate, Time, "Created User - " + " " & TxtUserName & ""
                End With
            End If
    ResetMode
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    If Conn.State = 0 Then
        'Set Conn = Nothing
        'Conn.Open
        ResetMode
    End If
        POPULATEGRID
        POPULATEPROFILES
        centerform Me
        GlbCurrentForm = EnumSecurity
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
        
End Sub

Private Sub G_Click()
On Error GoTo ErrorHandler
    If G.Row < 1 Then Exit Sub
    If RsRecords.State = adStateOpen Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM USERS WHERE USERNAME = '" & G.TextMatrix(G.Row, 0) & "'", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
                TxtUserName = RsRecords!UserName
                TxtFullName = RsRecords!fullname
                TxtPassword = RsRecords!Password
                SSTab1.Tab = 0
                POPULATEPROFILES
            End If
        Set RsRecords = Nothing
    Exit Sub
Exit Sub
ErrorHandler:
MsgBox Err.Number & Err.Description
End Sub


Private Sub TxtFullName_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Or KeyAscii = 34 Or KeyAscii = 124 Then KeyAscii = 0
   If KeyAscii <> vbKeyBack Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub TxtUsername_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Or KeyAscii = 34 Or KeyAscii = 124 Then KeyAscii = 0
   If KeyAscii <> vbKeyBack Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If

End Sub
