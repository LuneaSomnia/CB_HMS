VERSION 5.00
Begin VB.Form FrmConfirmPassword 
   Caption         =   "Change Password"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "FrmConfirmPassword.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   5520
      TabIndex        =   9
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton CmdClose 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reset Password"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.TextBox TxtConfirmPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox TxtNewPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox TxtOldPassword 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox TxtUserName 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "User Name"
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
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirm New Password"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "New Password"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Old Password"
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
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmConfirmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset
Dim PassEncrypt As New Encryption.EncryptDecrypt

Private Sub Command1_Click()
On Error GoTo ErrorHandler
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM USERS WHERE USERNAME = '" & TxtUserName & "'", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
            
                'Compare Passwords
                If TxtNewPassword <> TxtConfirmPassword Then MsgBox "New and Confirmed Passwords do not match. Please Re-enter", vbExclamation: Exit Sub
                'Refuse New Password to be same as Old Password.
                If TxtNewPassword = TxtOldPassword Then MsgBox "New Password cannot be the same as the Old Password. Please Re-enter", vbExclamation: TxtNewPassword = "": TxtConfirmPassword = "": Exit Sub
                PassEncrypt.Text = TxtNewPassword.Text
                PassEncrypt.Keystring = "SOLOMON"
                PassEncrypt.DoXor
                PassEncrypt.Stretch
                    
                RsRecords!UserName = TxtUserName
                RsRecords!Password = PassEncrypt.Text
                RsRecords!CHANGEPASSWORD = "0"
                RsRecords.Update
                MsgBox "Password Changed Succesfully", vbInformation
            End If
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub CMDCLOSE_Click()
    If GlbPasswordResetLogin = True Then
        MsgBox "Password was not changed. Login will not be allowed", vbInformation
        End
    Else
        Unload Me
    End If
End Sub

Private Sub CmdSave_Click()
On Error GoTo ErrorHandler
    If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM USERS WHERE USERNAME = '" & TxtUserName & "'", Conn, adOpenStatic, adLockOptimistic
            If RsRecords.BOF = False And RsRecords.EOF = False Then
            
                'Compare Passwords
                If TxtNewPassword <> TxtConfirmPassword Then MsgBox "New and Confirmed Passwords do not match. Please Re-enter", vbExclamation: Exit Sub
                
                'Decrypt Password
                                    
                    PassEncrypt.Text = TxtOldPassword.Text
                    PassEncrypt.Keystring = "SOLOMON"
                    PassEncrypt.Shrink
                    PassEncrypt.DoXor
                    
                    TxtOldPassword = PassEncrypt.Text
                
                PassEncrypt.Text = TxtNewPassword.Text
                PassEncrypt.Keystring = "SOLOMON"
                PassEncrypt.DoXor
                PassEncrypt.Stretch
                
                TxtNewPassword = PassEncrypt.Text
                'Compare if New and Old passwords are the same
                If TxtNewPassword = TxtOldPassword Then MsgBox "New Password and Old Password MUST not be the same. Please Re-enter", vbExclamation: Exit Sub
                    
                RsRecords!UserName = TxtUserName
                RsRecords!Password = PassEncrypt.Text
                RsRecords!CHANGEPASSWORD = False
                RsRecords.Update
                MsgBox "Password Changed Succesfully", vbInformation
                GlbPasswordResetLogin = False
                Unload Me
                MDIMain.Show
            End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
    TxtUserName = GlbCurrentUser
    centerform Me
End Sub

