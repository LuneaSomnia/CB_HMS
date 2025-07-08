VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3255
   ClientLeft      =   6660
   ClientTop       =   4560
   ClientWidth     =   9015
   ForeColor       =   &H000080FF&
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1923.161
   ScaleMode       =   0  'User
   ScaleWidth      =   8464.597
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TmrLogin 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   240
      Top             =   2520
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login Credentials"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   6735
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   2025
         TabIndex        =   0
         Top             =   480
         Width           =   4590
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2025
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1440
         Width           =   4590
      End
      Begin VB.TextBox TxtFullName 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2025
         TabIndex        =   1
         Top             =   945
         Width           =   4590
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   720
         TabIndex        =   9
         Top             =   495
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   510
      Left            =   7080
      TabIndex        =   4
      Top             =   2520
      Width           =   1740
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   1200
      Width           =   1740
   End
   Begin VB.Label LbLSysDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   8775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nairobi Cardiovascular Clinic"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Dim RsLogin As New ADODB.Recordset
Dim EcryptDecrypt As New Encryption.EncryptDecrypt
Private Sub CmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    End
End Sub

Private Sub CmdOk_Click()
    On Error Resume Next
    'Establish System Connection
        
    'check for correct Credentials
    If RsLogin.State = 1 Then Set RsLogin = Nothing
    RsLogin.Open "Select * from Users where username = '" & Trim(txtUserName) & "'", Conn, adOpenStatic, adLockOptimistic
        If RsLogin.EOF = False Then
            
            'Decrypt password for comparison
                EcryptDecrypt.Text = txtPassword.Text
                EcryptDecrypt.Keystring = "SOLOMON"
                EcryptDecrypt.DoXor
                EcryptDecrypt.Stretch
                TxtFullName = RsLogin!fullname
                
            If RsLogin!Password = EcryptDecrypt.Text Then
                'place code to here to pass the
                'success to the calling sub
                'setting a global var is the easiest
                GlbCurrentUser = txtUserName
                LoginSucceeded = True
                TmrLogin.Enabled = False
                txtPassword = ""
                
                'CHANGE
                    If FindRecord("USERS", "CHANGEPASSWORD", "USERNAME = '" & txtUserName & "'") = True Then
                        Me.Hide
                        GlbPasswordResetLogin = True
                        FrmConfirmPassword.Show
                    Else
                        Me.Hide
                        MDIMain.Show
                    End If
            Else
                MsgBox "Invalid Password", vbInformation
                txtPassword = ""
            End If
        Else
                MsgBox "Invalid User Name or Password, Please try again!", , "Login"
                txtPassword.SetFocus
                'SendKeys "{Home}+{End}"
        End If
End Sub

Private Sub Form_Load()
    LbLSysDate = Format(GlbSysDate, "DDDD,  DD - MMMM - YYYY")
    TmrLogin.Enabled = True
End Sub

Private Sub TmrLogin_Timer()
    If LbLSysDate.Caption = "" Then
        LbLSysDate = Format(GlbSysDate, "DDDD,  DD - MMMM - YYYY")
    Else
        LbLSysDate = ""
    End If
End Sub

Private Sub txtUserName_LostFocus()
    If RsLogin.State = 1 Then Set RsLogin = Nothing
    RsLogin.Open "Select * from Users where username = '" & Trim(txtUserName) & "'", Conn, adOpenStatic, adLockOptimistic
        If RsLogin.EOF = False Then TxtFullName = RsLogin!fullname
        txtUserName = UCase(txtUserName)
    Set RsLogin = Nothing
End Sub
