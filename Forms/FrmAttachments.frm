VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAttachments 
   Caption         =   "Attachments "
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7680
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Width           =   4815
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   2640
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   12375
      Begin VB.TextBox TxtPatientsNames 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   8535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   5040
      TabIndex        =   2
      Top             =   1080
      Width           =   7455
      Begin VB.PictureBox Picture1 
         Height          =   7455
         Left            =   120
         ScaleHeight     =   7395
         ScaleWidth      =   7155
         TabIndex        =   13
         Top             =   360
         Width           =   7215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attachment Type"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
      Begin VB.CommandButton CmdMRI 
         Caption         =   "MRI"
         Height          =   495
         Left            =   2640
         TabIndex        =   10
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton CmdCTScan 
         Caption         =   "CT Scan"
         Height          =   495
         Left            =   2640
         TabIndex        =   9
         Top             =   2360
         Width           =   2055
      End
      Begin VB.OptionButton Option4 
         Caption         =   "MRI"
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
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ct Scan"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2480
         Width           =   1575
      End
      Begin VB.CommandButton CmdUltrSound 
         Caption         =   "Ultrasound"
         Height          =   495
         Left            =   2640
         TabIndex        =   6
         Top             =   1480
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ultrasound"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1600
         Width           =   2175
      End
      Begin VB.CommandButton CmdXray 
         Caption         =   "Upload X Ray"
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "X Ray Results"
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
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub centerform(X As Form)
    Screen.MousePointer = 11
    If X.MDIChild = False Then
        On Error Resume Next: X.Top = (Screen.Height / 2) - X.Height / 2
        X.Left = (Screen.Width / 2) - (X.Width / 2)
    Else
        'MDI.Arrange cascade
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    centerform Me
End Sub
