VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Box"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOkay 
      Caption         =   "Okay"
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
      Left            =   12960
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      ItemData        =   "FrmAbout.frx":0000
      Left            =   11400
      List            =   "FrmAbout.frx":0037
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      Picture         =   "FrmAbout.frx":012C
      ScaleHeight     =   5535
      ScaleWidth      =   11055
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
   Begin VB.Line Line1 
      X1              =   11280
      X2              =   11280
      Y1              =   480
      Y2              =   5640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdOkay_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    centerform Me
End Sub
