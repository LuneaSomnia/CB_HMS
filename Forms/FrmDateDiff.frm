VERSION 5.00
Begin VB.Form FrmDateDiff 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date Difference Notification"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "FrmDateDiff.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdSawa 
      Caption         =   "&Sawa !!"
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
      Left            =   12720
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "The Software date differs from the Machine date. if the Software date is inaccurate, then you may need to run &End Of Day Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "FrmDateDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub CmdSawa_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    centerform Me
End Sub
