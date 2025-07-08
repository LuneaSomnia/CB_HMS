VERSION 5.00
Begin VB.Form FrmFullImage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   12240
      TabIndex        =   1
      Top             =   10200
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14295
      Begin VB.Image ImgFull 
         DragMode        =   1  'Automatic
         Height          =   9615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   14055
      End
   End
End
Attribute VB_Name = "FrmFullImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    centerform Me
End Sub
