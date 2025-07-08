VERSION 5.00
Begin VB.Form FrmQuantity 
   Caption         =   "Quantity"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
   Icon            =   "FrmQuantity.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAccept 
      Caption         =   "Accept"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quantity By Distribution Unit "
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox TxtTotalAmount 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox TxtPrice 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox TxtQuantity 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtUnit 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Total Amount Due"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Price Per Unit"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Desired Number of Units"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Minumum Distribution Unit "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmQuantity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAccept_Click()
   GlbUnitQuantity = TxtQuantity
   Unload Me
End Sub

Private Sub TxtQuantity_Change()
    ValidateDataType TxtQuantity, 0, "FrmQuantity", "TxtQuantity"
    If TxtPrice = "" Or TxtUnit = "" Or TxtQuantity = "" Then TxtTotalAmount = "": Exit Sub
    If Not IsNumeric(TxtQuantity) Then Exit Sub
    TxtTotalAmount = (TxtPrice / TxtUnit) * TxtQuantity
End Sub
Private Sub CalculateTotals()

End Sub
