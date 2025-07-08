VERSION 5.00
Begin VB.Form FrmDiscount 
   Caption         =   "Discount Adjustment"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7785
   Icon            =   "FrmDiscount.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   7575
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton CmdAccept 
         Caption         =   "Allow Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      Begin VB.TextBox TxtDiscountItem 
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   4935
      End
      Begin VB.TextBox TxtDiscountAmount 
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox TxtOriginalAmount 
         Height          =   495
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount Item"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount Amount"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Original Amount"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Label LblPatient 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "FrmDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAccept_Click()
    'UPDATE PRESCRIPTION
    If TxtDiscountAmount = "" Then Exit Sub
    MsgBox "ARE YOU SURE YOU WISH TO CONFIRM THE AMOUNT ADJUSTMENT?", vbYesNo + vbQuestion
    Conn.Execute "UPDATE PRESCRIPTION SET CASHAMOUNT = '" & TxtDiscountAmount & "' WHERE PRESCRIPTIONID = '" & GetID_NameFromCombo(TxtDiscountItem, 1) & "'"
    AuditTrail GlbCurrentUser, EnumCashier, GlbSysDate, Now, "Changed Price for " + GetID_NameFromCombo(TxtDiscountItem, 2) + " from " + TxtOriginalAmount + " To " + TxtDiscountAmount + " For CardNumber " + LblPatient
    Unload Me
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

