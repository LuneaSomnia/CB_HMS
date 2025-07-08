VERSION 5.00
Begin VB.Form FrmExitStock 
   Caption         =   "Edit Number of Stock"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save Changes"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Edit Medicine Stock Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox TxtNewStock 
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox TxtCurrentStock 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   2415
      End
      Begin VB.ComboBox CboCategory 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   5655
      End
      Begin VB.ComboBox CboProduct 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   "New Stock Count"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Current Stock Count"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Medicine Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmExitStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Conn.Execute "UPDATE STOCK_ENTRY SET LASTSTOCKCOUNT = '" & TxtNewStock & "' WHERE PRODUCTID = '" & GetID_NameFromCombo(CboProduct, 1) & "'"
    MsgBox "Stock Count For  '" & GetID_NameFromCombo(CboProduct, 2) & "'  has been changed Succesfully to  '" & TxtNewStock & "'"
    
    AuditTrail GlbCurrentUser, EnumPharmacy, GlbSysDate, Time, "Edited Stock Count for " & GetID_NameFromCombo(CboProduct, 2) & " from " & TxtCurrentStock & " to " & TxtNewStock & ""
End Sub

Private Sub Form_Load()
    centerform Me
End Sub

Private Sub TxtNewStock_Change()
    ValidateDataType TxtNewStock, 0, "FrmExitStock", "TxtNewStock"
End Sub
