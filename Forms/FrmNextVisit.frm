VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNextVisit 
   BackColor       =   &H0000C0C0&
   Caption         =   "Next Visit"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C0C0&
      Height          =   3015
      Left            =   11040
      TabIndex        =   5
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton CmdDrop 
         BackColor       =   &H0000C0C0&
         Caption         =   "Drop"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      Caption         =   "TCA "
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.TextBox TxtProceedure 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1560
         Width           =   8055
      End
      Begin MSComCtl2.DTPicker DtNextVisit 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   8421504
         Format          =   130023425
         CurrentDate     =   43271
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label LblFullNames 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   8055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Visit Proceedure :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Next Visit Date :"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmNextVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDrop_Click()
On Error GoTo Errorhandler
    Unload Me
    'GlbDropCancel = False
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub CmdSave_Click()
On Error GoTo Errorhandler
    If TxtProceedure = "" Then MsgBox "Proceedure Cannot be blank", vbCritical: Exit Sub
    Conn.Execute "UPDATE COMPLAINS SET  NEXTPROCEEDURE = '" & TxtProceedure & "' WHERE  VISITNUMBER = '" & GlbDropNumber & "'"
    MsgBox "Proceedure Details Saved Succesfully", vbInformation, "TCA"
    GlbDropCancel = False
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Errorhandler
    GlbDropCancel = True
    'If GlbDropView = True Then Me.Width = 11235
     centerform Me
   
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

