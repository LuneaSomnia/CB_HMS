VERSION 5.00
Begin VB.Form FrmAddSubGroup 
   Caption         =   "Sub group Maintenance"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   Icon            =   "FrmAddSubGroup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   6495
      Begin VB.CommandButton CmdAccept 
         Caption         =   "Save"
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox TxtGroupID 
         Height          =   405
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox TxtSubGroupID 
         Height          =   405
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox TxtDescription 
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Complain Group"
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
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Sub Group ID N&o"
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
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Sub Group Description"
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
         TabIndex        =   4
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmAddSubGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAccept_Click()
On Error GoTo ErrorHandler
    If TxtDescription = "" Then Exit Sub
    If BlnEditingComplains = False Then
        Conn.Execute "INSERT INTO COMPLAIN_SUB_GROUPS (COMPLAINID,SubGroupDescription)VALUES('" & Mid(TxtGroupID, 1, 3) & "','" & TxtDescription & "')"
    Else
        Conn.Execute "UPDATE COMPLAIN_SUB_GROUPS SET SUBGROUPDESCRIPTION = '" & TxtDescription & "' WHERE SUBCOMPLAINID = '" & TxtSubGroupID & "'"
    End If
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

    centerform Me
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub


