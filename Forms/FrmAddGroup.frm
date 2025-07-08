VERSION 5.00
Begin VB.Form FrmAddGroup 
   Caption         =   "Complain Maintenance"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   Icon            =   "FrmAddGroup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   6495
      Begin VB.CommandButton CmdAccept 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   6
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox TxtComplainDescription 
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox TxtGroupID 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0675"
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Complain Description"
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
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Complain ID"
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
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmAddGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAccept_Click()
On Error GoTo ErrorHandler
    If TxtComplainDescription = "" Then Exit Sub
        If BlnEditingComplains = False Then
            Conn.Execute "INSERT INTO COMPLAIN_GROUPS (COMPLAINDESCRIPTION)VALUES('" & TxtComplainDescription & "')"
        Else
            Conn.Execute "UPDATE COMPLAIN_GROUPS SET COMPLAINDESCRIPTION = '" & TxtComplainDescription & "' WHERE COMPLAINID = '" & TxtGroupID & "'"
            MsgBox "Record has been Updated Succesfully", vbInformation, "Edited"
        End If
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    centerform Me
    
    'LOAD SELECTED RECORD FROM LIST
    If BlnEditing = True Then
        TxtComplainDescription = Mid(FrmComplains.LstGroups.Text, 7, Len(FrmComplains.LstGroups.Text))
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub
