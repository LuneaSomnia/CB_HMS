VERSION 5.00
Begin VB.Form FrmComplains 
   Caption         =   "Complains list"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   Icon            =   "FrmComplains.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdInsetComplain 
      Caption         =   "Insert Complain"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   3480
      TabIndex        =   11
      Top             =   360
      Width           =   3495
      Begin VB.ListBox LstSubGroups 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4395
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton CmdEditSubGroups 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1260
         TabIndex        =   13
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton CmdDeleteSubGroup 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   4680
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   3255
      Begin VB.ListBox LstGroups 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6540
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton CmdDeleteGroup 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   6840
         Width           =   855
      End
      Begin VB.CommandButton CmdEditGroup 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1140
         TabIndex        =   8
         Top             =   6840
         Width           =   855
      End
      Begin VB.CommandButton CmdAddGroup 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   6840
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cl&ose"
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
      Left            =   7440
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox TxtComplains 
      Height          =   2055
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Sub Groups"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Complain Groups"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "FrmComplains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCombo As New ADODB.Recordset
Dim RsRecords As New ADODB.Recordset

Private Sub cmdadd_Click()
On Error GoTo ErrorHandler
   If LstGroups.Text = "" Then MsgBox "Please Select Complain Group before adding Sub Group.", vbInformation, "Select Group": Exit Sub
   FrmAddSubGroup.TxtGroupID = LstGroups.Text
   FrmAddSubGroup.TxtSubGroupID = "0675"
   
   FrmAddSubGroup.Show 1
   PopulateSubCategory Mid(LstGroups.Text, 1, 3)
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub PopulateSubCategory(ByVal ComplainID)
On Error GoTo ErrorHandler
    If ComplainID = "" Then Exit Sub
    'POPULATE LIST VIEW FOR SUB CATEGORY
    LstSubGroups.Clear
    RsCombo.Open "SELECT SubComplainID,SubGroupDescription FROM COMPLAIN_SUB_GROUPS WHERE COMPLAINID = '" & ComplainID & "' ORDER BY SUBGROUPDESCRIPTION", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                LstSubGroups.AddItem String(3 - Len(!SUBCOMPLAINID), "0") & !SUBCOMPLAINID & " - " & !SUBGROUPDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub PopulateCategory()
    'POPULATE LIST VIEW FOR MAIN CATEGORY
    LstGroups.Clear
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    RsCombo.Open "SELECT COMPLAINID,COMPLAINDESCRIPTION FROM COMPLAIN_GROUPS ORDER BY COMPLAINDESCRIPTION", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                LstGroups.AddItem String(3 - Len(!ComplainID), "0") & !ComplainID & " - " & !COMPLAINDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close
End Sub

Private Sub CmdAddGroup_Click()
    FrmAddGroup.Show 1
    PopulateCategory
End Sub

Private Sub CmdClear_Click()
    TxtComplains = ""
End Sub

Private Sub CMDCLOSE_Click()
    Unload Me
End Sub

Private Sub CmdDeleteGroup_Click()
    If LstGroups.Text = "" Then Exit Sub
    Conn.Execute "DELETE FROM COMPLAIN_GROUPS WHERE COMPLAINID = '" & GetID_NameFromCombo(LstGroups, 1) & "'"
    PopulateCategory
End Sub

Private Sub CmdDeleteSubGroup_Click()
    If LstSubGroups.Text = "" Then Exit Sub
    Conn.Execute "DELETE FROM COMPLAIN_SUB_GROUPS WHERE SUBCOMPLAINID = '" & GetID_NameFromCombo(LstSubGroups, 1) & "'"
    PopulateSubCategory GetID_NameFromCombo(LstGroups.Text, 1)
End Sub

Private Sub CmdEditGroup_Click()
On Error GoTo ErrorHandler
    BlnEditingComplains = True
    FrmAddGroup.TxtGroupID = GetID_NameFromCombo(LstGroups.Text, 1)
    FrmAddGroup.TxtComplainDescription = GetID_NameFromCombo(LstGroups.Text, 2)
    FrmAddGroup.Show 1
    PopulateCategory
    BlnEditingComplains = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdEditSubGroups_Click()
On Error GoTo ErrorHandler
    BlnEditingComplains = True
    FrmAddSubGroup.TxtGroupID = Mid(LstGroups.Text, 1, 3)
    FrmAddSubGroup.TxtSubGroupID = GetID_NameFromCombo(LstSubGroups.Text, 1)
    FrmAddSubGroup.TxtDescription = GetID_NameFromCombo(LstSubGroups.Text, 2)
    FrmAddSubGroup.Show 1
    PopulateSubCategory GetID_NameFromCombo(LstGroups.Text, 1)
    BlnEditingComplains = False
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdInsetComplain_Click()
On Error GoTo ErrorHandler
    If TxtComplains = "" Then
        TxtComplains = UCase(GetID_NameFromCombo(LstGroups.Text, 2)) & " - " & GetID_NameFromCombo(LstSubGroups.Text, 2)
    Else
        TxtComplains = TxtComplains + " , " + UCase(GetID_NameFromCombo(LstGroups.Text, 2)) & " - " & GetID_NameFromCombo(LstSubGroups.Text, 2)
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdOk_Click()
On Error GoTo ErrorHandler
    'FrmTreatment.TxtComplaints = TxtComplains
        Unload FrmComplains
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()

    PopulateCategory
    centerform Me

End Sub

Private Sub LstGroups_Click()
    PopulateSubCategory Mid(LstGroups.Text, 1, 3)
End Sub

