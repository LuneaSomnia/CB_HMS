VERSION 5.00
Begin VB.Form FrmLabParameters 
   Caption         =   "Lab Test Maintenance"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   Icon            =   "FrmLabParameters.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CboLabTestGroups 
      Height          =   315
      Left            =   2040
      TabIndex        =   17
      Top             =   240
      Width           =   4575
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "Remove"
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
      Left            =   4680
      TabIndex        =   16
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
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
      Left            =   3600
      TabIndex        =   13
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
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
      Left            =   5760
      TabIndex        =   12
      Top             =   7680
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   2655
      Left            =   3480
      TabIndex        =   6
      Top             =   4920
      Width           =   3255
      Begin VB.ListBox LstSelected 
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   3480
      TabIndex        =   5
      Top             =   720
      Width           =   3255
      Begin VB.CommandButton CmdAddTest 
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
         TabIndex        =   10
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton CmdEditTest 
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
         TabIndex        =   9
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton CmdDeleteTest 
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
         TabIndex        =   8
         Top             =   3600
         Width           =   975
      End
      Begin VB.ListBox LstLabTests 
         Appearance      =   0  'Flat
         Height          =   3150
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3255
      Begin VB.CommandButton CmdAddGroup 
         Caption         =   "Add"
         Enabled         =   0   'False
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
         TabIndex        =   4
         Top             =   6360
         Width           =   975
      End
      Begin VB.CommandButton CmdEditGroup 
         Caption         =   "Edit"
         Enabled         =   0   'False
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
         TabIndex        =   3
         Top             =   6360
         Width           =   975
      End
      Begin VB.CommandButton CmdDeleteGroup 
         Caption         =   "Delete"
         Enabled         =   0   'False
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
         TabIndex        =   2
         Top             =   6360
         Width           =   975
      End
      Begin VB.ListBox LstGroups 
         Appearance      =   0  'Flat
         Height          =   5880
         ItemData        =   "FrmLabParameters.frx":0442
         Left            =   120
         List            =   "FrmLabParameters.frx":0449
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Lab Test Groups"
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
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FrmLabParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsLab As New ADODB.Recordset
Dim RsRecords As New ADODB.Recordset

Private Sub CboLabTestGroups_Click()
    'POPULATE LIST BOX FOR LAB TEST SUB-GROUPS
    LstGroups.Clear
    RsLab.Open "SELECT GROUPID, SUBGROUPID, SUBGROUPDESCRIPTION FROM LAB_TEST_SUB_GROUPS WHERE GROUPID = '" & GetID_NameFromCombo(CboLabTestGroups, 1) & "' ORDER BY SUBgroupID ASC", Conn, adOpenStatic, adLockOptimistic
    
        With RsLab
            While .BOF = False And .EOF = False
                LstGroups.AddItem String(2 - Len(!SUBGROUPID), "0") & !SUBGROUPID & " - " & !SUBGROUPDESCRIPTION
                .MoveNext
            Wend
        End With
    RsLab.Close
End Sub

Private Sub CmdAddTest_Click()
    If LstGroups.Text = "" Then MsgBox "Please Select Sub-Group before Adding", vbInformation: Exit Sub
    FrmAddLabTest.Show 1
    PopulateLabTests
End Sub

Private Sub CmdClear_Click()
    LstSelected.Clear
End Sub

Private Sub CMDCLOSE_Click()
    Unload Me
End Sub

Private Sub PopulateLabTests()
    LstLabTests.Clear
    If CboLabTestGroups = "" Then Exit Sub
    'POPULATE COMBO FOR LAB TESTS
    RsLab.Open "SELECT * FROM LABTESTPARAMETERS WHERE TESTGROUP = '" & Val(GetID_NameFromCombo(CboLabTestGroups, 1)) & "' AND TESTSUBGROUP = '" & Val(GetID_NameFromCombo(LstGroups, 1)) & "' ORDER BY TESTID", Conn, adOpenStatic, adLockOptimistic
    
        With RsLab
            While .BOF = False And .EOF = False
                LstLabTests.AddItem String(2 - Len(!TESTID), "0") & !TESTID & " - " & !TESTDESCRIPTION
                .MoveNext
            Wend
        End With
    RsLab.Close
End Sub

Private Sub CmdDeleteTest_Click()
    Conn.Execute "DELETE FROM LABTESTPARAMETERS WHERE TESTID = '" & GetID_NameFromCombo(LstLabTests.Text, 1) & "'"
    PopulateLabTests
End Sub

Private Sub CmdEditTest_Click()
    BlnEditingLabTests = True
    
    FrmAddLabTest.TxtLabTestDescription = LstLabTests.Text
    FrmAddLabTest.Show 1
    PopulateLabTests
End Sub

Private Sub CmdOk_Click()
    For i = 0 To LstSelected.ListCount
        'If FrmTreatment.LstLabTests = "" Then
            FrmTreatment.LstLabTests.AddItem LstSelected.List(i)
       ' Else
            'FrmTreatment.TxtLabRequest = FrmTreatment.TxtLabRequest + LstSelected.List(i) & vbCrLf
        'End If
    Next
    Unload Me
End Sub

Private Sub CmdRemove_Click()
    If LstSelected.ListIndex = -1 Then Exit Sub
    LstSelected.RemoveItem LstSelected.ListIndex
End Sub

Private Sub Form_Load()
    PopulateLabTests
    centerform Me
    
    'POPULATE COMBO FOR LAB TEST GROUPS
    RsLab.Open "SELECT TESTGROUPID, TESTGROUPDESCRIPTION FROM LAB_TESTGROUPS ORDER BY TESTgroupID ASC", Conn, adOpenStatic, adLockOptimistic
    
        With RsLab
            While .BOF = False And .EOF = False
                CboLabTestGroups.AddItem String(2 - Len(!TESTGROUPID), "0") & !TESTGROUPID & " - " & !TESTGROUPDESCRIPTION
                .MoveNext
            Wend
        End With
    RsLab.Close
    
    'POPULATE TEST GROUPS COMBO BOX
''    If RsRecords.State = 1 Then Set RsRecords = Nothing
''
''    RsRecords.Open "SELECT * FROM LAB_TESTGROUPS", Conn, adOpenStatic, adLockOptimistic
''        While RsRecords.EOF = False
''            CboLabTestGroups.AddItem RsRecords!TESTGROUPID & " - " & RsRecords!TESTGROUPDESCRIPTION
''            RsRecords.MoveNext
''        Wend
''    RsRecords.Close
    
End Sub

Private Sub LstGroups_Click()
    PopulateLabTests
End Sub

Private Sub LstLabTests_DblClick()
    LstSelected.AddItem GetID_NameFromCombo(LstLabTests, 2)
End Sub
