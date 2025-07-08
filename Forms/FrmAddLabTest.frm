VERSION 5.00
Begin VB.Form FrmAddLabTest 
   Caption         =   "Lab Test Maintenance"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   Icon            =   "FrmAddLabTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   7095
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve Test"
         Height          =   495
         Left            =   4200
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test Name"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox TxtLabTestAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox TxtLabTestDescription 
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   5895
      End
      Begin VB.TextBox TxtLabTestID 
         Height          =   375
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0675"
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Amount Charged"
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
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Lab Test Description"
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
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Lab Test ID"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmAddLabTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset

Private Sub CmdSave_Click()
    If TxtLabTestDescription = "" Then Exit Sub
    TxtLabTestAmount = 0
    If BlnEditingLabTests = True Then
        Conn.Execute "UPDATE LABTESTPARAMETERS SET TESTDESCRIPTION = '" & GetID_NameFromCombo(TxtLabTestDescription, 2) & "',AMOUNT = '" & TxtLabTestAmount & "' WHERE TESTDESCRIPTION = '" & GetID_NameFromCombo(FrmLabParameters.LstLabTests, 2) & "'"
        BlnEditingLabTests = False
    Else
        Conn.Execute "INSERT INTO LABTESTPARAMETERS(TESTGROUP,TESTSUBGROUP,TESTDESCRIPTION,AMOUNT)VALUES('" & Val(GetID_NameFromCombo(FrmLabParameters.CboLabTestGroups, 1)) & "','" & Val(GetID_NameFromCombo(FrmLabParameters.LstGroups.Text, 1)) & "','" & UCase(TxtLabTestDescription) & "','" & TxtLabTestAmount & "')"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    centerform Me
End Sub

