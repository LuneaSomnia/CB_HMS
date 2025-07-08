VERSION 5.00
Begin VB.Form FrmNonpatients 
   Caption         =   "Non Patient Visits"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   Icon            =   "FrmNonPatients.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   7680
      TabIndex        =   8
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox TxtCardNumber 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtVisitorName 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   960
         Width           =   4695
      End
      Begin VB.ComboBox CboDoctors 
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   2280
         Width           =   4695
      End
      Begin VB.TextBox TxtReason 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Text            =   "Official"
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Card Number"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Person to be Seen"
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
         TabIndex        =   7
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Rea 
         Alignment       =   1  'Right Justify
         Caption         =   "Reason for Visit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Visitor's Name"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmNonpatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdNew_Click()
On Error GoTo ERRORHANDLER
    If CmdNew.Caption = "New" Then
        ClearText FrmNonpatients
        TxtCardNumber = "9999/99"
        TxtReason = "Official"
        CmdNew.Caption = "Save"
    Else
        If TxtVisitorName = "" Then MsgBox "Please enter the Visitor Name before Saving", vbInformation: Exit Sub
        If TxtReason = "" Then TxtReason = "Official"
        If CboDoctors = "" Then MsgBox "Please Select doctor to be seen before Saving", vbInformation: Exit Sub
        
        Conn.Execute "INSERT INTO COMPLAINS (CARDNUMBER,DOCTOR,TODOCTORS,VISITDATE,TIMEIN)VALUES('" & TxtCardNumber & "','" & CboDoctors & "','TRUE','" & GlbSysDate & "','" & Time & "')"
        lvVisitNumber = FindRecord("COMPLAINS", "VISITNUMBER", "CARDNUMBER = '9999/99' ORDER BY VISITNUMBER DESC")
        Conn.Execute "INSERT INTO NO_OFFICIAL_APPOINTMENTS (CARDNUMBER,VISITNUMBER,NAMES)VALUES('" & TxtCardNumber & "','" & lvVisitNumber & "','" & TxtVisitorName & "')"
        Conn.Execute "UPDATE GENERALPARAMs SET ITEMVALUE = '" & lvVisitNumber & "' WHERE ITEMNAME = 'NonPatientsRunNumber'"
        CmdNew.Caption = "New"
        MsgBox "Visit Saved Succesfully", vbInformation
        
    End If
ERRORHANDLER:
    Exit Sub
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ERRORHANDLER
    centerform Me
    
   'POPULATE DOCTORS COMBO BOX
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    CboDoctors.AddItem "<< ANY >>"
    RsRecords.Open "SELECT * FROM PROFILES WHERE RIGHTS = 'DOCTORS' AND ACCESS = 'TRUE'", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboDoctors.AddItem RsRecords!UserName
            RsRecords.MoveNext
        Wend
Exit Sub
ERRORHANDLER:
MsgBox Err.Description
End Sub

