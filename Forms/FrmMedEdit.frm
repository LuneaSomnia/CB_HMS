VERSION 5.00
Begin VB.Form FrmMedEdit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Medicine Edit"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "FrmMedEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   11535
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   10080
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
         Height          =   375
         Left            =   8520
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.TextBox TxtRemarks 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox CboDosage 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6240
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox TxtDays 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtType 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtMedName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Dosage Remarks"
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
         Left            =   7920
         TabIndex        =   5
         Top             =   230
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dosage"
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
         Left            =   6360
         TabIndex        =   4
         Top             =   230
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Days"
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
         Left            =   5160
         TabIndex        =   3
         Top             =   230
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   3600
         TabIndex        =   2
         Top             =   230
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Name"
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
         TabIndex        =   1
         Top             =   225
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmMedEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ArrDosage
Dim DosageNumber As Integer
Dim RsRecords As New ADODB.Recordset

Private Sub CboDosage_Click()
    DosageNumber = CboDosage.ListIndex
    StrArr = Split(ArrDosage, ",")
    TxtRemarks = FindRecord("DOSAGES", "DOSAGE", "DOSAGEID = '" & GetID_NameFromCombo(StrArr(DosageNumber), 1) & "'")
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
Dim lvNewQuantity As Integer
    If CboDosage = "" Then Exit Sub
    If TxtDays = "" Then Exit Sub
    lvNewQuantity = (Mid(CboDosage, 1, 1) * Right(CboDosage, 1)) * TxtDays
    
   If InStr(1, TxtType, "TABLET") = 1 Then 'CALCULATE IF ITS TABLET
        GlbUnitQuantity = (Mid(CboDosage, 1, 1) * Right(CboDosage, 1)) * TxtDays
   ElseIf InStr(1, TxtType, "CAPSULE") = 1 Then  'CALCULATE IF ITS CAPSULE
        GlbUnitQuantity = (Mid(CboDosage, 1, 1) * Right(CboDosage, 1)) * TxtDays
   Else  ' DO NOT CALCULATE IF ITS SYRUP OR CREAM OR POWDER OR OTHER
         GlbUnitQuantity = 1
   End If
    lvNewQuantity = GlbUnitQuantity
    Conn.Execute "UPDATE PRE_DRUGS_SALES SET DOSAGE = '" & CboDosage & "', DOSAGEREMARKS = '" & TxtRemarks & "', QUANTITY = '" & lvNewQuantity & "',DAYS = '" & TxtDays & "' WHERE PRODUCTID = '" & GlbMedEditID & "' AND SOLDBY = '" & GlbCurrentUser & "'"
    Unload Me
End Sub

Private Sub Form_Load()
    centerform Me
    PopulateDosage
End Sub
Private Sub PopulateDosage()
        'POPULATE COMBO FOR DOSAGE
        If RsRecords.State = 1 Then Set RsRecords = Nothing
        RsRecords.Open "SELECT * FROM DOSAGES", Conn, adOpenDynamic, adLockOptimistic
            ArrDosage = ""
            CboDosage.Clear
            With RsRecords
                While .BOF = False And .EOF = False
                    If ArrDosage = "" Then
                        ArrDosage = String(3 - Len(!DOSAGEID), "0") & !DOSAGEID & " - " & !DOSAGECODE
                    Else
                        ArrDosage = ArrDosage + "," & String(3 - Len(!DOSAGEID), "0") & !DOSAGEID & " - " & !DOSAGECODE
                    End If
                    'Debug.Print ArrDosage
                    'CboDosage.AddItem String(3 - Len(!DOSAGEID), "0") & !DOSAGEID & " - " & !DOSAGECODE
                    CboDosage.AddItem !DOSAGECODE
                    .MoveNext
                Wend
            End With
        RsRecords.Close
End Sub

