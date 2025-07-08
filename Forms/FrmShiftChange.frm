VERSION 5.00
Begin VB.Form FrmShifChange 
   Caption         =   "Shift Change Window"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   Icon            =   "FrmShiftChange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6015
      Begin VB.CommandButton CmbChangeShift 
         Caption         =   "Change Shift"
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
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox CboShift 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "FrmShiftChange.frx":0442
         Left            =   3240
         List            =   "FrmShiftChange.frx":044C
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Current Shift"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmShifChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbChangeShift_Click()
    If CboShift.Text = "" Then MsgBox "Please Select Shift before Saving Changes.", vbExclamation: Exit Sub
        If InStr(1, CboShift.Text, "DAY") >= 1 Then
            If VerifyAccess(GlbCurrentUser, "System Administrator") <> True Then MsgBox "You do not have Sufficient Privileges to access this Module", vbExclamation: Exit Sub
            Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = 0 WHERE ITEMNAME = 'DayShift_0_NightShift_1'"
        ElseIf InStr(1, CboShift.Text, "NIGHT") >= 1 Then
            Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = 1 WHERE ITEMNAME = 'DayShift_0_NightShift_1'"
        End If
        MsgBox "Shift Changes have been Updated Succesfully.", vbInformation
End Sub

Private Sub Form_Load()
    centerform Me
    If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'DayShift_0_NightShift_1'") = 1 Then
        CboShift.Text = "NIGHT SHIFT"
    Else
        CboShift.Text = "DAY   SHIFT"
    End If
End Sub
