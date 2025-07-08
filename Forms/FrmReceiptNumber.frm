VERSION 5.00
Begin VB.Form FrmReceiptNumber 
   Caption         =   "Direct Sale Receipt Details"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   9480
      TabIndex        =   5
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   280
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "RECEIPT NUMBER DETAILS"
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.TextBox TxtCardNumber 
         Height          =   315
         Left            =   7080
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox ChkPaid 
         Caption         =   "Customer already Paid in Full"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   9
         Top             =   0
         Width           =   3015
      End
      Begin VB.TextBox TxtSaleNumber 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TxtCustomerName 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox TxtReceiptNumber 
         Height          =   315
         Left            =   7080
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Sale Number"
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Customer Name"
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
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Receipt Number"
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
         Left            =   7080
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmReceiptNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()

End Sub

Private Sub ChkPaid_Click()
    If ChkPaid.Value = 1 Then
        TxtCardNumber.Visible = True
        TxtCustomerName.Enabled = False
        TxtReceiptNumber.Enabled = False
    Else
        TxtCardNumber.Visible = False
    End If
End Sub

Private Sub CmdSave_Click()
    If ChkPaid.Value = 0 Then
        If Not IsNumeric(TxtReceiptNumber) = True Then MsgBox "Only Numbers are allowed as Receipt Number", vbInformation: Exit Sub
        If TxtCustomerName.Text = "" Then MsgBox "Please Enter the Customer Name as It appears on the Receipt Before Proceeding", vbExclamation: Exit Sub
        If TxtReceiptNumber.Text = "" Then MsgBox "Please Enter the Receipt Number as it appears on the Receipt Before Proceeding", vbExclamation: Exit Sub
    End If
    If ChkPaid.Value = 1 And TxtCardNumber = "" Then MsgBox "Please Enter the Patient Card Number Before Proceeding", vbExclamation: TxtCardNumber.SetFocus: Exit Sub
    TxtCustomerName.Text = "Customer Already Paid in full - Card No " & Trim(TxtCardNumber)
    TxtReceiptNumber = 2
    'UPDATE DRUGS_SALES TO SHOW DRUGSUBMITTED = TRUE
    Conn.Execute "UPDATE DRUGS_SALES SET DRUGSUBMITTED = 1,PAYMENTMODE = 1,CUSTOMERNAME = '" & UCase(TxtCustomerName) & "',RECEIPTNUMBER = '" & TxtReceiptNumber & "' WHERE SALENUMBER = '" & TxtSaleNumber & "'"
    BlnReceiptDetails = True
    Unload Me
End Sub

Private Sub Form_Load()
    centerform Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TxtCustomerName.Text = "" Then
        Resp = MsgBox("Are you sure you wish to Exit without Saving Customer Name?", vbQuestion + vbYesNo)
        If Resp = vbYes Then
            BlnReceiptDetails = False: Exit Sub
        Else
            Cancel = True
            Exit Sub
        End If
    End If
    If TxtReceiptNumber.Text = "" Then
        Resp = MsgBox("Are you sure you wish to Exit without Saving Receipt Number?", vbQuestion + vbYesNo)
        If Resp = vbYes Then
            BlnReceiptDetails = False: Exit Sub
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub TxtReceiptNumber_Change()
    ValidateDataType TxtReceiptNumber, 0, "FrmReceiptNumber", "TxtReceiptNumber"
    
End Sub
