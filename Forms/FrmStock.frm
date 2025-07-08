VERSION 5.00
Begin VB.Form FrmStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medicine Inventory"
   ClientHeight    =   2610
   ClientLeft      =   1095
   ClientTop       =   375
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5775
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5775
      TabIndex        =   12
      Top             =   2310
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1213
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Valid"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1660
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SalePrice"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ReOrderLevel"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Text            =   "30"
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MinLevel"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Text            =   "10"
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ProductName"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ProductGroupID"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Text            =   "4550"
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Valid:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "SalePrice:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ReOrderLevel:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MinLevel:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ProductName:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ProductGroupID:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "FrmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsStock As New ADODB.Recordset

Private Sub Form_Load()
    centerform Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

''''Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
''''  'This is where you would put error handling code
''''  'If you want to ignore errors, comment out the next line
''''  'If you want to trap them, add code here to handle them
''''  MsgBox "Data error event hit err:" & Description
''''End Sub
''''
''''Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
''''  'This will display the current record position for this recordset
''''  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
''''End Sub
''''
''''Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
''''  'This is where you put validation code
''''  'This event gets called when the following actions occur
''''  Dim bCancel As Boolean
''''
''''  Select Case adReason
''''  Case adRsnAddNew
''''  Case adRsnClose
''''  Case adRsnDelete
''''  Case adRsnFirstChange
''''  Case adRsnMove
''''  Case adRsnRequery
''''  Case adRsnResynch
''''  Case adRsnUndoAddNew
''''  Case adRsnUndoDelete
''''  Case adRsnUndoUpdate
''''  Case adRsnUpdate
''''  End Select
''''
''''  If bCancel Then adStatus = adStatusCancel
''''End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    Set RsStock = Nothing
    Dim strID
    RsStock.Open "SELECT * FROM PRODUCTS ORDER BY PRODUCTID DESC", Conn, adOpenStatic, adLockOptimistic
        strID = RsStock!PRODUCTID
        With RsStock
                    .AddNew
                    !PRODUCTID = strID + 1
                    !ProductName = txtFields(1)
                    !saleprice = txtFields(4)
                    !amount = !saleprice
                    .Update
                    MsgBox "Medicine Maintained Successfully ", vbInformation
        End With
  Exit Sub
AddErr:
  MsgBox Err.Description
  'Resume
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
'  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  'datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
''    Set RsStock = Nothing
''    RsStock.Open "SELECT * FROM PRODUCTS WHERE PROVIDERID = '" & TxtCompanyCode & "'", Conn, adOpenStatic, adLockOptimistic
''        With RsStock
''
''                    !ProductName = txtFields(1)
''                    !saleprice = txtFields(4)
''                    !amount = !saleprice
''                    .Update
''                    MsgBox "Billing Company Successfully Maintained", vbInformation
''        End With
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

