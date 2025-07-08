VERSION 5.00
Begin VB.Form FrmLaboratory 
   Caption         =   "Laboratory Request and Results"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12120
   Icon            =   "FrmLaboratory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "Post Patient"
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   7920
      Width           =   9495
      Begin VB.CommandButton CmdPost 
         Caption         =   "POST"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptDoctor 
         Caption         =   "TO DOCTOR"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   6120
         TabIndex        =   25
         Top             =   400
         Width           =   1335
      End
      Begin VB.OptionButton OptCashier 
         Caption         =   "TO CASHIER"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4560
         TabIndex        =   24
         Top             =   400
         Width           =   1335
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO OBSERVATION"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2355
         TabIndex        =   23
         Top             =   400
         Width           =   1935
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO CONSULTATION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   400
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Payement Details"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   6960
      Width           =   9495
      Begin VB.ComboBox CboPaymentMode 
         Height          =   315
         ItemData        =   "FrmLaboratory.frx":0442
         Left            =   2880
         List            =   "FrmLaboratory.frx":044C
         TabIndex        =   18
         Text            =   "1 - CASH"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtLabAmount 
         Height          =   315
         Left            =   6720
         TabIndex        =   16
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Payment Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Laboratory Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Save && Post "
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   9720
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
      Begin VB.CommandButton CmdConclude 
         Caption         =   "Conclude Treatment"
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Close"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   6360
         Width           =   2055
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lab Results"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   9495
      Begin VB.TextBox TxtResult 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lab Request "
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   9495
      Begin VB.ListBox LstRequests 
         Height          =   2085
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   27
         Top             =   360
         Width           =   9255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Details"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.TextBox TxtRequestDate 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8760
         TabIndex        =   13
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox TxtDoc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TxtCardNumber 
         Enabled         =   0   'False
         Height          =   405
         Left            =   4800
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtPatientNames 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Request Date"
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
         Left            =   8760
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Requesting Doctor"
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
         Left            =   6600
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Card Number"
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
         Left            =   4800
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Patients Name"
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
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmLaboratory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsLab As New ADODB.Recordset

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPost_Click()
On Error GoTo ErrorHandler
    'IF THIS WAS A MIS POST TO LAB THEN JUST SEND IT BACK WITHOUT CHANGES.
    If LstRequests.Text = "NO LAB TEST REQUEST(S) SENT !!!" Then GoTo SendWithoutSaving
    
    'SAVE IF NOT PREVIOUSELY POSTED.
    Conn.Execute "UPDATE COMPLAINS SET LABRESULTS = '" & TxtResult & "',LABTECH = '" & GlbCurrentUser & "',INUSE = '0' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
    'ADD LAB-TESTS TO INVOICE. GET THE DATA FROM COMPLAINS AND INSERT INTO PRESCRIPTION TABLE
    If TxtLabAmount = "" Then MsgBox ("Amount Must be entered before Sending to Lab Results"), vbCritical: Exit Sub
    RsLab.Open "SELECT * FROM COMPLAINS WHERE CARDNUMBER = '" & TxtCardNumber & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsLab
            If .EOF = False Then
                Conn.Execute "INSERT INTO PRESCRIPTION(CARDNUMBER,VISITNUMBER,VISITDATE,BILLINGCO,CODE,DESCRIPTION,CASHAMOUNT,PAYMENTMODE)" & _
                             "VALUES ('" & TxtCardNumber & "','" & StrDocVisitNumber & "','" & Format(StrDocVisitDate, "ddmmmyyyy") & "','001','002','Lab Test(s)','" & TxtLabAmount & "','" & GetID_NameFromCombo(CboPaymentMode, 1) & "')"
            MsgBox "Lab Tests Included in Invoice", vbInformation
            End If
        End With
        
SendWithoutSaving:
    'NOW SEND TO REQUESTING DOCTOR
    If OptCashier.Item(3).Value = True Then
        DMY = SendPatient(EnumCashier, StrDocCardNo, Format(StrDocVisitDate))
    ElseIf OptDoctor.Item(4).Value = True Then
        DMY = SendPatient(EnumDoctors, StrDocCardNo, Format(StrDocVisitDate))
    ElseIf OptObservation.Item(2).Value = True Then
        DMY = SendPatient(EnumObservation, StrDocCardNo, Format(StrDocVisitDate))
    End If
    ClearText FrmLaboratory
    CmdPost.Enabled = False
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdSave_Click()
    'If TxtResult <> "" Then
    '''If TxtLabAmount = "" Then MsgBox "Please enter amount to be chaged before Saving", vbInformation: Exit Sub
    lvBillingCompany = "001" 'Todo
        Conn.Execute "UPDATE COMPLAINS SET LABRESULTS = '" & TxtResult & "' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
        'CHECK IF RECORD HAS ALREADY BEEN CHARGED AND UPDATE, IF NOT THEN INSERT
        
'''        If TxtDoc <> "FROM CASHIER" Then
'''            If CboPaymentMode = "1 - CASH" Then
'''            'CASH
'''                Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER, VISITNUMBER,BILLINGCO,VISITDATE,CODE,DESCRIPTION,QUANTITY,CASHAMOUNT,PAYDATE,PAYMENTMODE,CASHIER)" & _
'''                             "VALUES('" & StrDocCardNo & "', '" & StrDocVisitNumber & "','" & lvBillingCompany & "','" & Format(Date, "DD MMM YYYY") & "','002', 'Lab Test','1' ,'" & TxtLabAmount & "','" & Format(Date, "DD MMM YYYY") & "','" & Mid(CboPaymentMode, 1, 1) & "','" & GlbCurrentUser & "')"
'''            Else
'''            'CREDIT
'''                Conn.Execute "INSERT INTO PRESCRIPTION  (CARDNUMBER, VISITNUMBER,BILLINGCO,VISITDATE,CODE,DESCRIPTION,QUANTITY,CREDITAMOUNT,PAYDATE,PAYMENTMODE, CASHIER)" & _
'''                             "VALUES('" & StrDocCardNo & "', '" & StrDocVisitNumber & "','" & lvBillingCompany & "','" & Format(Date, "DD MMM YYYY") & "','002', 'Lab Test','1' ,'" & TxtLabAmount & "','" & Format(Date, "DD MMM YYYY") & "','" & Mid(CboPaymentMode, 1, 1) & "','" & GlbCurrentUser & "')"
'''            End If
'''        End If
        
        MsgBox "Test Result saved Saved Succesfully But Not Posted To Requesting Doctor", vbInformation
        CmdPost.Enabled = True
    'End If
End Sub

Private Sub CmdToCashier_Click()
    If TxtLabAmount = "" Then MsgBox "Please enter amount to be chaged before Saving", vbInformation: Exit Sub
    'SAVE IF NOT PREVIOUSELY POSTED.
    Conn.Execute "UPDATE COMPLAINS SET LABRESULTS = '" & TxtResult & "' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
    'NOW SEND TO REQUESTING DOCTOR
    DMY = SendPatient(EnumCashier, StrDocCardNo, Format(StrDocVisitDate))
    ClearText FrmLaboratory
End Sub

Private Sub CmdConclude_Click()
    If LstRequests.ListCount < 1 Then Exit Sub
    Conn.Execute "UPDATE COMPLAINS SET TOLABORATORY = '0' WHERE CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'"
    ClearText FrmLaboratory
    MsgBox "Treatment Completed Succesfully", vbInformation
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim StrArrRequests
    centerform Me
    GlbCurrentForm = EnumLab
    If StrDocCardNo = "" Then GreyOut FrmLaboratory: CmdSave.Enabled = False: CmdConclude.Enabled = False: Exit Sub
    If StrDocVisitNumber = "" Then GreyOut FrmLaboratory: CmdSave.Enabled = False: CmdConclude.Enabled = False: Exit Sub
    Dim lvFirstName, lvSecondName, lvSurname As String
    If RsLab.State = 1 Then Set RsLab = Nothing
    RsLab.Open "SELECT * FROM COMPLAINS WHERE CARDNUMBER = '" & StrDocCardNo & "' AND visitnumber = '" & StrDocVisitNumber & "'", Conn, adOpenStatic, adLockBatchOptimistic
        If RsLab.EOF = False Then
            lvFirstName = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
            lvSecondName = FindRecord("PATIENT_DETAILS", "SECONDNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
            lvSurname = FindRecord("PATIENT_DETAILS", "SURNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
            TxtPatientNames = lvSurname & " " & lvFirstName & " " & lvSecondName
            TxtCardNumber = StrDocCardNo
            If IsNull(RsLab!DOCTOR) = True Then
                TxtDoc = "NONE"
                
                
                Else
                TxtDoc = RsLab!DOCTOR
            End If
            TxtRequestDate = RsLab!VisitDate
            
            If IsNull(RsLab!LABREQUEST) Then LstRequests.AddItem "NO LAB TEST REQUEST(S) SENT !!!": CmdPost.Enabled = True: CmdSave.Enabled = False: Exit Sub

            
            StrArrRequests = Split(RsLab!LABREQUEST, Chr(13))
            For i = 0 To UBound(StrArrRequests) '-1
                LstRequests.AddItem StrArrRequests(i)
            Next
            If RsLab!LABRESULTS <> Null Then
                TxtResult.Text = RsLab!LABRESULTS
            End If
            '***********grey out hapa na pale
            If TxtDoc = "FROM CASHIER" Then
                OptCashier.Item(3).Enabled = False
                OptDoctor.Item(4).Enabled = True
                TxtLabAmount.Enabled = False
                CboPaymentMode.Enabled = False
            Else
            
            End If
            '***********
            TxtLabAmount = StrLabAmount
        End If
        If TxtCardNumber = "CARD NUMBER" Then
            CmdSave.Enabled = False
            CmdPost.Enabled = False
            GreyOut FrmLaboratory
''            OptCashier.Item(1).Enabled = True
''            CmdPost.Enabled = True
            Frame8.Enabled = True
        End If
    RsLab.Close
    centerform Me
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description
    'Resume
End Sub

