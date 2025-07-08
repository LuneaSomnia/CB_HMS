VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmReports 
   Caption         =   "Reports"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   Icon            =   "FrmReports.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   14280
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   7200
      TabIndex        =   27
      Top             =   3240
      Width           =   6975
      Begin VB.ComboBox CboProduct 
         Height          =   315
         Left            =   3720
         TabIndex        =   31
         Top             =   480
         Width           =   3135
      End
      Begin VB.ComboBox CboCategory 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Medicine Description"
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
         Left            =   3720
         TabIndex        =   30
         Top             =   220
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Category"
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
         TabIndex        =   28
         Top             =   220
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   7200
      TabIndex        =   16
      Top             =   2280
      Width           =   6975
      Begin VB.TextBox TxtCardNumber 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Card Number"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   7200
      TabIndex        =   10
      Top             =   960
      Width           =   6975
      Begin VB.OptionButton Option6 
         Caption         =   "Report All"
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
         Left            =   5400
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptDateRange 
         Caption         =   "Report By Date Range"
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
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton OptSingleDate 
         Caption         =   "Report By Single Date"
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
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTEndDate 
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   720
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         _Version        =   393216
         Format          =   114819073
         CurrentDate     =   39163
      End
      Begin MSComCtl2.DTPicker DTStartDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   114819073
         CurrentDate     =   39163
      End
      Begin VB.TextBox txtDate2 
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Output Order"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   7200
      TabIndex        =   9
      Top             =   4320
      Width           =   3645
      Begin VB.OptionButton Option2 
         Caption         =   "Descending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   400
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ascending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   400
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Output Options"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   10920
      TabIndex        =   8
      Top             =   4320
      Width           =   3285
      Begin VB.OptionButton Option5 
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   400
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   400
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   400
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Print to Screen"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   7200
      TabIndex        =   4
      Top             =   5160
      Width           =   6975
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin Crystal.CrystalReport CrstlRpt 
         Left            =   3480
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame FraAssetsAndLiabilities 
      Caption         =   "Selection Criteria"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   7200
      TabIndex        =   2
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox CboBillingCompany 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Billing Company"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Of Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.ListBox LstReports 
         Height          =   5910
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
End
Attribute VB_Name = "FrmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsRecords As New ADODB.Recordset

Private Sub CboCategory_Click()
On Error GoTo ERRORHANDLER
    Dim lvPrescriptionCategoryID As Long
    'POPULATE COMBO FOR DRUGS BY CATEGORY
    CboProduct.Clear
    lvPrescriptionCategoryID = Mid(CboCategory, 1, 3)
    RsRecords.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboProduct.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboProduct.AddItem !PRODUCTID & " - " & !ProductName
                End If
                .MoveNext
            Wend
        End With
    RsRecords.Close
Exit Sub
ERRORHANDLER:
    MsgBox Err.Description + " Please Contact System Vendor on 0722-729-365 For assistance.", vbExclamation
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdPrint_Click()
On Error GoTo ERRORHANDLER

'SEARCH CRITERIA NOTIFICATIONS
    If LstReports.Text = "DAILY  CASH  TRANSACTIONS - BY CARD NUMBER" And TxtCardNumber = "" Then
        MsgBox "Please Input the card Number to Print on the Selection Criteria", vbExclamation, "Criteria Required": Exit Sub
    End If
'*****************************
        'CLEAR ALL PREVIOUS SELECTION FORMULAS AND FORMULAS
        CrstlRpt.SelectionFormula = ""
        'StrCompanyName = "BANK OF AFRICA"
        'STRReportName = "ASSETS AND LIABILITIES"
        Select Case LstReports.Text
           Case "PATIENT VISITS BY DATE"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT VISIT BY DATE REPORT"
                   '.Connect = Conn.ConnectionString
                    .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Patients.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                        If OptSingleDate.Value = True Then
                            If TxtCardNumber <> "" Then
                                .SelectionFormula = "{COMPLAINS.CARDNUMBER} = '" & TxtCardNumber & "' AND {COMPLAINS.VISITDATE} =date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            Else
                                .SelectionFormula = "{COMPLAINS.VISITDATE} = date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            End If
                        ElseIf OptDateRange.Value = True Then
                            If TxtCardNumber <> "" Then
                                .SelectionFormula = "{COMPLAINS.CARDNUMBER} = '" & TxtCardNumber & "' AND {COMPLAINS.VISITDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            Else
                                .SelectionFormula = "{COMPLAINS.VISITDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {COMPLAINS.VISITDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                            End If
                        ElseIf TxtCardNumber <> "" Then
                            .SelectionFormula = "{COMPLAINS.CARDNUMBER} = '" & TxtCardNumber & "'"
                        Else
                            'PRINT ALL
                        End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "PATIENT BIODATA"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Patient BioData.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .Destination = 0
                   .Action = 1
                End With
            Case "SUMMARY OF VISIT NUMBERS BY DATE"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\PATIENTCOUNT.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   If OptSingleDate.Value = True Then
                        .SelectionFormula = "{COMPLAINS.VISITDATE} = Date(" & Format(DTStartDate, "YYYY,MM,DD") & ")"
                   End If
                   .Destination = 0
                   .Action = 1
                End With
                
            Case "INVOICE PER PATIENT PER VISIT"
                With CrstlRpt
                    Dim RptConn As New ADODB.Connection
                    'RptConn.ConnectionString = "Provider=SQLOLEDB.1;Password=Today123;Persist Security Info=True;User ID=SA;Initial Catalog=OUTPATIENT;Data Source=SYB-KEN-NB-002\SQL2005"
                    RptConn.ConnectionString = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"
                    RptConn.Open
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Invoice.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .SelectionFormula = "{COMPLAINS.VISITDATE} = Date(" & Format(GlbSysDate, "YYYY,MM,DD") & ")"
                   .Destination = 0
                   .Action = 1
                End With
            Case "STATEMENT PER BILLING COMPANY BY DATE RANGE"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\STATEMENTS.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY BILLING COMPANY"
                   '.SelectionFormula = "{COMPLAINS.VISITDATE} = Date(" & Format(DtVisitDate.Value, "YYYY,MM,DD") & ")"
                   .Destination = 0
                   .Action = 1
                End With
            Case "SUMMARY OF DAILY COLLECTION BY DATE"
                If VerifyAccess(GlbCurrentUser, "Cashier") = False Then MsgBox "You do not have Sufficient Privileges to access this Report", vbExclamation: Exit Sub
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Daily Cash Summary.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .Destination = 0
                   .Action = 1
                End With
            Case "DETAILED MEDICINE DISTRIBUTION BY PATIENT"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\MedicineDistribution.rpt"
                   .WindowTitle = StrCompanyName & " - " & "MEDICINE DISTRIBUTION BY PATIENT"
                        If OptSingleDate.Value = True Then
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} = date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                        ElseIf OptDateRange.Value = True Then
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {PRESCRIPTION.PAYDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                        End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "SUMMARY MEDICINE DISTRIBUTION"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\MedicineDistributionSummary.rpt"
                   .WindowTitle = StrCompanyName & " - " & "SUMMARY MEDICINE DISTRIBUTION"
                        If OptSingleDate.Value = True Then
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} = date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                        ElseIf OptDateRange.Value = True Then
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {PRESCRIPTION.PAYDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                        End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "SUMMARY  MEDICINE STOCK LEVELS"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Stock Levels.rpt"
                   .WindowTitle = StrCompanyName & " - " & "STOCK LEVELS"
                        If OptSingleDate.Value = True Then
                                '.SelectionFormula = "{PRESCRIPTION.PAYDATE} = date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                        ElseIf OptDateRange.Value = True Then
                                '.SelectionFormula = "{PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {PRESCRIPTION.PAYDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                        End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "DAILY  CASH  TRANSACTIONS - ALL", "DAILY  CASH  TRANSACTIONS - BY CARD NUMBER"
                If VerifyAccess(GlbCurrentUser, "Cashier") = False Then MsgBox "You do not have Sufficient Privileges to access this Report", vbExclamation: Exit Sub
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "DAILY CASH TRANSACTIONS - ALL"
                   '.Connect = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Daily Cash Collection.rpt"
                   .WindowTitle = StrCompanyName & " - " & " DAILY CASH TRANSACTIONS"
                        If OptSingleDate.Value = True Then
                            If TxtCardNumber <> "" Then
                                .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & TxtCardNumber & "' AND {PRESCRIPTION.PAYDATE} =date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            Else
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} = date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            End If
                        ElseIf OptDateRange.Value = True Then
                            If TxtCardNumber <> "" Then
                                .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & TxtCardNumber & "' AND {PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            Else
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {PRESCRIPTION.PAYDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                            End If
                        ElseIf TxtCardNumber <> "" Then
                            .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & TxtCardNumber & "'"
                        Else
                            'PRINT ALL
                        End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "DAILY CREDIT TRANSACTIONS - ALL", "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER"
                If VerifyAccess(GlbCurrentUser, "Cashier") = False Then MsgBox "You do not have Sufficient Privileges to access this Report", vbExclamation: Exit Sub
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "DAILY CASH TRANSACTIONS - ALL"
                   '.Connect = "DSN=NCC;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Creditors by Date.RPT"
                   .WindowTitle = StrCompanyName & " - " & " DAILY CASH TRANSACTIONS"
                        If OptSingleDate.Value = True Then
                            If TxtCardNumber <> "" Then
                                .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & TxtCardNumber & "' AND {PRESCRIPTION.PAYDATE} =date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            Else
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} =date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            End If
                        ElseIf OptDateRange.Value = True Then
                            If TxtCardNumber <> "" Then
                                .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & TxtCardNumber & "'" 'AND {PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {PRESCRIPTION.PAYDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                            Else
                                'CREDIT ENTRIES DO NOT HAVE PAY DATES THUS SHOULD NOT FILTER BY DATE
                                '.SelectionFormula = "{PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {PRESCRIPTION.PAYDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                            End If
                        ElseIf TxtCardNumber <> "" Then
                            .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & TxtCardNumber & "'"
                        Else
                            'PRINT ALL
                        End If
                   .Destination = 0
                   .Action = 1
                End With
                
            Case "MEDICINE SALES PER - SHIFT"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "MEDICINE SALES GROUPED BY SHIFT"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & "" ';DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\MedicineSalesByShift.RPT"
                   .WindowTitle = StrCompanyName & " - " & " MEDICINE SALES TRANSACTIONS"
                    'If LstReports.Text = "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER" Then
                        If OptSingleDate.Value = True Then
                            .SelectionFormula = "{DRUG_SALES_REPORT.SALEDATE} = DATE('" & Format(DTStartDate, "DD MMM YYYY") & "')"
                        End If
                    'End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "MEDICINE SALES PER - RECEIPT NUMBER"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "MEDICINE SALES GROUPED BY RECEIPT NUMBER"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & "" ';DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\DirectSalesByReceiptNumber.RPT"
                   .WindowTitle = StrCompanyName & " - " & " MEDICINE SALES TRANSACTIONS"
                    'If LstReports.Text = "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER" Then
                        If OptSingleDate.Value = True Then
                            .SelectionFormula = "{DRUGS_SALES.PAYDATE} = DATE('" & Format(DTStartDate, "DD MMM YYYY") & "')"
                        End If
                    'End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "PHARMACY DIRECT SALES - PHARMACY"
                'If VerifyAccess(GlbCurrentUser, "Cashier") = False Then MsgBox "You do not have Sufficient Privileges to access this Report", vbExclamation: Exit Sub
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "DAILY CASH TRANSACTIONS - ALL"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & "" ';DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\PharmacySales.RPT"
                   .WindowTitle = StrCompanyName & " - " & " PHARMACY CASH TRANSACTIONS"
                    'If LstReports.Text = "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER" Then
                        If OptSingleDate.Value = True Then
                            .SelectionFormula = "{DRUGS_SALES.PAYDATE} = DATE('" & Format(DTStartDate, "DD MMM YYYY") & "')"
                        End If
                    'End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "MINISTRY OF HEALTH - WEEKLY REPORT"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "MINISTRY OF HEALTH - WEEKLY REPORT"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & "" ';DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\MOHWeekly.RPT"
                   .WindowTitle = StrCompanyName & " - " & " MINISTRY OF HEALTH - WEEKLY REPORT"
                    'If LstReports.Text = "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER" Then
                        '.SelectionFormula = "{PRESCRIPTION.PAYDATE} = '" & Format(Date, "DD/MM/YYYY") & "'"
                    'End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "MINISTRY OF HEALTH - MONTHLY REPORT"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "MINISTRY OF HEALTH - WEEKLY REPORT"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & "" ';DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\MOHmONTHLY.RPT"
                   .WindowTitle = StrCompanyName & " - " & " MINISTRY OF HEALTH - WEEKLY REPORT"
                    'If LstReports.Text = "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER" Then
                        '.SelectionFormula = "{PRESCRIPTION.PAYDATE} = '" & Format(Date, "DD/MM/YYYY") & "'"
                    'End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "MINISTRY OF HEALTH - TALLY SHEET - DIAGNOSIS UNDER 5 YEARS"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\TallyReportUnder5.rpt"
                   .WindowTitle = StrCompanyName & " - " & " TALLY SHEET - DIAGNOSIS UNDER 5 YEARS"
                   If OptSingleDate.Value = True Then
                        .SelectionFormula = "{DOC_DIAGNOSIS.VISITDATE} = Date(" & Format(DTStartDate, "YYYY,MM,DD") & ")"
                    ElseIf OptDateRange.Value = True Then
                        .SelectionFormula = "{DOC_DIAGNOSIS.VISITDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {DOC_DIAGNOSIS.VISITDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                   End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "MINISTRY OF HEALTH - TALLY SHEET - DIAGNOSIS OVER 5 YEARS"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\TallyReportOVER5.rpt"
                   .WindowTitle = StrCompanyName & " - " & " TALLY SHEET - DIAGNOSIS OVER 5 YEARS"
                   If OptSingleDate.Value = True Then
                        .SelectionFormula = "{DOC_DIAGNOSIS.VISITDATE} = Date(" & Format(DTStartDate, "YYYY,MM,DD") & ")"
                    ElseIf OptDateRange.Value = True Then
                        .SelectionFormula = "{DOC_DIAGNOSIS.VISITDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {DOC_DIAGNOSIS.VISITDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                   End If
                   .Destination = 0
                   .Action = 1
                End With
                Case "LABORATORY TESTS SUMMARY REPORT"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   .Connect = "DSN=NCC;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\LaboratoryTestsReport.rpt"
                   .WindowTitle = StrCompanyName & " - " & " LABORATORY TESTS SUMMARY REPORT"
                   If OptSingleDate.Value = True Then
                        .SelectionFormula = "{PRESCRIPTION.VISITDATE} = Date(" & Format(DTStartDate, "YYYY,MM,DD") & ")"
                    ElseIf OptDateRange.Value = True Then
                        .SelectionFormula = "{PRESCRIPTION.VISITDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "') and {PRESCRIPTION.VISITDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
                   End If
                   .Destination = 0
                   .Action = 1
                End With
        End Select
        AuditTrail GlbCurrentUser, EnumReports, GlbSysDate, Time, "Printed Report - " + "" & LstReports.Text & ""
                MousePointer = vbNormal
Exit Sub
ERRORHANDLER:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    centerform Me
    
    'SET THE DATE PICKERS TO SHOW CURRENT DATE AND NOT THE DATE THEY WERE PLACED THERE.
    For Each CTL In Me.Controls
        If TypeOf CTL Is DTPicker Then
            CTL = GlbSysDate
        End If
    Next

    
    LstReports.AddItem "PATIENT BIODATA"
    LstReports.AddItem "PATIENT VISITS BY DATE"
    LstReports.AddItem "INVOICE PER PATIENT PER VISIT"
    LstReports.AddItem "PATIENT PRESCRIPTION PER VISIT"
    LstReports.AddItem "STATEMENT PER BILLING COMPANY BY DATE RANGE"
    LstReports.AddItem "SUMMARY OF DAILY COLLECTION BY DATE"
    LstReports.AddItem "SUMMARY OF VISIT NUMBERS BY DATE"
    LstReports.AddItem "   "
    LstReports.AddItem "DAILY  CASH  TRANSACTIONS - ALL"
    LstReports.AddItem "DAILY  CASH  TRANSACTIONS - BY CARD NUMBER"
    LstReports.AddItem "DAILY CREDIT TRANSACTIONS - ALL"
    LstReports.AddItem "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER"
    LstReports.AddItem "    "
    LstReports.AddItem "SUMMARY MEDICINE DISTRIBUTION"
    LstReports.AddItem "DETAILED MEDICINE DISTRIBUTION BY PATIENT"
    LstReports.AddItem "SUMMARY  MEDICINE STOCK LEVELS"
    LstReports.AddItem "MEDICINE SALES PER - SHIFT"
    LstReports.AddItem "MEDICINE SALES PER - RECEIPT NUMBER"
    LstReports.AddItem "PHARMACY DIRECT SALES - PHARMACY"
    LstReports.AddItem "    "
    LstReports.AddItem "LABORATORY TESTS SUMMARY REPORT"
    LstReports.AddItem "    "
    LstReports.AddItem "MINISTRY OF HEALTH - WEEKLY REPORT"
    LstReports.AddItem "MINISTRY OF HEALTH - MONTHLY REPORT"
    LstReports.AddItem "MINISTRY OF HEALTH - TALLY SHEET - DIAGNOSIS UNDER 5 YEARS"
    LstReports.AddItem "MINISTRY OF HEALTH - TALLY SHEET - DIAGNOSIS OVER 5 YEARS"
    '*******************************************************************************************************
    DTStartDate.Value = GlbSysDate
    DTEndDate.Value = GlbSysDate
    
        'POPULATE COMBO FOR PRESCRIPTION CATEGORY
    RsRecords.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY ORDER BY PRODUCTGROUPID", Conn, adOpenDynamic, adLockOptimistic
    
        With RsRecords
            While .BOF = False And .EOF = False
                CboCategory.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
                .MoveNext
            Wend
        End With
    RsRecords.Close
    
    'POPULATE COMBO FOR BILLING COMPANIES
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM SERVICE_PROVIDER", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboBillingCompany.AddItem RsRecords!COMPANYCODE & " - " & RsRecords!SERVICEPROVIDER
            RsRecords.MoveNext
        Wend
    RsRecords.Close

End Sub

Private Sub OptDateRange_Click()
    Label4.Visible = True
    DTEndDate.Visible = True
End Sub

Private Sub OptSingleDate_Click()
    Label4.Visible = False
    DTEndDate.Visible = False
End Sub
