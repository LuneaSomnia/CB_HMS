VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmFinancialReports 
   Caption         =   "Financial Reports"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   Icon            =   "FrmFinancialReports.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   7215
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   6975
      Begin VB.TextBox TxtCardNumber 
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1695
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
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   6975
      Begin VB.OptionButton OptDateRange 
         Caption         =   "Report By Date Range"
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
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton OptSingleDate 
         Caption         =   "Report By Single Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   22
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTEndDate 
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   720
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         Format          =   54919169
         CurrentDate     =   39163
      End
      Begin MSComCtl2.DTPicker DTStartDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   25
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   54919169
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
         Left            =   3840
         TabIndex        =   13
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
         Left            =   480
         TabIndex        =   12
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
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   3645
      Begin VB.OptionButton Option2 
         Caption         =   "Descending"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   400
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ascending"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   400
         Width           =   1095
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
      Left            =   3840
      TabIndex        =   8
      Top             =   6000
      Width           =   3285
      Begin VB.OptionButton Option5 
         Caption         =   "Printer"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   400
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "File"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   400
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Screen"
         Height          =   255
         Left            =   120
         TabIndex        =   19
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
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   6840
      Width           =   6975
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin Crystal.CrystalReport CrstlRpt 
         Left            =   3240
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
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   6975
      Begin VB.ComboBox Combo1 
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.ListBox LstReports 
         Height          =   2205
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6735
      End
   End
End
Attribute VB_Name = "FrmFinancialReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdPrint_Click()
On Error GoTo ErrorHandler

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
                   .Connect = Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\Patients.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   '.SelectionFormula = "{SUMMARY_A_and_L.OPERATING_DATE} >= Date(" & Format(DTStartDate.value, "YYYY,MM,DD") & ") AND {SUMMARY_A_and_L.OPERATING_DATE} <= Date(" & Format(DTEndDate.value, "YYYY,MM,DD") & ")"
                   '.SelectionFormula = """Operating_Date"""" BETWEEN {ts '2007-03-02 00:00:00.000'} AND {ts '2007-03-05 00:00:00.000'}"
                   '.Formulas(0) = ("Client='" & StrCompanyName & "'")
                   '.Formulas(1) = ("ReportName ='" & STRReportName & "'")
                   .Destination = 0
                   .Action = 1
                End With
            Case "PATIENT BIODATA"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Patient BioData.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .Destination = 0
                   .Action = 1
                End With
            Case "SUMMARY OF VISIT NUMBERS BY DATE"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\PATIENTCOUNT.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .Destination = 0
                   .Action = 1
                End With
                
            Case "INVOICE PER PATIENT PER VISIT"
                With CrstlRpt
                    Dim RptConn As New ADODB.Connection
                    'RptConn.ConnectionString = "Provider=SQLOLEDB.1;Password=Today123;Persist Security Info=True;User ID=SA;Initial Catalog=OUTPATIENT;Data Source=SYB-KEN-NB-002\SQL2005"
                    RptConn.ConnectionString = "DSN=OUTPATIENTS;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"
                    RptConn.Open
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Invoice.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .SelectionFormula = "{COMPLAINS.VISITDATE} = Date(" & Format(Date, "YYYY,MM,DD") & ")"
                   .Destination = 0
                   .Action = 1
                End With
            Case "STATEMENT PER BILLING COMPANY BY DATE RANGE"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\STATEMENTS.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY BILLING COMPANY"
                   '.SelectionFormula = "{COMPLAINS.VISITDATE} = Date(" & Format(DtVisitDate.Value, "YYYY,MM,DD") & ")"
                   .Destination = 0
                   .Action = 1
                End With
            Case "SUMMARY OF DAILY COLLECTION BY DATE"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "PATIENT BIODATA"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Daily Cash Summary.rpt"
                   .WindowTitle = StrCompanyName & " - " & " PATIENT VISITS BY DATE IN ASCENDING ORDER"
                   .Destination = 0
                   .Action = 1
                End With
                
            Case "DAILY  CASH  TRANSACTIONS - ALL", "DAILY  CASH  TRANSACTIONS - BY CARD NUMBER"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "DAILY CASH TRANSACTIONS - ALL"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
                   .ReportFileName = App.Path & "\REPORTS\Daily Cash Collection.RPT"
                   .WindowTitle = StrCompanyName & " - " & " DAILY CASH TRANSACTIONS"
                        If OptSingleDate.Value = True Then
                            If TxtCardNumber <> "" Then
                                .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & TxtCardNumber & "' AND {PRESCRIPTION.PAYDATE} =date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            Else
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} =date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            End If
                        ElseIf OptDateRange.Value = True Then
                            If TxtCardNumber <> "" Then
                                .SelectionFormula = "{PRESCRIPTION.CARDNUMBER} = '" & TxtCardNumber & "' AND {PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "')"
                            Else
                                .SelectionFormula = "{PRESCRIPTION.PAYDATE} >= date('" & Format(DTStartDate, "dd mmmm yyyy") & "')" 'and {PRESCRIPTION.PAYDATE} <= date('" & Format(DTEndDate, "dd mmmm yyyy") & "')"
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
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "DAILY CASH TRANSACTIONS - ALL"
                   '.Connect = "DSN=OUTPATIENTS;UID=sa;PWD=CES123;DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & ""
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
            Case "PHARMACY  DIRECT SALES - PHARMACY"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "DAILY CASH TRANSACTIONS - ALL"
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & "" ';DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\PharmacySales.RPT"
                   .WindowTitle = StrCompanyName & " - " & " PHARMACY CASH TRANSACTIONS"
                    'If LstReports.Text = "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER" Then
                        '.SelectionFormula = "{PRESCRIPTION.PAYDATE} = '" & Format(Date, "DD/MM/YYYY") & "'"
                    'End If
                   .Destination = 0
                   .Action = 1
                End With
            Case "MINISTRY OF HEALTH - WEEKLY REPORT"
                With CrstlRpt
                   .SelectionFormula = ""
                    STRReportName = "MINISTRY OF HEALTH - WEEKLY REPORT"
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & "" ';DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
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
                   .Connect = "DSN=OUTPATIENTS;UID=" & DBUser & ";PWD=" & DBPassword & "" ';DSQ=SYB-KEN-NB-002\SQL2005;"   'Conn.ConnectionString
                   .ReportFileName = App.Path & "\REPORTS\MOHmONTHLY.RPT"
                   .WindowTitle = StrCompanyName & " - " & " MINISTRY OF HEALTH - WEEKLY REPORT"
                    'If LstReports.Text = "DAILY CREDIT TRANSACTIONS - BY CARD NUMBER" Then
                        '.SelectionFormula = "{PRESCRIPTION.PAYDATE} = '" & Format(Date, "DD/MM/YYYY") & "'"
                    'End If
                   .Destination = 0
                   .Action = 1
                End With
        End Select
                MousePointer = vbNormal
Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub Form_Load()
    centerform Me
    
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
    LstReports.AddItem "PHARMACY  DIRECT SALES - PHARMACY"
    LstReports.AddItem "    "
    LstReports.AddItem "MINISTRY OF HEALTH - WEEKLY REPORT"
    LstReports.AddItem "MINISTRY OF HEALTH - MONTHLY REPORT"
End Sub

Private Sub OptDateRange_Click()
    Label4.Visible = True
    DTEndDate.Visible = True
End Sub

Private Sub OptSingleDate_Click()
    Label4.Visible = False
    DTEndDate.Visible = False
End Sub
