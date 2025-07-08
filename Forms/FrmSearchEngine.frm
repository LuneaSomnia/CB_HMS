VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FrmSearchEngine 
   Caption         =   "Search Engine"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15735
   Icon            =   "FrmSearchEngine.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "General Search"
      TabPicture(0)   =   "FrmSearchEngine.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Search By ICD10 Classified Condition"
      TabPicture(1)   =   "FrmSearchEngine.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GridICD10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame21"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "LstICD10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "LstSelectedICD10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "CmdICD10Search"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CboRemove"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton CboRemove 
         Caption         =   "Remove From List"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13080
         TabIndex        =   66
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CommandButton CmdICD10Search 
         BackColor       =   &H8000000D&
         Caption         =   "Search by ICD 10"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13080
         MaskColor       =   &H00FF0000&
         TabIndex        =   65
         Top             =   4920
         Width           =   2295
      End
      Begin VB.ListBox LstSelectedICD10 
         Height          =   1425
         Left            =   7800
         TabIndex        =   64
         Top             =   3960
         Width           =   5175
      End
      Begin VB.ListBox LstICD10 
         Height          =   1425
         Left            =   240
         TabIndex        =   63
         Top             =   3960
         Width           =   7455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Date Criteria"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   240
         TabIndex        =   52
         Top             =   360
         Width           =   15135
         Begin VB.CommandButton Command1 
            BackColor       =   &H8000000D&
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   12840
            MaskColor       =   &H00FF0000&
            TabIndex        =   57
            Top             =   480
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Report By Single Date"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   56
            Top             =   360
            Width           =   2895
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Report By Date Range"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   55
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Report All"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10320
            TabIndex        =   54
            Top             =   360
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   7080
            TabIndex        =   58
            Top             =   840
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            _Version        =   393216
            Format          =   139853825
            CurrentDate     =   39163
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   300
            Left            =   3360
            TabIndex        =   59
            Top             =   840
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            Format          =   139853825
            CurrentDate     =   39163
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   8040
            TabIndex        =   53
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "Start Date"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   61
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "End Date"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   60
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "ICD 10 Clasification"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   240
         TabIndex        =   46
         Top             =   1920
         Width           =   15135
         Begin VB.Frame Frame23 
            Height          =   1455
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   14895
            Begin VB.ComboBox CboICD10SubCategory 
               Height          =   315
               Left            =   120
               TabIndex        =   49
               Top             =   1005
               Width           =   14655
            End
            Begin VB.ComboBox CboICD10Category 
               Height          =   315
               Left            =   120
               TabIndex        =   48
               Top             =   405
               Width           =   14655
            End
            Begin VB.Label Label51 
               Caption         =   "Select Disease Sub - Category"
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
               TabIndex        =   51
               Top             =   720
               Width           =   2775
            End
            Begin VB.Label Label49 
               Caption         =   "Select Disease Category"
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
               TabIndex        =   50
               Top             =   120
               Width           =   2775
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search Criteria"
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
         Height          =   8295
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   4815
         Begin VB.CheckBox Check1 
            Caption         =   "Search By Medical History"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   2775
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Search By Current Medication"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   2040
            Width           =   3375
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Search By Family Medical History"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   3840
            Width           =   3735
         End
         Begin VB.ComboBox CboDiagnosisCategory 
            Height          =   315
            Left            =   600
            TabIndex        =   33
            Top             =   960
            Width           =   3975
         End
         Begin VB.ComboBox CboDiagnosis 
            Height          =   315
            Left            =   600
            TabIndex        =   32
            Top             =   1560
            Width           =   3975
         End
         Begin VB.ComboBox CboDrugs 
            Height          =   315
            Left            =   600
            TabIndex        =   31
            Top             =   3240
            Width           =   3975
         End
         Begin VB.ComboBox CboPrescriptionCategory 
            Height          =   315
            Left            =   600
            TabIndex        =   30
            Top             =   2640
            Width           =   3975
         End
         Begin VB.ComboBox CboDiagnosis2 
            Height          =   315
            Left            =   600
            TabIndex        =   29
            Top             =   5040
            Width           =   3975
         End
         Begin VB.ComboBox CboDiagnosisCategory2 
            Height          =   315
            Left            =   600
            TabIndex        =   28
            Top             =   4440
            Width           =   3975
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Search By Suppliment Prescription"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   5520
            Width           =   3735
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Search By Diagnosis"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   6600
            Width           =   3255
         End
         Begin VB.ComboBox CboSuppliment 
            Height          =   315
            Left            =   600
            TabIndex        =   25
            Top             =   6120
            Width           =   3975
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   480
            TabIndex        =   24
            Top             =   7800
            Width           =   4095
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   480
            TabIndex        =   23
            Top             =   7200
            Width           =   4095
         End
         Begin VB.OptionButton OptMedicalHistory 
            Caption         =   "Search By Medical History"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton OptCurrentMedication 
            Caption         =   "Search By Current Medication"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   2040
            Width           =   2655
         End
         Begin VB.OptionButton OptFamilyMedicalHistory 
            Caption         =   "Search By Family Medical History"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   3840
            Width           =   2655
         End
         Begin VB.OptionButton OptSuppliment 
            Caption         =   "Search By Suppliment Prescription"
            Height          =   315
            Left            =   240
            TabIndex        =   19
            Top             =   5520
            Width           =   2775
         End
         Begin VB.OptionButton OptDiagnosis 
            Caption         =   "Search By Diagnosis"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   6600
            Width           =   2055
         End
         Begin VB.Label Label17 
            Caption         =   "Condition Category"
            Height          =   255
            Left            =   600
            TabIndex        =   45
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label16 
            Caption         =   "Condition/Disease"
            Height          =   255
            Left            =   600
            TabIndex        =   44
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Prescription Category"
            Height          =   255
            Left            =   600
            TabIndex        =   43
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Label Label18 
            Caption         =   "Prescription Drug"
            Height          =   255
            Left            =   600
            TabIndex        =   42
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label Label21 
            Caption         =   "Condition/Disease"
            Height          =   255
            Left            =   600
            TabIndex        =   41
            Top             =   4800
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Condition Category"
            Height          =   255
            Left            =   600
            TabIndex        =   40
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Supliment Prescribed"
            Height          =   255
            Left            =   600
            TabIndex        =   39
            Top             =   5880
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Condition/Disease"
            Height          =   255
            Left            =   480
            TabIndex        =   38
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Condition Category"
            Height          =   255
            Left            =   480
            TabIndex        =   37
            Top             =   6960
            Width           =   1815
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   4560
            Y1              =   5400
            Y2              =   5400
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   4560
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   4560
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   4560
            Y1              =   6480
            Y2              =   6480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  Date Criteria"
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
         Height          =   8295
         Left            =   -69840
         TabIndex        =   1
         Top             =   360
         Width           =   10335
         Begin VB.Frame Frame4 
            Height          =   855
            Left            =   0
            TabIndex        =   12
            Top             =   7440
            Width           =   10215
            Begin VB.TextBox TTxtResultsCount 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Left            =   5040
               TabIndex        =   14
               Top             =   300
               Width           =   1935
            End
            Begin VB.CommandButton CmdExit 
               Caption         =   "Exit"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   8040
               TabIndex        =   13
               Top             =   240
               Width           =   2055
            End
            Begin VB.Label Label4 
               Caption         =   "Number of Results"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   15
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1215
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   10095
            Begin VB.TextBox txtDate2 
               Height          =   285
               Left            =   6600
               TabIndex        =   9
               Top             =   720
               Visible         =   0   'False
               Width           =   375
            End
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
               Left            =   6360
               TabIndex        =   6
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
               Left            =   3240
               TabIndex        =   5
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
               TabIndex        =   4
               Top             =   240
               Width           =   2295
            End
            Begin VB.CommandButton CmdSearch 
               BackColor       =   &H8000000D&
               Caption         =   "Search"
               BeginProperty Font 
                  Name            =   "Garamond"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   7800
               MaskColor       =   &H00FF0000&
               TabIndex        =   3
               Top             =   360
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker DTEndDate 
               Height          =   300
               Left            =   5760
               TabIndex        =   7
               Top             =   720
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   529
               _Version        =   393216
               Format          =   173408257
               CurrentDate     =   39163
            End
            Begin MSComCtl2.DTPicker DTStartDate 
               Height          =   300
               Left            =   1560
               TabIndex        =   8
               Top             =   720
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   529
               _Version        =   393216
               Format          =   173408257
               CurrentDate     =   39163
            End
            Begin VB.Label Label5 
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
               Left            =   4440
               TabIndex        =   11
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label6 
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
               TabIndex        =   10
               Top             =   720
               Width           =   1095
            End
         End
         Begin VSFlex6DAOCtl.vsFlexGrid GridResults 
            Height          =   4815
            Left            =   120
            TabIndex        =   16
            Top             =   1560
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   8493
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483624
            ForeColor       =   16711680
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   65535
            ForeColorSel    =   16711680
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483624
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   4
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   3
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0   'False
            ShowComboButton =   -1  'True
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
         End
      End
      Begin VSFlex6DAOCtl.vsFlexGrid GridICD10 
         Height          =   3135
         Left            =   240
         TabIndex        =   62
         Top             =   5520
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   5530
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483624
         ForeColor       =   16711680
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   65535
         ForeColorSel    =   16711680
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483624
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   4
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
   End
End
Attribute VB_Name = "FrmSearchEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCombo As New ADODB.Recordset
Dim RsFilter As New ADODB.Recordset
Dim RsSearch As New ADODB.Recordset
Dim RsDrill As New ADODB.Recordset
Dim RsLoop As New ADODB.Recordset
Public Enum EnumSearch
    EnumMedicalHistory = 1
    EnumCurrentMedication = 2
    EnumFamilyHistory = 3
    EnumSuppliments = 4
    EnumDiagnosis = 5
End Enum
Dim SearchCriteria As Integer

Private Sub CboDiagnosisCategory_Click()
On Error GoTo Errorhandler
    Dim lvDiagnosisCategoryID As Long
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    CboDiagnosis.Clear
    lvDiagnosisCategoryID = Mid(CboDiagnosisCategory, 1, 3)
    RsFilter.Open "SELECT DIAGNOSISID, DIAGNOSISDESCRIPTION FROM DIAGNOSIS WHERE DIAGNOSISCATEGORY = '" & lvDiagnosisCategoryID & "'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsFilter
            While .BOF = False And .EOF = False
                CboDiagnosis.AddItem !DIAGNOSISID & " - " & !DIAGNOSISDESCRIPTION
                .MoveNext
            Wend
        End With
    RsFilter.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboDiagnosisCategory2_Click()
On Error GoTo Errorhandler
    Dim lvDiagnosisCategoryID As Long
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    CboDiagnosis.Clear
    lvDiagnosisCategoryID = Mid(CboDiagnosisCategory2, 1, 3)
    RsFilter.Open "SELECT DIAGNOSISID, DIAGNOSISDESCRIPTION FROM DIAGNOSIS WHERE DIAGNOSISCATEGORY = '" & lvDiagnosisCategoryID & "'", Conn, adOpenDynamic, adLockOptimistic
    
        With RsFilter
            While .BOF = False And .EOF = False
                CboDiagnosis2.AddItem !DIAGNOSISID & " - " & !DIAGNOSISDESCRIPTION
                .MoveNext
            Wend
        End With
    RsFilter.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboICD10Category_Click()
'POPULATE COMBO FOR DISEASE SUB-CATEGORY
    CboICD10SubCategory.Clear
    'LstICD10.Clear
    RsCombo.Open "SELECT DISEASESUBCATEGORY, SUBCATEGORYDESCRIPTION FROM ICD10_SUBCATEGORY WHERE DISEASECATEGORY = '" & GetID_NameFromCombo(CboICD10Category, 1) & "' ORDER BY DISEASESUBCATEGORY", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboICD10SubCategory.AddItem String(3 - Len(!DiseaseSubCategory), "0") & !DiseaseSubCategory & " - " & !SUBCATEGORYDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close
End Sub

Private Sub PopulateICD10_CODES(ByRef DiseaseCategory, ByRef DiseaseSubCategory)
    'POPULATE LIST VIEW FOR MAIN CATEGORY
    LstICD10.Clear
    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
    If RsCombo.State = 1 Then Set RsCombo = Nothing
    RsCombo.Open "SELECT ICD10CHARACTER,ICD10NUMBER,DISEASEDESCRIPTION FROM ICD10_CODES WHERE DISEASECATEGORY = '" & DiseaseCategory & "' AND DISEASESUBCATEGORY = '" & GetID_NameFromCombo(DiseaseSubCategory, 1) & "' ORDER BY ICD10NUMBER", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                LstICD10.AddItem UCase(!ICD10CHARACTER) & String(2 - Len(!ICD10NUMBER), "0") & !ICD10NUMBER & " - " & !DISEASEDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close
End Sub

Private Sub CboICD10SubCategory_Click()
    PopulateICD10_CODES GetID_NameFromCombo(CboICD10Category, 1), CboICD10SubCategory.Text
End Sub

Private Sub CboPrescriptionCategory_Click()
On Error GoTo Errorhandler
    Dim lvPrescriptionCategoryID
    'POPULATE COMBO FOR PRESCRIPTION
    CboDrugs.Clear
    lvPrescriptionCategoryID = Mid(CboPrescriptionCategory, 1, 3)
    RsFilter.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' order by productname", Conn, adOpenDynamic, adLockOptimistic
    
        With RsFilter
            While .BOF = False And .EOF = False
                If Len(!PRODUCTID) = 3 Then
                    CboDrugs.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
                Else
                    CboDrugs.AddItem !PRODUCTID & " - " & !ProductName
                End If
                .MoveNext
            Wend
        End With
    RsFilter.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub

Private Sub CboRemove_Click()
    If LstSelectedICD10.ListIndex < 0 Then Exit Sub
    LstSelectedICD10.RemoveItem LstSelectedICD10.ListIndex
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdICD10Search_Click()

    GridICD10.Rows = 1: GridICD10.Cols = 5
    For i = 0 To LstSelectedICD10.ListCount - 1
        LstSelectedICD10.Selected(i) = True
        GridICD10.FormatString = "CARD NUMBER | VISIT NUMBER | FULL NAMES | ICD10 CODE | ICD10 DESCRIPTION "
        GridICD10.ColWidth(4) = 5000
        GridICD10.ColWidth(2) = 5000
        
        If RsSearch.State = 1 Then Set RsSearch = Nothing
        
        'SELECT ALL CARD NUMBERS THAT MEET THE CRITERIA FIRST.
        RsSearch.Open "Select DISTINCT CARDNUMBER from DOCTOR_DIAGNOSIS WHERE DIAGNOSISTYPE = '3' AND DIAGNOSISDESCRIPTION = '" & GetID_NameFromCombo(LstSelectedICD10, 1) & "'", Conn, adOpenStatic, adLockOptimistic
                While RsSearch.EOF = False
                    If RsLoop.State = 1 Then Set RsLoop = Nothing
                    RsLoop.Open "SELECT TOP 1 * FROM DOCTOR_DIAGNOSIS WHERE CARDNUMBER = '" & RsSearch!CardNumber & "' AND DIAGNOSISTYPE = '3' AND DIAGNOSISDESCRIPTION = '" & GetID_NameFromCombo(LstSelectedICD10, 1) & "' ORDER BY VISITNUMBER DESC", Conn, adOpenStatic, adLockOptimistic
                    With RsLoop
                        If .EOF = False Then
                            'INSERT INTO GRID
                            'ENRICH DATA FROM OTHER TABLES
                            lvFirstName = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & !CardNumber & "'")
                            lvSecondName = FindRecord("PATIENT_DETAILS", "SECONDNAME", "CARDNUMBER = '" & !CardNumber & "'")
                            lvSurname = FindRecord("PATIENT_DETAILS", "SURNAME", "CARDNUMBER = '" & !CardNumber & "'")
                            lvFullNames = lvFirstName + " " + lvSecondName + " " + lvSurname
                            LVDIAGNOSIS = FindRecord("ICD10_CODES", "DISEASEDESCRIPTION", "ICD10CHARACTER = '" & Mid(!DIAGNOSISDESCRIPTION, 1, 1) & "' And ICD10NUMBER = '" & Mid(!DIAGNOSISDESCRIPTION, 2, 2) & "'")
                            GridICD10.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & lvFullNames & vbTab & !DIAGNOSISDESCRIPTION & vbTab & LVDIAGNOSIS
                        End If
                    End With
                    RsSearch.MoveNext
                Wend
    Next i
    If RsSearch.State = 1 Then Set RsSearch = Nothing
End Sub

Private Sub CmdSearch_Click()
    Select Case SearchCriteria
        Case 1
            SearchMedicalHistory
        Case 2
            SearchCurrentMedication
        Case 3
            SearchFamilyMedicalHistory
        Case 4
            SearchSuppliment
    End Select
End Sub
Private Sub SearchMedicalHistory()
    Dim lvCount As Double
    GridResults.Clear: GridResults.Rows = 1: GridResults.Cols = 3
    GridResults.FormatString = " CARD NUMBER | VISIT NUMBER |   CLIENT NAME             | DIAGNOSIS DESCRIPTION         |"
    'MEDICAL HISTORY
    RsSearch.Open "SELECT CARDNUMBER,VISITNUMBER,VISITDATE FROM COMPLAINS WHERE VISITDATE = '" & Format(DTStartDate, "DDMMMYYYY") & "'", Conn, adOpenStatic, adLockOptimistic
        With RsSearch
            While .EOF = False
                'SEARCH IN MEDICAL HISTORY
                RsDrill.Open "SELECT * FROM MEDICAL_HISTORY WHERE CARDNUMBER = '" & !CardNumber & "' AND VISITNUMBER = '" & !VISITNUMBER & "' AND DIAGNOSISID = '" & GetID_NameFromCombo(CboDiagnosis, 1) & "'", Conn, adOpenStatic, adLockOptimistic
                    While RsDrill.EOF = False
                        lvFullNames = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & RsDrill!CardNumber & "'")
                        lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SECONDNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
                        lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SURNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
                        GridResults.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & lvFullNames & vbTab & CboDiagnosis
                        lvCount = lvCount + 1
                        RsDrill.MoveNext
                    Wend
                .MoveNext
                RsDrill.Close
            Wend
        End With
    RsSearch.Close
    TTxtResultsCount = lvCount
End Sub
Private Sub SearchCurrentMedication()
    Dim lvCount As Double
    GridResults.Clear: GridResults.Rows = 1: GridResults.Cols = 3
    GridResults.FormatString = " CARD NUMBER | VISIT NUMBER |   CLIENT NAME     | PRESCRIPTION DESCRIPTION         |"
    'MEDICAL HISTORY
    RsSearch.Open "SELECT CARDNUMBER,VISITNUMBER,VISITDATE FROM COMPLAINS WHERE VISITDATE = '" & Format(DTStartDate, "DDMMMYYYY") & "'", Conn, adOpenStatic, adLockOptimistic
        With RsSearch
            While .EOF = False
                'SEARCH IN MEDICAL HISTORY
                RsDrill.Open "SELECT * FROM CURRENT_MEDICATION WHERE CARDNUMBER = '" & !CardNumber & "' AND VISITNUMBER = '" & !VISITNUMBER & "' AND PRESCRIPTIONID = '" & GetID_NameFromCombo(CboDrugs, 1) & "'", Conn, adOpenStatic, adLockOptimistic
                    While RsDrill.EOF = False
                        lvFullNames = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & RsDrill!CardNumber & "'")
                        lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SECONDNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
                        lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SURNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
                        GridResults.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & lvFullNames & vbTab & CboDrugs
                        lvCount = lvCount + 1
                        RsDrill.MoveNext
                    Wend
                .MoveNext
                RsDrill.Close
            Wend
        End With
    RsSearch.Close
    TTxtResultsCount = lvCount
End Sub
Private Sub SearchFamilyMedicalHistory()
    Dim lvCount As Double
    GridResults.Clear: GridResults.Rows = 1: GridResults.Cols = 3
    GridResults.FormatString = " CARD NUMBER | VISIT NUMBER |   CLIENT NAME     | PRESCRIPTION DESCRIPTION         |"
    'MEDICAL HISTORY
    RsSearch.Open "SELECT CARDNUMBER,VISITNUMBER,VISITDATE FROM COMPLAINS WHERE VISITDATE = '" & Format(DTStartDate, "DDMMMYYYY") & "'", Conn, adOpenStatic, adLockOptimistic
        With RsSearch
            While .EOF = False
                'SEARCH IN MEDICAL HISTORY
                RsDrill.Open "SELECT * FROM FAMILY_MEDICAL_HISTORY WHERE CARDNUMBER = '" & !CardNumber & "' AND VISITNUMBER = '" & !VISITNUMBER & "' AND DIAGNOSISID = '" & GetID_NameFromCombo(CboDiagnosis2, 1) & "'", Conn, adOpenStatic, adLockOptimistic
                    While RsDrill.EOF = False
                        lvFullNames = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & RsDrill!CardNumber & "'")
                        lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SECONDNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
                        lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SURNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
                        GridResults.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & lvFullNames & vbTab & CboDiagnosis2
                        lvCount = lvCount + 1
                        RsDrill.MoveNext
                    Wend
                .MoveNext
                RsDrill.Close
            Wend
        End With
    RsSearch.Close
    TTxtResultsCount = lvCount
End Sub
Private Sub SearchSuppliment()
    Dim lvCount As Double
    GridResults.Clear: GridResults.Rows = 1: GridResults.Cols = 3
    GridResults.FormatString = " CARD NUMBER | VISIT NUMBER |   CLIENT NAME     | PRESCRIPTION DESCRIPTION         |"
    'MEDICAL HISTORY
    RsSearch.Open "SELECT CARDNUMBER,VISITNUMBER,VISITDATE FROM COMPLAINS WHERE VISITDATE = '" & Format(DTStartDate, "DDMMMYYYY") & "'", Conn, adOpenStatic, adLockOptimistic
        With RsSearch
            While .EOF = False
                'SEARCH IN MEDICAL HISTORY
                RsDrill.Open "SELECT * FROM PRESCRIPTION WHERE CARDNUMBER = '" & !CardNumber & "' AND VISITNUMBER = '" & !VISITNUMBER & "' AND CODE = '" & GetID_NameFromCombo(CboSuppliment, 1) & "'", Conn, adOpenStatic, adLockOptimistic
                    While RsDrill.EOF = False
                        lvFullNames = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & RsDrill!CardNumber & "'")
                        lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SECONDNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
                        lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SURNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
                        GridResults.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & lvFullNames & vbTab & CboSuppliment
                        lvCount = lvCount + 1
                        RsDrill.MoveNext
                    Wend
                .MoveNext
                RsDrill.Close
            Wend
        End With
    RsSearch.Close
    TTxtResultsCount = lvCount
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    centerform Me
    
'''    'POPULATE COMBO FOR DIAGNOSIS CATEGORY
'''    RsCombo.Open "SELECT DIAGNOSISCATEGORYID, DIAGNOSISCATEGORYDESC FROM DIAGNOSISCATEGORY ORDER BY DIAGNOSISCATEGORYDESC", Conn, adOpenDynamic, adLockOptimistic
'''
'''        With RsCombo
'''            While .BOF = False And .EOF = False
'''                CboDiagnosisCategory.AddItem String(3 - Len(!DIAGNOSISCATEGORYID), "0") & !DIAGNOSISCATEGORYID & " - " & !DIAGNOSISCATEGORYDESC
'''                CboDiagnosisCategory2.AddItem String(3 - Len(!DIAGNOSISCATEGORYID), "0") & !DIAGNOSISCATEGORYID & " - " & !DIAGNOSISCATEGORYDESC
'''                .MoveNext
'''            Wend
'''        End With
'''    RsCombo.Close
'''
'''    'POPULATE COMBO FOR PRESCRIPTION CATEGORY
'''    RsCombo.Open "SELECT PRODUCTGROUPID, PRODUCTGROUP FROM PRODUCTCATEGORY ORDER BY PRODUCTGROUP", Conn, adOpenDynamic, adLockOptimistic
'''
'''        With RsCombo
'''            While .BOF = False And .EOF = False
'''                CboPrescriptionCategory.AddItem String(3 - Len(!PRODUCTGROUPID), "0") & !PRODUCTGROUPID & " - " & !PRODUCTGROUP
'''                .MoveNext
'''            Wend
'''        End With
'''    RsCombo.Close
'''
'''
'''    'POPULATE COMBO FOR SUPPLIMENTS
'''    lvPrescriptionCategoryID = 14
'''    RsCombo.Open "SELECT PRODUCTID, PRODUCTNAME FROM PRODUCTS WHERE CATEGORYID = ' " & lvPrescriptionCategoryID & "' order by productname", Conn, adOpenDynamic, adLockOptimistic
'''
'''        With RsCombo
'''            While .BOF = False And .EOF = False
'''                If Len(!PRODUCTID) = 3 Then
'''                    CboSuppliment.AddItem String(3 - Len(!PRODUCTID), "0") & !PRODUCTID & " - " & !ProductName
'''                Else
'''                    CboSuppliment.AddItem !PRODUCTID & " - " & !ProductName
'''                End If
'''                .MoveNext
'''            Wend
'''        End With
'''    RsCombo.Close


    'SET THE DATE PICKERS TO SHOW CURRENT DATE AND NOT THE DATE THEY WERE PLACED THERE.
    For Each CTL In Me.Controls
        If TypeOf CTL Is DTPicker Then
            CTL = GlbSysDate
        End If
    Next


    'POPULATE COMBO FOR DISEASE CATEGORY
    RsCombo.Open "SELECT CATEGORYID, CATEGORYDESCRIPTION FROM ICD10_CATEGORY ORDER BY CATEGORYID ASC", Conn, adOpenDynamic, adLockOptimistic
    
        With RsCombo
            While .BOF = False And .EOF = False
                CboICD10Category.AddItem String(3 - Len(!CATEGORYID), "0") & !CATEGORYID & " - " & !CATEGORYDESCRIPTION
                .MoveNext
            Wend
        End With
    RsCombo.Close



End Sub

Private Sub GridICD10_DblClick()
    On Error GoTo Errorhandler
    If GridICD10.Row = 0 Then Me.Caption = "There are no Patients on the Search result List.": Exit Sub
    StrDocCardNo = GridICD10.TextMatrix(GridICD10.Row, 0)
    StrDocVisitNumber = GridICD10.TextMatrix(GridICD10.Row, 1)
    StrDocVisitDate = FindRecord("COMPLAINS", "VISITDATE", "CARDNUMBER = '" & StrDocCardNo & "' AND VISITNUMBER = '" & StrDocVisitNumber & "'")

    FrmTreatment.Show
    Exit Sub
Errorhandler:
    MsgBox Err.Description
    
End Sub

Private Sub GridResults_DblClick()
    'ASSIGN VALUES TO GLOBAL VARIABLES.
    If GridResults.Row = 0 Then Exit Sub
    StrDocCardNo = GridResults.TextMatrix(GridResults.Row, 0)
    StrDocVisitNumber = GridResults.TextMatrix(GridResults.Row, 1)
    'StrDocVisitDate = Grid.TextMatrix(Grid.Row, 5)
    BlnHISTORY = True
    
    lvFullNames = FindRecord("PATIENT_DETAILS", "FIRSTNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
    lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SECONDNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
    lvFullNames = lvFullNames + " " + FindRecord("PATIENT_DETAILS", "SURNAME", "CARDNUMBER = '" & StrDocCardNo & "'")
    AuditTrail GlbCurrentUser, EnumPatientHistory, GlbSysDate, Time, "Loaded Patient History For CardNumber - " + "" & StrDocCardNo & "" + " - " + "" & lvFullNames & ""
    FrmTreatment.Show
End Sub

Private Sub LstICD10_Click()
    LstSelectedICD10.AddItem LstICD10  'GetID_NameFromCombo(LstLabTests, 2)
End Sub

Private Sub OptCurrentMedication_Click()
    DisableControls
    If OptCurrentMedication.Value = True Then
        CboPrescriptionCategory.Enabled = True
        CboDrugs.Enabled = True
    End If
    SearchCriteria = EnumCurrentMedication
End Sub

Private Sub OptDateRange_Click()
    Label5.Visible = True
    DTEndDate.Visible = True
    DTStartDate.Enabled = True
End Sub

Private Sub OptFamilyMedicalHistory_Click()
    If OptFamilyMedicalHistory.Value = True Then
    DisableControls
    ClearControls
    CboDiagnosis2.Enabled = True
    CboDiagnosisCategory2.Enabled = True
    End If
    SearchCriteria = EnumFamilyHistory
End Sub

Private Sub OptMedicalHistory_Click()
    DisableControls
    If OptMedicalHistory.Value = True Then
        CboDiagnosisCategory.Enabled = True
        CboDiagnosis.Enabled = True
        
    End If
    SearchCriteria = EnumMedicalHistory
End Sub
Private Sub DisableControls()
    Dim CTL As Control
    For Each CTL In Me.Controls
        If TypeOf CTL Is TextBox Then
            CTL.Enabled = False
        ElseIf TypeOf CTL Is ComboBox Then
            CTL.Enabled = False
        ElseIf TypeOf CTL Is DTPicker Then
            CTL.Enabled = False
        End If
    Next
End Sub
Private Sub ClearControls()
    Dim CTL As Control
    For Each CTL In Me.Controls
        If TypeOf CTL Is TextBox Then
            CTL.Text = ""
        ElseIf TypeOf CTL Is ComboBox Then
            CTL.Text = ""
        End If
    Next
End Sub

Private Sub OptSingleDate_Click()
    Label5.Visible = False
    DTEndDate.Visible = False
    DTStartDate.Enabled = True
End Sub

Private Sub OptSuppliment_Click()
    If OptSuppliment.Value = True Then
    DisableControls
    ClearControls
    CboSuppliment.Enabled = True
    End If
    SearchCriteria = EnumSuppliments
End Sub
