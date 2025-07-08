VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmTreatmentNCC 
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17250
   Icon            =   "FrmTreatmentNCC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   17250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame14 
      Height          =   855
      Left            =   15240
      TabIndex        =   78
      Top             =   8760
      Width           =   1935
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
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
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Width           =   17025
      _ExtentX        =   30030
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "History of Presenting illness"
      TabPicture(0)   =   "FrmTreatmentNCC.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "CAD Risk Factors"
      TabPicture(1)   =   "FrmTreatmentNCC.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Previous Medical History"
      TabPicture(2)   =   "FrmTreatmentNCC.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label11"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "SSTab2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text7"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text8"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text9"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Social History"
      TabPicture(3)   =   "FrmTreatmentNCC.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Physical Examination"
      TabPicture(4)   =   "FrmTreatmentNCC.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame17"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame18"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame19"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame20"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Diagnosis"
      TabPicture(5)   =   "FrmTreatmentNCC.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame13"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame16"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame15"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Plan"
      TabPicture(6)   =   "FrmTreatmentNCC.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame21"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame21 
         Caption         =   "Frame21"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   136
         Top             =   360
         Width           =   16695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save Text"
         Height          =   495
         Left            =   14040
         TabIndex        =   135
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Frame Frame20 
         Caption         =   "Abdomen"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -66360
         TabIndex        =   134
         Top             =   3720
         Width           =   8175
      End
      Begin VB.Frame Frame19 
         Caption         =   "Respiratory"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74880
         TabIndex        =   133
         Top             =   3720
         Width           =   8415
      End
      Begin VB.Frame Frame18 
         Caption         =   "Extremities"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2055
         Left            =   -75000
         TabIndex        =   103
         Top             =   1680
         Width           =   16815
         Begin VB.TextBox Text28 
            Height          =   495
            Left            =   8400
            TabIndex        =   132
            Text            =   "Text28"
            Top             =   600
            Width           =   4095
         End
         Begin VB.TextBox Text27 
            Height          =   285
            Left            =   15000
            TabIndex        =   131
            Text            =   "Text27"
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Text24 
            Height          =   285
            Left            =   15000
            TabIndex        =   129
            Text            =   "Text24"
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox Text26 
            Height          =   285
            Left            =   13200
            TabIndex        =   127
            Text            =   "Text26"
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Text25 
            Height          =   285
            Left            =   13200
            TabIndex        =   125
            Text            =   "Text25"
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox Text23 
            Height          =   285
            Left            =   5520
            TabIndex        =   120
            Text            =   "Text16"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text22 
            Height          =   285
            Left            =   5520
            TabIndex        =   119
            Text            =   "Text16"
            Top             =   680
            Width           =   1095
         End
         Begin VB.TextBox Text21 
            Height          =   285
            Left            =   5520
            TabIndex        =   118
            Text            =   "Text16"
            Top             =   1120
            Width           =   1095
         End
         Begin VB.TextBox Text20 
            Height          =   285
            Left            =   5520
            TabIndex        =   117
            Text            =   "Text16"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   2160
            TabIndex        =   112
            Text            =   "Text16"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   2160
            TabIndex        =   111
            Text            =   "Text16"
            Top             =   680
            Width           =   1095
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   2160
            TabIndex        =   110
            Text            =   "Text16"
            Top             =   1120
            Width           =   1095
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   2160
            TabIndex        =   109
            Text            =   "Text16"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label37 
            Caption         =   "MURMURS"
            Height          =   255
            Left            =   15000
            TabIndex        =   130
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label36 
            Caption         =   "S3  S4"
            Height          =   255
            Left            =   15000
            TabIndex        =   128
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label35 
            Caption         =   "S1  S2"
            Height          =   255
            Left            =   13200
            TabIndex        =   126
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label34 
            Caption         =   "APICAL IMPULSE"
            Height          =   255
            Left            =   13200
            TabIndex        =   124
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label33 
            Caption         =   "CARDIAC"
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
            Left            =   14280
            TabIndex        =   123
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "OTHERS"
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
            Left            =   7440
            TabIndex        =   122
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "PULSE RIGHT"
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
            Left            =   3600
            TabIndex        =   121
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "PT"
            Height          =   255
            Left            =   5040
            TabIndex        =   116
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "DP"
            Height          =   255
            Left            =   5040
            TabIndex        =   115
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "PA"
            Height          =   255
            Left            =   5040
            TabIndex        =   114
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "FA"
            Height          =   255
            Left            =   5040
            TabIndex        =   113
            Top             =   280
            Width           =   375
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "PT"
            Height          =   255
            Left            =   1680
            TabIndex        =   108
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "DP"
            Height          =   255
            Left            =   1680
            TabIndex        =   107
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "PA"
            Height          =   255
            Left            =   1680
            TabIndex        =   106
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "FA"
            Height          =   255
            Left            =   1680
            TabIndex        =   105
            Top             =   280
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "PULSE LEFT"
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
            Left            =   480
            TabIndex        =   104
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Higher Functions"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   -74880
         TabIndex        =   92
         Top             =   480
         Width           =   16695
         Begin VB.TextBox Text15 
            Height          =   315
            Left            =   13800
            TabIndex        =   102
            Text            =   "Text15"
            Top             =   360
            Width           =   2655
         End
         Begin VB.CheckBox Check20 
            Caption         =   "THYROMEGALLY"
            Height          =   255
            Left            =   9720
            TabIndex        =   100
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check19 
            Caption         =   "BRUITS"
            Height          =   255
            Left            =   7440
            TabIndex        =   99
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox Check18 
            Caption         =   "JVP"
            Height          =   255
            Left            =   7440
            TabIndex        =   98
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text14 
            Height          =   315
            Left            =   1560
            TabIndex        =   96
            Text            =   "Text14"
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox Text13 
            Height          =   315
            Left            =   1560
            TabIndex        =   94
            Text            =   "Text13"
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label21 
            Caption         =   "ADENOPATHY"
            Height          =   255
            Left            =   12480
            TabIndex        =   101
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "NECK"
            Height          =   255
            Left            =   6480
            TabIndex        =   97
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "OROPHARYNX"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "HEAD"
            Height          =   255
            Left            =   960
            TabIndex        =   93
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Lab Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -66120
         TabIndex        =   89
         Top             =   2760
         Width           =   7935
         Begin VB.CommandButton CmdViewScan 
            Caption         =   "View Full Image"
            Height          =   375
            Left            =   5280
            TabIndex        =   90
            Top             =   1680
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtLabResults 
            Height          =   1935
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   91
            Top             =   240
            Width           =   7695
         End
         Begin VB.Image ImgPreview 
            Height          =   1995
            Left            =   120
            Picture         =   "FrmTreatmentNCC.frx":0506
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   5280
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Lab Request"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -66120
         TabIndex        =   87
         Top             =   480
         Width           =   7935
         Begin VB.TextBox TxtLabRequest 
            Height          =   1935
            HideSelection   =   0   'False
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   88
            ToolTipText     =   "Tests Requested From laboratory"
            Top             =   240
            Width           =   7695
         End
      End
      Begin VB.Frame Frame13 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   80
         Top             =   480
         Width           =   6975
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   2520
            TabIndex        =   84
            Text            =   "Combo2"
            Top             =   360
            Width           =   4335
         End
         Begin VB.ListBox List1 
            Height          =   1425
            Left            =   120
            TabIndex        =   83
            Top             =   840
            Width           =   6735
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   2400
            TabIndex        =   82
            Text            =   "Combo2"
            Top             =   2400
            Width           =   4455
         End
         Begin VB.ListBox List2 
            Height          =   1425
            Left            =   120
            TabIndex        =   81
            Top             =   2880
            Width           =   6735
         End
         Begin VB.Label Label16 
            Caption         =   "ACCUTE"
            Height          =   255
            Left            =   1680
            TabIndex        =   86
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "OTHER"
            Height          =   255
            Left            =   1560
            TabIndex        =   85
            Top             =   2400
            Width           =   735
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Social History"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4575
         Left            =   -74880
         TabIndex        =   59
         Top             =   480
         Width           =   16695
         Begin VB.Frame Frame11 
            Caption         =   "Review of Systems"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3375
            Left            =   120
            TabIndex        =   64
            Top             =   1080
            Width           =   16455
            Begin VB.CheckBox Check17 
               Caption         =   "PREVIOUS COLONOSCOPY"
               Height          =   495
               Left            =   7080
               TabIndex        =   76
               Top             =   840
               Width           =   2535
            End
            Begin VB.CheckBox Check16 
               Caption         =   "SNORRING"
               Height          =   495
               Left            =   7080
               TabIndex        =   75
               Top             =   360
               Width           =   1215
            End
            Begin VB.CheckBox Check15 
               Caption         =   "PARAESTHESIAS"
               Height          =   495
               Left            =   4080
               TabIndex        =   74
               Top             =   1800
               Width           =   1815
            End
            Begin VB.CheckBox Check14 
               Caption         =   "PROSTATISM"
               Height          =   495
               Left            =   4080
               TabIndex        =   73
               Top             =   1320
               Width           =   1575
            End
            Begin VB.CheckBox Check13 
               Caption         =   "DYSPEPSIA"
               Height          =   495
               Left            =   4080
               TabIndex        =   72
               Top             =   840
               Width           =   1215
            End
            Begin VB.CheckBox Check12 
               Caption         =   "WHEEZE"
               Height          =   495
               Left            =   4080
               TabIndex        =   71
               Top             =   360
               Width           =   1215
            End
            Begin VB.CheckBox Check11 
               Caption         =   "CNS:    HEADACHE"
               Height          =   495
               Left            =   360
               TabIndex        =   70
               Top             =   1800
               Width           =   2415
            End
            Begin VB.CheckBox Check10 
               Caption         =   "GIT:    NOCTURIA"
               Height          =   495
               Left            =   360
               TabIndex        =   69
               Top             =   1320
               Width           =   2415
            End
            Begin VB.CheckBox Check9 
               Caption         =   "GIT:    CONSTIPATION"
               Height          =   495
               Left            =   360
               TabIndex        =   68
               Top             =   840
               Width           =   2415
            End
            Begin VB.CheckBox Check8 
               Caption         =   "RESPIRATORY:COUGH"
               Height          =   495
               Left            =   360
               TabIndex        =   67
               Top             =   360
               Width           =   2415
            End
            Begin VB.TextBox Text12 
               Height          =   375
               Left            =   360
               TabIndex        =   66
               Top             =   2760
               Width           =   9375
            End
            Begin VB.Label Label15 
               Caption         =   "MUSCULOSKELETAL:"
               Height          =   255
               Left            =   360
               TabIndex        =   65
               Top             =   2520
               Width           =   1815
            End
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   7800
            TabIndex        =   63
            Text            =   "Text11"
            Top             =   480
            Width           =   4335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2880
            TabIndex        =   61
            Text            =   "Combo1"
            Top             =   480
            Width           =   2175
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   16440
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label14 
            Caption         =   "Occupation"
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
            Left            =   6600
            TabIndex        =   62
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Marital Status"
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
            Left            =   1440
            TabIndex        =   60
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -62040
         TabIndex        =   56
         Top             =   4680
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -65520
         TabIndex        =   54
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -68760
         TabIndex        =   52
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -71040
         TabIndex        =   50
         Top             =   4680
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   41
         Top             =   600
         Width           =   16695
         _ExtentX        =   29448
         _ExtentY        =   6376
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Previous Cardiac Evaluation"
         TabPicture(0)   =   "FrmTreatmentNCC.frx":15890
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame5"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Past Medical History"
         TabPicture(1)   =   "FrmTreatmentNCC.frx":158AC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Past Surgical History"
         TabPicture(2)   =   "FrmTreatmentNCC.frx":158C8
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame7"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            Height          =   3135
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   16455
            Begin VB.TextBox Text5 
               Height          =   2775
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   16215
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frame6"
            Height          =   3135
            Left            =   -74880
            TabIndex        =   44
            Top             =   360
            Width           =   16455
            Begin VB.TextBox Text4 
               Height          =   2775
               Left            =   120
               TabIndex        =   45
               Text            =   "Text4"
               Top             =   240
               Width           =   16215
            End
         End
         Begin VB.Frame Frame5 
            Height          =   3135
            Left            =   -74880
            TabIndex        =   42
            Top             =   360
            Width           =   16455
            Begin VB.TextBox Text3 
               Height          =   2775
               Left            =   120
               TabIndex        =   43
               Text            =   "Text3"
               Top             =   240
               Width           =   16215
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "CAD Risk Factors"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4695
         Left            =   -74880
         TabIndex        =   31
         Top             =   360
         Width           =   16695
         Begin VB.CheckBox Check1 
            Caption         =   "SMOKING"
            Height          =   375
            Left            =   1200
            TabIndex        =   40
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox Check7 
            Caption         =   "FAMILY HISTORY"
            Height          =   375
            Left            =   5280
            TabIndex        =   39
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox Check6 
            Caption         =   "CHRONIC KIDNEY DISEASES"
            Height          =   255
            Left            =   5280
            TabIndex        =   38
            Top             =   900
            Width           =   2535
         End
         Begin VB.CheckBox Check5 
            Caption         =   "ALCOHOL"
            Height          =   375
            Left            =   9240
            TabIndex        =   37
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox Check4 
            Caption         =   "DIABETES MELLITUS/IFG"
            Height          =   255
            Left            =   5280
            TabIndex        =   36
            Top             =   480
            Width           =   2415
         End
         Begin VB.CheckBox Check3 
            Caption         =   "DYSLIPIDAEMIA"
            Height          =   375
            Left            =   1200
            TabIndex        =   35
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox Check2 
            Caption         =   "HYPERTENSION"
            Height          =   375
            Left            =   1200
            TabIndex        =   34
            Top             =   900
            Width           =   1695
         End
         Begin VB.Frame Frame4 
            Caption         =   "Over the Counter  Medication / Suppliments"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   2655
            Left            =   120
            TabIndex        =   32
            Top             =   1920
            Width           =   16455
            Begin VB.TextBox Text2 
               BackColor       =   &H00FFFFFF&
               Height          =   2175
               Left            =   120
               TabIndex        =   33
               Top             =   360
               Width           =   16215
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "History of presenting illness"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4695
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   13815
         Begin VB.TextBox Text1 
            Height          =   4215
            Left            =   120
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   360
            Width           =   13575
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Mammogram"
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
         Left            =   -63360
         TabIndex        =   55
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Pap smear"
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
         Left            =   -66720
         TabIndex        =   53
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69360
         TabIndex        =   51
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Para "
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
         Left            =   -71640
         TabIndex        =   49
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "PAST OBSTETRICS AND GYNAECOLOGICAL HISTORY"
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
         Left            =   -74760
         TabIndex        =   48
         Top             =   4305
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observations"
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   14895
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   11760
         TabIndex        =   58
         Text            =   "Text10"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox TxtSecondName 
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox TxtFirstname 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox TxtBp 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtWeight 
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtHeight 
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtBMI 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DtCurrDate 
         Height          =   285
         Left            =   13440
         TabIndex        =   14
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   53346305
         CurrentDate     =   39415
      End
      Begin VB.Label Label12 
         Caption         =   "Surname"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10920
         TabIndex        =   57
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Second Name"
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
         Left            =   5760
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name"
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
         Left            =   480
         TabIndex        =   25
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Blood Pressure"
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
         TabIndex        =   24
         Top             =   705
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Weight"
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
         Left            =   3120
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Height"
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
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Date"
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
         Left            =   12840
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10320
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label BMI 
         Alignment       =   2  'Center
         Caption         =   "BMI"
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
         Left            =   7560
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Post Patient"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   8760
      Width           =   15015
      Begin VB.OptionButton OptPharmacy 
         Caption         =   "TO PHARMACY"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   8790
         TabIndex        =   11
         Top             =   465
         Width           =   1695
      End
      Begin VB.OptionButton OptCashier 
         Caption         =   "TO CASHIER"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   6660
         TabIndex        =   10
         Top             =   465
         Width           =   1455
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO OBSERVATION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3930
         TabIndex        =   9
         Top             =   465
         Width           =   2055
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO CONSULTATION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   465
         Width           =   2175
      End
      Begin VB.CommandButton CmdPost 
         Caption         =   "POST"
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
         Left            =   12960
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton OptLab 
         Caption         =   "TO LAB"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   11160
         TabIndex        =   6
         Top             =   465
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Patient Previous Visits Ordered by Visit Number"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   14895
      Begin VB.CheckBox ChkClearHistory 
         Caption         =   "Show Current Visit Information"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9720
         TabIndex        =   28
         Top             =   0
         Width           =   3615
      End
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   3201
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
   Begin VB.Frame Frame12 
      Height          =   3495
      Left            =   15120
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "Phase ll Results"
         Height          =   495
         Left            =   120
         TabIndex        =   77
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton CMDSchedule 
         Caption         =   "Schedule Visit"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton CmdMeasurements 
         Caption         =   "View Measurements"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmTreatmentNCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    SSTab1.Tab = 0
    SSTab2.Tab = 0
End Sub

