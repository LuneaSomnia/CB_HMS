VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FrmObservation 
   Caption         =   "Nurse Observations"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   Icon            =   "FrmObservation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtCount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   11760
      TabIndex        =   92
      Text            =   "0"
      Top             =   3240
      Width           =   1215
   End
   Begin TabDlg.SSTab TabObservation 
      Height          =   4815
      Left            =   120
      TabIndex        =   39
      Top             =   3600
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   8493
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Previous Medical History"
      TabPicture(0)   =   "FrmObservation.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Vital Signs"
      TabPicture(1)   =   "FrmObservation.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameVitalSigns"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrameClinicReview"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "OptComprehensive"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "OptClinicReview"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Chief Complaint"
      TabPicture(2)   =   "FrmObservation.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.OptionButton OptClinicReview 
         Caption         =   "Clinic Review"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   300
         Left            =   -67680
         TabIndex        =   94
         Top             =   450
         Width           =   2535
      End
      Begin VB.OptionButton OptComprehensive 
         Caption         =   "Comprehensive Capture"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   300
         Left            =   -71400
         TabIndex        =   93
         Top             =   450
         Width           =   3495
      End
      Begin VB.Frame Frame5 
         Caption         =   "Allergies"
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
         Height          =   1455
         Left            =   120
         TabIndex        =   52
         Top             =   3240
         Width           =   12615
         Begin VB.TextBox TxtOtherAllergies 
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
            Left            =   3720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            ToolTipText     =   "OTHERS"
            Top             =   720
            Width           =   8655
         End
         Begin VB.CheckBox ChkOthersMedicalHistory 
            Caption         =   "OTHERS"
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
            Left            =   3720
            TabIndex        =   49
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox ChkSulphurDrugs 
            Caption         =   "SULPHUR DRUGS"
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
            Left            =   1560
            TabIndex        =   48
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox ChkPeniciline 
            Caption         =   "PENICILINE"
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
            Left            =   1560
            TabIndex        =   47
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Previous Immunization"
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
         Height          =   1335
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   12615
         Begin VB.CheckBox ChkHepatitisBvaccine 
            Caption         =   "HEPATITIS B VACCINE"
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
            Left            =   1560
            TabIndex        =   46
            Top             =   840
            Width           =   2775
         End
         Begin VB.CheckBox ChkPheumoniaVaccine 
            Caption         =   "PHEUMONIA VACCINE"
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
            Left            =   1560
            TabIndex        =   45
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Medical History"
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
         Height          =   1575
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   12615
         Begin VB.CheckBox ChkCancer 
            Caption         =   "CANCER"
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
            Left            =   5760
            TabIndex        =   44
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox ChkAsthma 
            Caption         =   "ASTHMA"
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
            Left            =   5760
            TabIndex        =   43
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox ChkGout 
            Caption         =   "GOUT"
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
            Left            =   5760
            TabIndex        =   42
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox ChkHighCholestrol 
            Caption         =   "HIGH CHOLESTROL"
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
            Left            =   1560
            TabIndex        =   41
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox ChkHypertension 
            Caption         =   "HYPERTENSION"
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
            Left            =   1560
            TabIndex        =   23
            Top             =   660
            Width           =   1575
         End
         Begin VB.CheckBox ChkDiabetes 
            Caption         =   "DIABETES"
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
            Left            =   1560
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Chief Complaint"
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
         Height          =   4335
         Left            =   -74880
         TabIndex        =   78
         Top             =   360
         Width           =   12615
         Begin VB.TextBox TxtChiefComplaint 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            Top             =   360
            Width           =   12375
         End
      End
      Begin VB.Frame FrameClinicReview 
         Caption         =   "Clinic Review"
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
         Height          =   3975
         Left            =   -74880
         TabIndex        =   95
         Top             =   720
         Width           =   12615
         Begin VB.TextBox TxtNurse 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   10560
            TabIndex        =   118
            Text            =   "Sarah"
            Top             =   3480
            Width           =   1935
         End
         Begin VB.CheckBox ChkHighlightSugarInRed 
            Alignment       =   1  'Right Justify
            Caption         =   "Highlight In Red"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   480
            TabIndex        =   116
            Top             =   2730
            Width           =   1815
         End
         Begin VB.CheckBox ChkHighlighBPtInRed 
            Alignment       =   1  'Right Justify
            Caption         =   "Highlight In Red"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   600
            TabIndex        =   110
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox TxtPlan 
            Height          =   1695
            Left            =   7080
            TabIndex        =   108
            Top             =   600
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.TextBox TxtSPO2Review 
            Height          =   300
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   21
            Top             =   3360
            Width           =   1575
         End
         Begin VB.TextBox TxtSugarReview 
            Height          =   300
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   20
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox TxtPulseReview 
            Height          =   300
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   17
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox TxtBPReview 
            Height          =   300
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   15
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox TxtRRReview 
            Height          =   300
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   19
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox TxtWeightReview 
            Height          =   300
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   16
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox TxtTempReview 
            Height          =   300
            Left            =   3240
            MaxLength       =   7
            TabIndex        =   18
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox TxTAssesment 
            Height          =   1575
            Left            =   7080
            TabIndex        =   96
            Top             =   360
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.Label Label41 
            Caption         =   "Observed By :"
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
            Left            =   10560
            TabIndex        =   117
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   "PLAN :"
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
            Left            =   6360
            TabIndex        =   109
            Top             =   2040
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            Caption         =   "BP :"
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
            Left            =   1920
            TabIndex        =   107
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "SPO2 :"
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
            Left            =   1560
            TabIndex        =   106
            Top             =   3375
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "%"
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
            Left            =   4845
            TabIndex        =   105
            Top             =   3375
            Width           =   255
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "SUGAR :"
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
            Left            =   2040
            TabIndex        =   104
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "PULSE: "
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
            Left            =   2040
            TabIndex        =   103
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "mmHg STANDING"
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
            Left            =   4800
            TabIndex        =   102
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "WEIGHT :"
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
            Left            =   1680
            TabIndex        =   101
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "RR :"
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
            Left            =   1560
            TabIndex        =   100
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "b/min"
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
            Left            =   4680
            TabIndex        =   99
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "TEMPERATURE :"
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
            Left            =   1800
            TabIndex        =   98
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "ASSESMENT :"
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
            Left            =   5760
            TabIndex        =   97
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame FrameVitalSigns 
         Caption         =   "Vitals"
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
         Height          =   3975
         Left            =   -74880
         TabIndex        =   54
         Top             =   720
         Width           =   12615
         Begin VB.OptionButton OptRBS 
            Caption         =   "RBS"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8160
            TabIndex        =   90
            Top             =   3120
            Width           =   735
         End
         Begin VB.OptionButton OptFBS 
            Caption         =   "FBS"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6960
            TabIndex        =   89
            Top             =   3120
            Width           =   735
         End
         Begin VB.OptionButton OptOtherVitals 
            Caption         =   "OTHER"
            Height          =   255
            Left            =   5160
            TabIndex        =   88
            Top             =   3600
            Width           =   975
         End
         Begin VB.OptionButton OptSelf 
            Caption         =   "SELF"
            Height          =   255
            Left            =   4080
            TabIndex        =   87
            Top             =   3600
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox TxtOtherVitalSigns 
            Height          =   375
            Left            =   6360
            TabIndex        =   50
            Top             =   3540
            Width           =   4215
         End
         Begin VB.TextBox TxtRR 
            Height          =   300
            Left            =   2640
            MaxLength       =   3
            TabIndex        =   6
            Top             =   2635
            Width           =   1575
         End
         Begin VB.TextBox TxtBloodSugar 
            Height          =   300
            Left            =   8880
            MaxLength       =   3
            TabIndex        =   14
            Top             =   3105
            Width           =   1695
         End
         Begin VB.TextBox TxtWaist 
            Height          =   300
            Left            =   8880
            MaxLength       =   3
            TabIndex        =   13
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox TxtBSA 
            Height          =   285
            Left            =   8880
            TabIndex        =   12
            Top             =   2115
            Width           =   1695
         End
         Begin VB.TextBox TxtBMI 
            Height          =   300
            Left            =   8880
            TabIndex        =   11
            Top             =   1665
            Width           =   1695
         End
         Begin VB.TextBox TxtHeight 
            Height          =   300
            Left            =   8880
            MaxLength       =   4
            TabIndex        =   10
            Top             =   1230
            Width           =   1695
         End
         Begin VB.TextBox TxtWeight 
            Height          =   300
            Left            =   8880
            MaxLength       =   3
            TabIndex        =   9
            Top             =   735
            Width           =   1695
         End
         Begin VB.TextBox TxtBPSittingRight 
            Height          =   300
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   1
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtTemperature 
            Height          =   300
            Left            =   8880
            MaxLength       =   3
            TabIndex        =   8
            Top             =   345
            Width           =   1695
         End
         Begin VB.TextBox TxtSOP2 
            Height          =   300
            Left            =   2640
            MaxLength       =   3
            TabIndex        =   7
            Top             =   3105
            Width           =   1335
         End
         Begin VB.TextBox TxtHR 
            Height          =   285
            Left            =   2640
            MaxLength       =   3
            TabIndex        =   5
            Top             =   2183
            Width           =   1575
         End
         Begin VB.TextBox TxtBPStanding 
            Height          =   285
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   4
            Top             =   1731
            Width           =   1575
         End
         Begin VB.TextBox TxtSulpine 
            Height          =   285
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   3
            Top             =   1279
            Width           =   1575
         End
         Begin VB.TextBox TxtLeft 
            Height          =   285
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   2
            Top             =   827
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "TEMPERATURE"
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
            Left            =   6720
            TabIndex        =   62
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "mmHg STANDING"
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
            Left            =   4200
            TabIndex        =   85
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "b/min"
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
            Left            =   4080
            TabIndex        =   84
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "b/min"
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
            Left            =   4080
            TabIndex        =   83
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "Source of Referral"
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
            Left            =   2040
            TabIndex        =   82
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "mmHg STANDING"
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
            Left            =   4200
            TabIndex        =   80
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "mmHg STANDING"
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
            Left            =   4200
            TabIndex        =   79
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "mmHg STANDING"
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
            Left            =   4200
            TabIndex        =   77
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "RR"
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
            Left            =   960
            TabIndex        =   76
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "MMOL/L"
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
            Left            =   10680
            TabIndex        =   75
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Centimeters"
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
            Left            =   10680
            TabIndex        =   74
            Top             =   2685
            Width           =   1455
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "WAIST CIRCUMFERENCE"
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
            Left            =   6720
            TabIndex        =   73
            Top             =   2685
            Width           =   2055
         End
         Begin VB.Label Label19 
            Caption         =   "M2"
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
            Left            =   10680
            TabIndex        =   72
            Top             =   2145
            Width           =   495
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "BSA"
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
            Left            =   6720
            TabIndex        =   71
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label17 
            Caption         =   "KG/M2"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10635
            TabIndex        =   70
            Top             =   1710
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Meters"
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
            Left            =   10635
            TabIndex        =   68
            Top             =   1275
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "HEIGHT"
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
            Left            =   6720
            TabIndex        =   67
            Top             =   1245
            Width           =   2055
         End
         Begin VB.Label Label13 
            Caption         =   "Kgs"
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
            Left            =   10560
            TabIndex        =   66
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "WEIGHT"
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
            Left            =   6720
            TabIndex        =   65
            Top             =   795
            Width           =   2055
         End
         Begin VB.Label Label11 
            Caption         =   "0"
            Height          =   255
            Left            =   10680
            TabIndex        =   64
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label10 
            Caption         =   "C"
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
            Left            =   10800
            TabIndex        =   63
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4005
            TabIndex        =   61
            Top             =   3120
            Width           =   255
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "SPO2"
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
            Left            =   960
            TabIndex        =   60
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "HR"
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
            Left            =   960
            TabIndex        =   59
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "STANDING   BP:"
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
            TabIndex        =   58
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "SUPINE   BP:"
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
            TabIndex        =   57
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "LEFT"
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
            Left            =   600
            TabIndex        =   56
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "BP: SITTING RIGHT"
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
            TabIndex        =   55
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
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
            Left            =   6720
            TabIndex        =   69
            Top             =   1680
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame6 
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
      TabIndex        =   30
      Top             =   8400
      Width           =   12855
      Begin VB.CommandButton CmdPost 
         Caption         =   "POST"
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
         Left            =   10680
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton OptConsultation 
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
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   35
         Top             =   350
         Width           =   2415
      End
      Begin VB.OptionButton OptDoctors 
         Caption         =   "TO DOCTORS"
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
         Height          =   375
         Index           =   2
         Left            =   2970
         TabIndex        =   34
         Top             =   350
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton OptCashier 
         Caption         =   "TO CASHIER"
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
         Height          =   375
         Index           =   3
         Left            =   5280
         TabIndex        =   33
         Top             =   350
         Width           =   1935
      End
      Begin VB.OptionButton OptPharmacy 
         Caption         =   "TO PHARMACY"
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
         Height          =   375
         Index           =   4
         Left            =   7350
         TabIndex        =   32
         Top             =   350
         Width           =   1935
      End
      Begin VB.OptionButton OptLab 
         Caption         =   "TO LAB"
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
         Height          =   375
         Index           =   0
         Left            =   9360
         TabIndex        =   31
         Top             =   350
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BMI Calculator"
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
      Left            =   11280
      TabIndex        =   29
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Navigation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   13080
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton CmdPacemaker 
         Caption         =   "Pacemaker Chart"
         Height          =   375
         Left            =   120
         TabIndex        =   115
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton CmdABI 
         Caption         =   "ABI Examination"
         Height          =   375
         Left            =   120
         TabIndex        =   114
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton CmdWarfarin 
         Caption         =   "Warfarin  Chart"
         Height          =   375
         Left            =   120
         TabIndex        =   113
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton CmdBloodSugarChart 
         Caption         =   "Blood Sugar Chart"
         Height          =   375
         Left            =   120
         TabIndex        =   112
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton BPChart 
         Caption         =   "Blood Pressure Cart"
         Height          =   375
         Left            =   120
         TabIndex        =   111
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Update Medical History"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   86
         Top             =   6120
         Width           =   1695
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton FrmExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   8520
         Width           =   1695
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh List"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update Vital Signs"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update Chief Complaint"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   7320
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Patient for Observation"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   12855
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   2535
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   4471
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
         BackColor       =   65535
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         BackColorAlternate=   49344
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
         Editable        =   -1  'True
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
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Patients to be Seen"
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
      Left            =   8400
      TabIndex        =   91
      Top             =   3240
      Width           =   3135
   End
End
Attribute VB_Name = "FrmObservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsInsert As New ADODB.Recordset
Dim RsGrid As New ADODB.Recordset
Dim lvObservationCardNumber As String
Dim lvObservationVisitNumber As Long
Dim ItemSelected As Integer
Public Function UpdateMedicalHistory(ByRef CardNo, ByRef VisitNo) As Boolean
    On Error GoTo Errorhandler
    If lvObservationCardNumber = "" Then Exit Function
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM OBSERVATION_MEDICAL_HISTORY WHERE CARDNUMBER = '" & lvObservationCardNumber & "' AND VISITNUMBER = '" & lvObservationVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !DIABETES = ChkDiabetes
                        !HYPERTENSION = ChkHypertension
                        !HIGHCHOLESTROL = ChkHighCholestrol
                        !GOUT = ChkGout
                        !ASTHMA = ChkAsthma
                        !CANCER = ChkCancer
                        !PHEUMONIAVACCINE = ChkPheumoniaVaccine
                        !HEPATITISBVACCINE = ChkHepatitisBvaccine
                        !PENICILINE = ChkPeniciline
                        !SULPHURDRUGS = ChkSulphurDrugs
                        !OTHERS = TxtOtherAllergies
                    .Update
                    UpdateMedicalHistory = True
            End With
Exit Function
Errorhandler:
        MsgBox Err.Description
End Function
Public Function UpdateVitalSigns(ByRef CardNo, ByRef VisitNo) As Boolean
    On Error GoTo Errorhandler
    If lvObservationCardNumber = "" Then Exit Function
    
    
    'CHECK IF IT IS A CLINIC REVIEW AND UPDATE
    If OptClinicReview.Value = True Then UpdateVitalSigns = UpdateClinicReview(CardNo, VisitNo): Exit Function
    
    
    'CHECK IF ALL DATA HAS BEEN INPUT (COMPREHENSIVE REVIEW)
    If TxtBPSittingRight = "" Then MsgBox "Please Enter the BP Sitting before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtLeft = "" Then MsgBox "Please Enter the BP Left before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtSulpine = "" Then MsgBox "Please Enter the SUPINE before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtBPStanding = "" Then MsgBox "Please Enter the BP Standing before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtHR = "" Then MsgBox "Please Enter the   HR   before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtRR = "" Then MsgBox "Please Enter the  RR  before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtSOP2 = "" Then MsgBox "Please Enter the   SPO2  before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtTemperature = "" Then MsgBox "Please Enter the Temperature before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtWeight = "" Then MsgBox "Please Enter the Weight before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtHeight = "" Then MsgBox "Please Enter the Height before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtWaist = "" Then MsgBox "Please Enter the Waist before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtBloodSugar = "" Then MsgBox "Please Enter the Blood Sugar before Posting", vbInformation, "Blank Field.": Exit Function
    
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM OBSERVATION_VITAL_SIGNS WHERE CARDNUMBER = '" & lvObservationCardNumber & "' AND VISITNUMBER = '" & lvObservationVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !BPsittingRight = TxtBPSittingRight
                        !BPLEFT = TxtLeft
                        !SULPINE = TxtSulpine
                        !STANDING = TxtBPStanding
                        !HR = TxtHR
                        !RR = TxtRR
                        !SPO2 = TxtSOP2
                        !Temperature = TxtTemperature
                        !Weight = TxtWeight
                        !Height = TxtHeight
                        !BMI = TxtBMI
                        !BSA = TxtBSA
                        !WAIST = TxtWaist
                        !FBS = OptFBS.Value
                        !RBS = OptRBS.Value
                        !BLOODSUGAR = TxtBloodSugar
                        !SELFREFERAL = OptSelf
                        !OTHER = TxtOtherVitalSigns
                    .Update
                    UpdateVitalSigns = True
            End With
Exit Function
Errorhandler:
        MsgBox Err.Description
End Function
Public Function UpdateClinicReview(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    If lvObservationCardNumber = "" Then Exit Function
    
    'CHECK IF ALL DATA HAS BEEN INPUT
    If TxtBPReview = "" Then MsgBox "Please Enter the BP Sitting before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtPulseReview = "" Then MsgBox "Please Enter the Pulse Measurements before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtRRReview = "" Then MsgBox "Please Enter the RR before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtSugarReview = "" Then MsgBox "Please Enter the Sugar before Posting", vbInformation, "Blank Field.": Exit Function
    If TxtSPO2Review = "" Then MsgBox "Please Enter the   SPO2  before Posting", vbInformation, "Blank Field.": Exit Function
    'If CboDiagnosisReview = "" Then MsgBox "Please Enter the  Diagnosis  before Posting", vbInformation, "Blank Field.": Exit Function
    
'''    If d = "" Then MsgBox "Please Enter the   SPO2  before Posting", vbInformation, "Blank Field.": Exit Sub
'''    If TxtTemperature = "" Then MsgBox "Please Enter the Temperature before Posting", vbInformation, "Blank Field.": Exit Sub
'''    If TxtWeight = "" Then MsgBox "Please Enter the Weight before Posting", vbInformation, "Blank Field.": Exit Sub
'''    If TxtHeight = "" Then MsgBox "Please Enter the Height before Posting", vbInformation, "Blank Field.": Exit Sub
'''    If TxtWaist = "" Then MsgBox "Please Enter the Waist before Posting", vbInformation, "Blank Field.": Exit Sub
'''    If TxtBloodSugar = "" Then MsgBox "Please Enter the Blood Sugar before Posting", vbInformation, "Blank Field.": Exit Sub
    
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM OBSERVATION_CLINIC_REVIEW WHERE CARDNUMBER = '" & lvObservationCardNumber & "' AND VISITNUMBER = '" & lvObservationVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !BP = TxtBPReview
                        !Weight = TxtWeightReview
                        !PULSE = TxtPulseReview
                        !Temperature = TxtTempReview
                        !RR = TxtRRReview
                        !SUGAR = TxtSugarReview
                        !SPO2 = TxtSPO2Review
                        '!ASSESMENT = TxTAssesment
                        '!REVIEWPLAN = TxtPlan
                        If ChkHighlighBPtInRed.Value = 1 Then !HIGHLIGHTBP = ChkHighlighBPtInRed.Value
                        If ChkHighlightSugarInRed.Value = 1 Then !HIGHLIGHTSUGAR = ChkHighlightSugarInRed.Value
                    .Update
                    UpdateClinicReview = True
            End With
Exit Function
Errorhandler:
        MsgBox Err.Description
End Function
Public Function UpdateChiefComplaint(ByRef CardNo, ByRef VisitNo) As Boolean
    On Error GoTo Errorhandler
    If lvObservationCardNumber = "" Then Exit Function
    If RsInsert.State = 1 Then Set RsInsert = Nothing
        RsInsert.Open "SELECT * FROM OBSERVATION_CHIEF_COMPLAINT WHERE CARDNUMBER = '" & lvObservationCardNumber & "' AND VISITNUMBER = '" & lvObservationVisitNumber & "'", Conn, adOpenStatic, adLockOptimistic
            With RsInsert
                If .EOF = True Then
                    .AddNew
                End If
                        !CardNumber = CardNo
                        !VISITNUMBER = VisitNo
                        !CHIEFCOMPLAINT = TxtChiefComplaint
                    .Update
                    UpdateChiefComplaint = True
            End With
Exit Function
Errorhandler:
        MsgBox Err.Description
End Function

Private Sub BPChart_Click()
    lvOptionalCardNo = lvObservationCardNumber
    lvOptionalVisitNo = lvObservationVisitNumber
    If lvOptionalCardNo = "" Then Exit Sub

    FrmBPChart.Show 1
End Sub

Private Sub ChkHighlighBPtInRed_Click()
    If ChkHighlighBPtInRed.Value = 1 Then
        TxtBPReview.BackColor = vbRed
    Else
        TxtBPReview.BackColor = vbWhite
    End If
End Sub

Private Sub ChkHighlightSugarInRed_Click()
    If ChkHighlightSugarInRed.Value = 1 Then
        TxtSugarReview.BackColor = vbRed
    Else
        TxtSugarReview.BackColor = vbWhite
    End If
End Sub

Private Sub CmdABI_Click()
    lvOptionalCardNo = lvObservationCardNumber
    lvOptionalVisitNo = lvObservationVisitNumber
    If lvOptionalCardNo = "" Then Exit Sub
    FrmABI.Show 1
End Sub

Private Sub CmdBloodSugarChart_Click()
    lvOptionalCardNo = lvObservationCardNumber
    lvOptionalVisitNo = lvObservationVisitNumber
    If lvOptionalCardNo = "" Then Exit Sub

    FrmBloodSugar.Show 1
End Sub

Private Sub CmdPacemaker_Click()
    lvOptionalCardNo = lvObservationCardNumber
    lvOptionalVisitNo = lvObservationVisitNumber
    If lvOptionalCardNo = "" Then Exit Sub

    FrmPacemaker.Show 1
End Sub

Private Sub CmdPost_Click()
    On Error GoTo Errorhandler
    
    'IF ITS BEING SENT TO CONSULTATION THEN DONT VALIDATE ANYTHING. JUST RETURN PATIENT TO CONSULTATION. BY ODER, TOWN CLERK.
    If ItemSelected = 1 Then GoTo SEND_TO_CONSULTATION
    
    'UPDATE THE THREE TABS FIRST THEN MOVE TO NEXT QUEUE
    If UpdateMedicalHistory(lvObservationCardNumber, lvObservationVisitNumber) = False Then Exit Sub
    If UpdateVitalSigns(lvObservationCardNumber, lvObservationVisitNumber) = False Then Exit Sub
    If UpdateChiefComplaint(lvObservationCardNumber, lvObservationVisitNumber) = False Then Exit Sub
    
    If OptClinicReview.Value = True Then
        Conn.Execute "UPDATE COMPLAINS SET BP = '" & TxtBPReview & "',WEIGHT = '" & TxtWeightReview & "',HEIGHT = '" & TxtHeight & "',BMINDEX = '" & TxtBMI & "',NURSE = '" & GlbCurrentUser & "',INUSE = '0' WHERE CARDNUMBER = '" & lvObservationCardNumber & "'"
    Else
        Conn.Execute "UPDATE COMPLAINS SET BP = '" & TxtBPSittingRight & "',WEIGHT = '" & TxtWeight & "',HEIGHT = '" & TxtHeight & "',BMINDEX = '" & TxtBMI & "',NURSE = '" & GlbCurrentUser & "',INUSE = '0' WHERE CARDNUMBER = '" & lvObservationCardNumber & "'"
    End If

SEND_TO_CONSULTATION:
    Select Case ItemSelected
   Case 1
            Resp = MsgBox("Are you sure you wish to send this record back to Front Desk ?", vbQuestion + vbYesNo + vbDefaultButton2)
                If Resp = vbNo Then Exit Sub
            DUMMY = SendPatient(EnumConsultation, lvObservationCardNumber, GlbSysDate, lvObservationVisitNumber)
            'FrmPatients.Show
            'Unload Me
    Case 2
        'To Doctor
            DUMMY = SendPatient(EnumDoctors, lvObservationCardNumber, GlbSysDate, lvObservationVisitNumber)
            If FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'NurseDoctorRolesCombined'") = 1 Then
                FrmWaitingRoom.Show
            End If
            'Unload Me
            CmdRefresh_Click
   Case 3
        'To Cashier
            DUMMY = SendPatient(EnumCashier, lvObservationCardNumber, GlbSysDate)
            'FrmCashier.Show
            'Unload Me
    Case 4
        'To Pharmacy
            DUMMY = SendPatient(EnumPharmacy, lvObservationCardNumber, GlbSysDate)
            'FrmPharmacy.Show
            'Unload Me
    End Select
    CmdRefresh_Click
Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub

Private Sub CmdRefresh_Click()
    ClearText FrmObservation
    FillObservation
    CmdPost.Enabled = False
End Sub

Private Sub CmdWarfarin_Click()
    lvOptionalCardNo = lvObservationCardNumber
    lvOptionalVisitNo = lvObservationVisitNumber
    If lvOptionalCardNo = "" Then Exit Sub
    FrmWarfarin.Show 1
End Sub

Private Sub Command2_Click()
    UpdateVitalSigns lvObservationCardNumber, lvObservationVisitNumber
End Sub

Private Sub Command3_Click()
    UpdateChiefComplaint lvObservationCardNumber, lvObservationVisitNumber
End Sub

Private Sub Command4_Click()
    UpdateMedicalHistory lvObservationCardNumber, lvObservationVisitNumber
End Sub

Private Sub Form_Load()
    FillObservation
    centerform Me
    
    TabObservation.Tab = 1
    TxtOtherVitalSigns.Enabled = False
    
     'ASSIGN FORM NAME AS CURRENT FORM
    GlbCurrentForm = EnumObservation
    ManageProcessFlow EnumObservation
    
    OptDoctors_Click (2)
    If OptDoctors.Item(2).Value = False Then OptDoctors.Item(2).Value = True
    
    TxtNurse = GlbCurrentUser
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BlnViewMeasurements = False
End Sub

Private Sub Frame7_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub FrmExit_Click()
    Unload Me
End Sub

Private Sub Grid_Click()
On Error GoTo Errorhandler
Dim i As Double
    
    For i = 1 To Grid.Rows - 1
        If Grid.TextMatrix(i, 9) = "SELECT" Then GoTo SKIPHEADER
        
        'TEMPORARY FIX
            'KARI = Grid.(Grid.Row)
        
         If Grid.TextMatrix(i, 9) = "" Then Grid.TextMatrix(i, 9) = 0
         If Grid.Rows <= 1 Then GoTo SKIPHEADER
         If CStr(Grid.TextMatrix(i, 9)) <> -1 Then
             Dim StrCardNumber, strVisitNumber, StrPatientName, StrBillingCompany, StrIDNumber, StrBP, StrWeight, StrHeight
             
             lvObservationCardNumber = Grid.TextMatrix(Grid.Row, 0)
             lvObservationVisitNumber = Grid.TextMatrix(Grid.Row, 1)
             StrCardNumber = Grid.TextMatrix(Grid.Row, 0)
             strVisitNumber = Grid.TextMatrix(Grid.Row, 1)
             StrPatientName = Grid.TextMatrix(Grid.Row, 2)
             StrBillingCompany = Grid.TextMatrix(Grid.Row, 3)
             StrIDNumber = Grid.TextMatrix(Grid.Row, 4)
             StrBP = Grid.TextMatrix(Grid.Row, 5)
             StrWeight = Grid.TextMatrix(Grid.Row, 6)
             StrHeight = Grid.TextMatrix(Grid.Row, 7)
             
             StrNurse = FindRecord("COMPLAINS", "NURSE", "VISITNUMBER = '" & strVisitNumber & "'")
             'If StrNurse = "" Then StrNurse = GlbCurrentUser
             
             Grid.Clear: Grid.Col = 1: Grid.Rows = 1
             Grid.FormatString = "CARD NUMBER|   PATIENTS FULL NAME  |   BILLING COMPANY     |ID NUMBER |BLOOD PRESSURE |WEIGHT | HEIGHT |  NURSE   ."
             Grid.AddItem StrCardNumber & vbTab & strVisitNumber & vbTab & StrPatientName & vbTab & StrBillingCompany & vbTab & StrIDNumber & vbTab & StrBP & vbTab & StrWeight & vbTab & StrHeight & vbTab & StrNurse
             Exit For
        End If
SKIPHEADER:
    Next
    
    PopulateMedicalHistory lvObservationCardNumber, lvObservationVisitNumber
    PopulateVitalSigns lvObservationCardNumber, lvObservationVisitNumber
    PopulateClinicReview lvObservationCardNumber, lvObservationVisitNumber
    PopulateChiefComplaint lvObservationCardNumber, lvObservationVisitNumber
    
    centerform Me
    OptComprehensive.Value = True
    OptComprehensive_Click
    If FindRecord("COMPLAINS", "TODOCTORS", "CARDNUMBER = '" & lvObservationCardNumber & "' AND VISITNUMBER = '" & lvObservationVisitNumber & "'") <> True Then
        CmdPost.Enabled = True
    End If
Exit Sub
Errorhandler:
        MsgBox Err.Description
Exit Sub
Resume
End Sub
Private Sub FillObservation()
   On Error GoTo Errorhandler
   KARI = GlbSysDate
    Grid.Clear: TxtCount = 0
    Grid.Rows = 1
    Grid.Cols = 9
    For i = 0 To Grid.Cols - 1
        Grid.ColAlignment(i) = flexAlignCenterCenter
    Next i
    
    Grid.ColAlignment(1) = flexAlignCenterCenter
 
    Grid.ColDataType(8) = flexDTBoolean
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER| VISIT NUMBER | PATIENTS FULL NAME                           |BILLING COMPANY     |ID NUMBER |BLOOD PRESSURE |WEIGHT | HEIGHT IN METERS |  BMI  |SELECT "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        If BlnViewMeasurements = True Then
            RsGrid.Open "SELECT PATIENT_DETAILS.*, COMPLAINS.VISITNUMBER AS VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND COMPLAINS.CARDNUMBER = '" & StrDocCardNo & "'", Conn, adOpenDynamic, adLockOptimistic
            
            CmdPost.Enabled = False
            
        Else
            RsGrid.Open "SELECT PATIENT_DETAILS.*,COMPLAINS.VISITNUMBER AS VISITNUMBER FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND COMPLAINS.TOOBSERVATION = '1' AND DISMISSED = 'FALSE'", Conn, adOpenDynamic, adLockOptimistic
        End If
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                        TxtCount = TxtCount + 1
                        .MoveNext
                    Wend
                End With
            End If
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub PopulateMedicalHistory(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM OBSERVATION_MEDICAL_HISTORY WHERE CARDNUMBER = '" & CardNo & "' ORDER BY VISITNUMBER DESC", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                ChkDiabetes = !DIABETES
                ChkHypertension = !HYPERTENSION
                ChkHighCholestrol = !HIGHCHOLESTROL
                ChkGout = !GOUT
                ChkAsthma = !ASTHMA
                ChkCancer = !CANCER
                ChkPheumoniaVaccine = !PHEUMONIAVACCINE
                ChkHepatitisBvaccine = !HEPATITISBVACCINE
                ChkPeniciline = !PENICILINE
                ChkSulphurDrugs = !SULPHURDRUGS
                TxtOtherAllergies = !OTHERS
            End With
        End If
    RsMeasurements.Close
    
Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub

Private Sub PopulateVitalSigns(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    Dim GetComprehensive As Boolean
    
LAST_COMPREHENSIVE_VITALS_RECORDED:
    If GetComprehensive = True Then
        Set RsMeasurements = Nothing
        RsMeasurements.Open "SELECT * FROM OBSERVATION_VITAL_SIGNS WHERE CARDNUMBER = '" & CardNo & "' order by VISITNUMBER DESC", Conn, adOpenStatic, adLockOptimistic
            If RsMeasurements.EOF = False Then
                GoTo PROCEED
            Else
                Exit Sub
            End If
    End If
    
    RsMeasurements.Open "SELECT * FROM OBSERVATION_VITAL_SIGNS WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = True Then GetComprehensive = True: GoTo LAST_COMPREHENSIVE_VITALS_RECORDED ' IF NOT, THEN CONTINUE.
        If RsMeasurements.EOF = False Then
PROCEED:
            With RsMeasurements
                TxtBPSittingRight = !BPsittingRight
                TxtLeft = !BPLEFT
                TxtSulpine = !SULPINE
                TxtBPStanding = !STANDING
                TxtHR = !HR
                TxtRR = !RR
                TxtSOP2 = !SPO2
                TxtTemperature = !Temperature
                TxtWeight = !Weight
                TxtHeight = !Height
                TxtBMI = !BMI
                TxtBSA = !BSA
                TxtWaist = !WAIST
                If IsNull(!FBS) Then OptFBS.Value = False Else OptFBS.Value = True
                If IsNull(!RBS) Then OptRBS.Value = False Else OptRBS.Value = True
                TxtBloodSugar = !BLOODSUGAR
                ChkSelfReferal = !SELFREFERAL
                TxtOtherVitalSigns = !OTHER
            End With
        End If
    RsMeasurements.Close
Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub
Private Sub PopulateClinicReview(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM OBSERVATION_CLINIC_REVIEW WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                TxtBPReview = !BP
                TxtWeightReview = !Weight
                TxtPulseReview = !PULSE
                TxtTempReview = !Temperature
                TxtRRReview = !RR
                TxtSugarReview = !SUGAR
                TxtSPO2Review = !SPO2
                If !REVIEWPLAN <> "" Then TxtPlan = !REVIEWPLAN
                If Not IsNull(!ASSESMENT) Then TxTAssesment.Text = !ASSESMENT
                If !HIGHLIGHTBP <> "" Then ChkHighlighBPtInRed = !HIGHLIGHTBP
                If !HIGHLIGHTSUGAR <> "" Then ChkHighlightSugarInRed = !HIGHLIGHTSUGAR
                If FindRecord("COMPLAINS", "NURSE", "VISITNUMBER = '" & VisitNo & "'") <> "" Then
                    TxtNurse = FindRecord("COMPLAINS", "NURSE", "VISITNUMBER = '" & VisitNo & "'")
                End If
            End With
        End If
    RsMeasurements.Close
    
Exit Sub
Errorhandler:
        MsgBox Err.Description
       ' Resume
End Sub
Private Sub PopulateChiefComplaint(ByRef CardNo, ByRef VisitNo)
    On Error GoTo Errorhandler
    Dim RsMeasurements As New ADODB.Recordset
    
    RsMeasurements.Open "SELECT * FROM OBSERVATION_CHIEF_COMPLAINT WHERE CARDNUMBER = '" & CardNo & "' AND VISITNUMBER = '" & VisitNo & "'", Conn, adOpenStatic, adLockOptimistic
        If RsMeasurements.EOF = False Then
            With RsMeasurements
                TxtChiefComplaint = !CHIEFCOMPLAINT 'ChiefComplaint
            End With
        End If
    RsMeasurements.Close
    
Exit Sub
Errorhandler:
        MsgBox Err.Description
End Sub

Public Sub ManageProcessFlow(ActiveForm)
    Dim RsControls As New ADODB.Recordset
    RsControls.Open "SELECT * FROM PROCESSFLOW WHERE SCREENID = '" & ActiveForm & "'", Conn, adOpenStatic, adLockOptimistic
        If RsControls.EOF = False Then
            With RsControls
                If !CONSULTATION = 1 Then OptConsultation.Item(1).Enabled = True
                If !DOCTORS = 1 Then OptDoctors.Item(2).Enabled = True
                If !CASHIER = 1 Then OptCashier.Item(3).Enabled = True
                If !PHARMACY = 1 Then OptPharmacy.Item(4).Enabled = True
                If !LAB = 1 Then OptLab.Item(5).Enabled = True
            End With
        End If
End Sub


Private Sub OptClinicReview_Click()
    FrameVitalSigns.Visible = False
    FrameClinicReview.Visible = True
End Sub

Private Sub OptComprehensive_Click()
    FrameVitalSigns.Visible = True
    FrameClinicReview.Visible = False
End Sub

Private Sub OptConsultation_Click(Index As Integer)
    ItemSelected = OptConsultation.Item(Index).Index
End Sub

Private Sub OptDoctors_Click(Index As Integer)
    ItemSelected = OptDoctors.Item(Index).Index
End Sub

Private Sub OptOtherVitals_Click()
    If OptOtherVitals.Value = True Then
        TxtOtherVitalSigns.Enabled = True
        TxtOtherVitalSigns = ""
    ElseIf OptOtherVitals.Value = False Then
        TxtOtherVitalSigns.Enabled = False
    End If
End Sub

Private Sub OptSelf_Click()
    If OptSelf.Value = True Then
        TxtOtherVitalSigns.Enabled = False
        TxtOtherVitalSigns = ""
    ElseIf OptSelf.Value = False Then
        TxtOtherVitalSigns.Enabled = True
    End If
End Sub

Private Sub TxtBloodSugar_Change()
    KISH = ValidateDataType_Advice(TxtBloodSugar, 0)
    If KISH = False Then TxtBloodSugar.Text = Mid(TxtBloodSugar, 1, Len(TxtBloodSugar) - 1)
End Sub

Private Sub TxtHeight_Change()

    KISH = ValidateDataType_Advice(TxtHeight, 0)
    If KISH = False Then TxtHeight.Text = Mid(TxtHeight, 1, Len(TxtHeight) - 1)
    
    If TxtHeight <> "" And TxtWeight <> "" Then TxtBMI = CalculateBMI(TxtWeight, TxtHeight)
End Sub

Private Sub TxtHR_Change()
    KISH = ValidateDataType_Advice(TxtHR, 0)
    If KISH = False Then TxtHR.Text = Mid(TxtHR, 1, Len(TxtHR) - 1)
End Sub

Private Sub TxtRR_Change()
    KISH = ValidateDataType_Advice(TxtRR, 0)
    If KISH = False Then TxtRR.Text = Mid(TxtRR, 1, Len(TxtRR) - 1)
End Sub

Private Sub TxtSOP2_Change()
    KISH = ValidateDataType_Advice(TxtSOP2, 0)
    If KISH = False Then TxtSOP2.Text = Mid(TxtSOP2, 1, Len(TxtSOP2) - 1)
End Sub

Private Sub TxtTemperature_Change()
    KISH = ValidateDataType_Advice(TxtTemperature, 0)
    If KISH = False Then TxtTemperature.Text = Mid(TxtTemperature, 1, Len(TxtTemperature) - 1)
End Sub

Private Sub TxtWaist_Change()
    KISH = ValidateDataType_Advice(TxtWaist, 0)
    If KISH = False Then TxtWaist.Text = Mid(TxtWaist, 1, Len(TxtWaist) - 1)
End Sub

Private Sub TxtWeight_Change()
    KISH = ValidateDataType_Advice(TxtWeight, 0)
    If KISH = False Then TxtWeight.Text = Mid(TxtWeight, 1, Len(TxtWeight) - 1)
    
    If TxtHeight <> "" And TxtWeight <> "" Then TxtBMI = CalculateBMI(TxtWeight, TxtHeight)
End Sub
