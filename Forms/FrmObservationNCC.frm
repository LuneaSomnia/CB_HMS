VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmObservationNCC 
   Caption         =   "Nurse Observations"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Previous Medical History"
      TabPicture(0)   =   "FrmObservationNCC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Vital Signs"
      TabPicture(1)   =   "FrmObservationNCC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Chief Complaint"
      TabPicture(2)   =   "FrmObservationNCC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame7 
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
         Height          =   5175
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   12615
         Begin VB.TextBox Text18 
            Height          =   375
            Left            =   6360
            TabIndex        =   82
            Top             =   4620
            Width           =   4215
         End
         Begin VB.CheckBox Check13 
            Caption         =   "OTHER"
            Height          =   255
            Left            =   5400
            TabIndex        =   81
            Top             =   4680
            Width           =   975
         End
         Begin VB.CheckBox Check12 
            Caption         =   "SELF"
            Height          =   255
            Left            =   3960
            TabIndex        =   80
            Top             =   4680
            Width           =   855
         End
         Begin VB.TextBox Text15 
            Height          =   300
            Left            =   2520
            TabIndex        =   70
            Top             =   3360
            Width           =   1575
         End
         Begin VB.TextBox Text14 
            Height          =   300
            Left            =   8880
            TabIndex        =   67
            Top             =   3825
            Width           =   1695
         End
         Begin VB.TextBox Text13 
            Height          =   300
            Left            =   8880
            TabIndex        =   64
            Top             =   3260
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   8880
            TabIndex        =   61
            Top             =   2713
            Width           =   1695
         End
         Begin VB.TextBox Text11 
            Height          =   300
            Left            =   8880
            TabIndex        =   58
            Top             =   2151
            Width           =   1695
         End
         Begin VB.TextBox Text10 
            Height          =   300
            Left            =   8880
            TabIndex        =   55
            Top             =   1589
            Width           =   1695
         End
         Begin VB.TextBox Text9 
            Height          =   300
            Left            =   8880
            TabIndex        =   52
            Top             =   1027
            Width           =   1695
         End
         Begin VB.TextBox Text8 
            Height          =   300
            Left            =   2520
            TabIndex        =   50
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text7 
            Height          =   300
            Left            =   8880
            TabIndex        =   47
            Top             =   465
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   2520
            TabIndex        =   44
            Top             =   3825
            Width           =   1335
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   2520
            TabIndex        =   42
            Top             =   2865
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   2520
            TabIndex        =   40
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2520
            TabIndex        =   38
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2520
            TabIndex        =   36
            Top             =   960
            Width           =   1575
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
            Left            =   2160
            TabIndex        =   83
            Top             =   4680
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
            Left            =   4080
            TabIndex        =   74
            Top             =   1440
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
            Left            =   4080
            TabIndex        =   73
            Top             =   960
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
            Left            =   4080
            TabIndex        =   71
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "b/min RR"
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
            Left            =   840
            TabIndex        =   69
            Top             =   3360
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
            Left            =   10725
            TabIndex        =   68
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "BLOOD SUGAR"
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
            TabIndex        =   66
            Top             =   3840
            Width           =   2055
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
            TabIndex        =   65
            Top             =   3285
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
            TabIndex        =   63
            Top             =   3280
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
            TabIndex        =   62
            Top             =   2760
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
            TabIndex        =   60
            Top             =   2760
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
            TabIndex        =   59
            Top             =   2160
            Width           =   855
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
            TabIndex        =   57
            Top             =   2160
            Width           =   2055
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
            TabIndex        =   56
            Top             =   1605
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
            TabIndex        =   54
            Top             =   1605
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
            Left            =   10680
            TabIndex        =   53
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "WIEGHT"
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
            TabIndex        =   51
            Top             =   1035
            Width           =   2055
         End
         Begin VB.Label Label11 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10680
            TabIndex        =   49
            Top             =   360
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
            TabIndex        =   48
            Top             =   480
            Width           =   375
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
            TabIndex        =   46
            Top             =   480
            Width           =   2055
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
            Left            =   3885
            TabIndex        =   45
            Top             =   3840
            Width           =   255
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "b/min SPO2"
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
            Left            =   840
            TabIndex        =   43
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "mmHg RR"
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
            Left            =   840
            TabIndex        =   41
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "STANDING"
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
            Left            =   360
            TabIndex        =   39
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "mmHg SULPINE"
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
            Left            =   360
            TabIndex        =   37
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "mmHg LEFT"
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
            TabIndex        =   35
            Top             =   960
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
            Left            =   360
            TabIndex        =   34
            Top             =   480
            Width           =   2055
         End
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
         ForeColor       =   &H8000000D&
         Height          =   2055
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   12615
         Begin VB.TextBox TxtOthers 
            Height          =   375
            Left            =   3240
            TabIndex        =   27
            Text            =   "OTHERS"
            ToolTipText     =   "OTHERS"
            Top             =   1320
            Width           =   9255
         End
         Begin VB.CheckBox Check11 
            Caption         =   "OTHERS"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox Check10 
            Caption         =   "SULPHUR DRUGS"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   2895
         End
         Begin VB.CheckBox Check9 
            Caption         =   "PENICILINE"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   480
            Width           =   2895
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
         ForeColor       =   &H8000000D&
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   12615
         Begin VB.CheckBox Check8 
            Caption         =   "HEPATITIS B VACCINE"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Width           =   2655
         End
         Begin VB.CheckBox Check7 
            Caption         =   "PHEUMONIA VACCINE"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   3255
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
         ForeColor       =   &H8000000D&
         Height          =   1575
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   12615
         Begin VB.CheckBox Check6 
            Caption         =   "CANCER"
            Height          =   375
            Left            =   4440
            TabIndex        =   32
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox Check5 
            Caption         =   "ASTHMA"
            Height          =   495
            Left            =   4440
            TabIndex        =   31
            Top             =   600
            Width           =   3135
         End
         Begin VB.CheckBox Check4 
            Caption         =   "GOUT"
            Height          =   375
            Left            =   4440
            TabIndex        =   30
            Top             =   240
            Width           =   2655
         End
         Begin VB.CheckBox Check3 
            Caption         =   "HIGH CHOLESTROL"
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   1080
            Width           =   3135
         End
         Begin VB.CheckBox Check2 
            Caption         =   "HYPERTENSION"
            Height          =   495
            Left            =   240
            TabIndex        =   28
            Top             =   600
            Width           =   3135
         End
         Begin VB.CheckBox Check1 
            Caption         =   "DIABETES"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   3975
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
         Height          =   5175
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Width           =   12615
         Begin VB.Frame Frame10 
            Caption         =   "Source of History"
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
            Height          =   1575
            Left            =   120
            TabIndex        =   78
            Top             =   3480
            Width           =   12375
            Begin VB.TextBox Text17 
               Height          =   1095
               Left            =   120
               TabIndex        =   79
               Top             =   360
               Width           =   12135
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Regular Medication "
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
            Height          =   1815
            Left            =   120
            TabIndex        =   76
            Top             =   1680
            Width           =   12375
            Begin VB.TextBox Text16 
               Height          =   1335
               Left            =   120
               TabIndex        =   77
               Top             =   360
               Width           =   12135
            End
         End
         Begin VB.TextBox Text5 
            Height          =   1215
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   12375
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Patient for Observation"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   12855
      Begin VSFlex6DAOCtl.vsFlexGrid Grid 
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3413
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
      TabIndex        =   8
      Top             =   8280
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   280
         Width           =   2055
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
         TabIndex        =   12
         Top             =   280
         Width           =   1935
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
         Left            =   5340
         TabIndex        =   11
         Top             =   280
         Width           =   1575
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
         Height          =   495
         Index           =   4
         Left            =   7350
         TabIndex        =   10
         Top             =   240
         Width           =   1575
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
         Height          =   495
         Index           =   0
         Left            =   9360
         TabIndex        =   9
         Top             =   240
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
      TabIndex        =   7
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
      Height          =   9015
      Left            =   13080
      TabIndex        =   0
      Top             =   120
      Width           =   1935
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   8400
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
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Update Medical History "
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   240
         TabIndex        =   3
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update Family History"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   5640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update Current Medication"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   6480
         Visible         =   0   'False
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmObservationNCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    SSTab1.Tab = 0
End Sub

Private Sub Frame8_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label23_Click()

End Sub

Private Sub SSTab1_DblClick()

End Sub
