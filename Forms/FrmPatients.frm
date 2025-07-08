VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPatients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patients Details"
   ClientHeight    =   9105
   ClientLeft      =   4545
   ClientTop       =   1800
   ClientWidth     =   14970
   Icon            =   "FrmPatients.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   14970
   Begin TabDlg.SSTab PatientsTab 
      Height          =   7935
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Patient Details"
      TabPicture(0)   =   "FrmPatients.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ChkConsultation"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ChkPreserve"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "PatientTab2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame16"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmdBrowse"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Appointments"
      TabPicture(1)   =   "FrmPatients.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame17"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "OptExistingClient"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "OptNewClient"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CmdBookSearch"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TxtBookCardNumber"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame7"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "TabAppointments"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label25"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame17 
         Caption         =   "Show Specific Appointment Date"
         ForeColor       =   &H00FF00FF&
         Height          =   2175
         Left            =   -64560
         TabIndex        =   126
         Top             =   3960
         Width           =   2655
         Begin VB.CommandButton CmdDateFilter 
            Caption         =   "Show"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   127
            Top             =   1440
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DtDateFilter 
            Height          =   375
            Left            =   120
            TabIndex        =   128
            Top             =   960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            Format          =   70451201
            CurrentDate     =   41202
         End
         Begin VB.CheckBox ChkDateFilter 
            Caption         =   "Appointments By Date"
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "Get Picture"
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
         Left            =   11760
         TabIndex        =   123
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Frame Frame16 
         Caption         =   "Picture"
         Height          =   3135
         Left            =   10200
         TabIndex        =   122
         Top             =   4680
         Width           =   2895
         Begin VB.CommandButton CmdExplode 
            Caption         =   "---"
            Height          =   195
            Left            =   720
            TabIndex        =   135
            Top             =   0
            Width           =   375
         End
         Begin VB.PictureBox Picture1 
            Height          =   1815
            Left            =   480
            ScaleHeight     =   1755
            ScaleWidth      =   1875
            TabIndex        =   124
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Image ImgPreview 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DragMode        =   1  'Automatic
            Height          =   2775
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Filter"
         Height          =   3855
         Left            =   -64560
         TabIndex        =   118
         Top             =   3960
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CommandButton Command2 
            Caption         =   "Show"
            Height          =   375
            Left            =   120
            TabIndex        =   121
            Top             =   1200
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CheckBox ChkAppointment 
            Caption         =   "New Appointments"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   360
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox ChkRevisit 
            Caption         =   "Scheduled Re-Visits"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   840
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Next of Kin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   10080
         TabIndex        =   46
         Top             =   360
         Width           =   3015
         Begin VB.TextBox TxtKINAddress 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   2208
            Width           =   2775
         End
         Begin VB.TextBox TxtKinRelationship 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   1632
            Width           =   2775
         End
         Begin VB.TextBox TxtKinFirstName 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox TxtKinTel 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   19
            Top             =   2784
            Width           =   2775
         End
         Begin VB.TextBox TxtKinSecondName 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   1056
            Width           =   2775
         End
         Begin VB.TextBox TxtKinEmail 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   3360
            Width           =   2775
         End
         Begin VB.Label Label33 
            Caption         =   "Physical Address"
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
            TabIndex        =   132
            Top             =   1950
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Relationship"
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
            TabIndex        =   112
            Top             =   1365
            Width           =   1095
         End
         Begin VB.Label Label8 
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
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label9 
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
            Left            =   120
            TabIndex        =   49
            Top             =   800
            Width           =   1815
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Tel/Mobile"
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
            TabIndex        =   48
            Top             =   2570
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Email Address"
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
            TabIndex        =   47
            Top             =   3120
            Width           =   2175
         End
      End
      Begin VB.OptionButton OptExistingClient 
         Caption         =   "Existing Client Appointment"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   -69480
         TabIndex        =   88
         Top             =   480
         Width           =   2895
      End
      Begin VB.OptionButton OptNewClient 
         Caption         =   "New Client Appointment"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   -72840
         TabIndex        =   87
         Top             =   480
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.CommandButton CmdBookSearch 
         Caption         =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -63240
         TabIndex        =   86
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBookCardNumber 
         Height          =   285
         Left            =   -64920
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame7 
         Caption         =   "Appointment Details"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   69
         Top             =   720
         Width           =   12975
         Begin VB.ComboBox CboEndAMPM 
            Height          =   315
            ItemData        =   "FrmPatients.frx":047A
            Left            =   8040
            List            =   "FrmPatients.frx":0484
            TabIndex        =   134
            Text            =   "AM"
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox CboStartAMPM 
            Height          =   315
            ItemData        =   "FrmPatients.frx":0490
            Left            =   4680
            List            =   "FrmPatients.frx":049A
            TabIndex        =   133
            Text            =   "AM"
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox ChkSearch 
            Caption         =   "Check1"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox CboDoctor 
            Height          =   315
            Left            =   9240
            TabIndex        =   36
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox TxtBookID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   12480
            TabIndex        =   82
            Text            =   "0"
            Top             =   1680
            Width           =   270
         End
         Begin MSComCtl2.DTPicker DTStartTime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "h:nn AM/PM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   315
            Left            =   3000
            TabIndex        =   35
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            OLEDropMode     =   1
            CheckBox        =   -1  'True
            CustomFormat    =   "12:00 AM"
            Format          =   70385666
            UpDown          =   -1  'True
            CurrentDate     =   41305
         End
         Begin VB.TextBox TxtBookSecondName 
            Height          =   315
            Left            =   8280
            TabIndex        =   31
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox TxtBookSurname 
            Height          =   315
            Left            =   2160
            TabIndex        =   32
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox TxtBookDoctor 
            Height          =   315
            Left            =   9960
            TabIndex        =   74
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox TxtBookTelephone 
            Height          =   315
            Left            =   8280
            TabIndex        =   33
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox TxtBookFirstName 
            Height          =   315
            Left            =   2160
            TabIndex        =   30
            Top             =   360
            Width           =   3375
         End
         Begin MSComCtl2.DTPicker DtAppointment 
            Height          =   315
            Left            =   480
            TabIndex        =   34
            Top             =   1680
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   70451201
            CurrentDate     =   40746
         End
         Begin MSComCtl2.DTPicker DTEndTime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "h:nn AM/PM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   315
            Left            =   6360
            TabIndex        =   130
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   70451202
            UpDown          =   -1  'True
            CurrentDate     =   40746
         End
         Begin VB.Label Label37 
            Caption         =   "Appointment End Time"
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
            TabIndex        =   131
            Top             =   1395
            Width           =   1815
         End
         Begin VB.Label Label19 
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
            Left            =   1080
            TabIndex        =   77
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Telephone Number"
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
            TabIndex        =   76
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label24 
            Caption         =   "Doctor To be Seen"
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
            Left            =   9240
            TabIndex        =   75
            Top             =   1425
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Appointment Start Time"
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
            Left            =   3000
            TabIndex        =   73
            Top             =   1395
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "Appointment Date"
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
            TabIndex        =   72
            Top             =   1395
            Width           =   1695
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
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
            Left            =   6840
            TabIndex        =   71
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Surname"
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
            Left            =   1320
            TabIndex        =   70
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3855
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   9855
         Begin VB.TextBox TxtFileNumber 
            Height          =   285
            Left            =   6960
            TabIndex        =   138
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox CboDoctors 
            Height          =   315
            Left            =   6960
            TabIndex        =   136
            Top             =   3360
            Width           =   2775
         End
         Begin VB.TextBox TxtPOBox 
            Height          =   315
            Left            =   1680
            TabIndex        =   10
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox TxtEmail 
            Height          =   285
            Left            =   4080
            TabIndex        =   14
            Top             =   3360
            Width           =   2655
         End
         Begin VB.TextBox TxtOfficeLine 
            Height          =   285
            Left            =   1560
            TabIndex        =   13
            Top             =   3360
            Width           =   1575
         End
         Begin VB.TextBox TxtTown 
            Height          =   285
            Left            =   6960
            TabIndex        =   12
            Top             =   2760
            Width           =   2775
         End
         Begin VB.TextBox TxtPoBoxCode 
            Height          =   315
            Left            =   4680
            TabIndex        =   11
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox TxtInsuranceNumber 
            Height          =   285
            Left            =   4080
            TabIndex        =   110
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox TxtSurname 
            Height          =   315
            Left            =   1440
            TabIndex        =   2
            Top             =   1320
            Width           =   4095
         End
         Begin VB.ComboBox CboBillingCompany 
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox TxtCardNumber 
            Enabled         =   0   'False
            Height          =   285
            Left            =   8400
            MaxLength       =   15
            TabIndex        =   54
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox TxtFirstName 
            Height          =   315
            Left            =   1440
            TabIndex        =   0
            Top             =   840
            Width           =   4095
         End
         Begin VB.ComboBox CboGender 
            Height          =   315
            Left            =   1440
            TabIndex        =   7
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox TxtAge 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3840
            TabIndex        =   5
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox TxtSecondName 
            Height          =   285
            Left            =   6960
            TabIndex        =   1
            Top             =   855
            Width           =   2775
         End
         Begin VB.TextBox TxtTel 
            Height          =   285
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   3
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox TxtAddress 
            Height          =   285
            Left            =   6960
            TabIndex        =   6
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox TxtID 
            Height          =   315
            Left            =   3840
            MaxLength       =   8
            TabIndex        =   8
            Top             =   2280
            Width           =   1695
         End
         Begin VB.CommandButton CmdBillingcompanies 
            Caption         =   "..."
            Height          =   315
            Left            =   3600
            TabIndex        =   53
            Top             =   480
            Width           =   375
         End
         Begin VB.ComboBox CboPaymentMode 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmPatients.frx":04A6
            Left            =   6960
            List            =   "FrmPatients.frx":04A8
            TabIndex        =   9
            Top             =   2280
            Width           =   2775
         End
         Begin VB.CommandButton CmdCardSearch 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9360
            TabIndex        =   52
            Top             =   480
            Width           =   375
         End
         Begin MSComCtl2.DTPicker DTDateofBirth 
            Height          =   315
            Left            =   1440
            TabIndex        =   4
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   70123521
            CurrentDate     =   39414
         End
         Begin VB.Label Label38 
            Caption         =   "File No"
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
            Left            =   6960
            TabIndex        =   139
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Doctor to be Seen"
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
            Left            =   6960
            TabIndex        =   137
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "Email Address"
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
            Left            =   3360
            TabIndex        =   117
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label34 
            Caption         =   "Office Line"
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
            TabIndex        =   116
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Town"
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
            Left            =   5640
            TabIndex        =   115
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label31 
            Caption         =   "P.O Box Code"
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
            TabIndex        =   114
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label30 
            Caption         =   "Post Office Box"
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
            TabIndex        =   113
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label29 
            Caption         =   "Insurance Card Number"
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
            TabIndex        =   109
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Surname"
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
            TabIndex        =   83
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   " Card No"
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
            Left            =   8280
            TabIndex        =   66
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "AGE"
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
            Left            =   3240
            TabIndex        =   65
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
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
            Left            =   5640
            TabIndex        =   64
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Billing Company"
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
            TabIndex        =   63
            Top             =   230
            Width           =   2415
         End
         Begin VB.Label Label4 
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
            Left            =   240
            TabIndex        =   62
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Date of Birth"
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
            TabIndex        =   61
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Gender"
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
            TabIndex        =   60
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Tel\Mobile"
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
            Left            =   5640
            TabIndex        =   59
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Address"
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
            Left            =   5640
            TabIndex        =   58
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "ID N&o"
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
            Left            =   3240
            TabIndex        =   57
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Payment Mode"
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
            Left            =   5595
            TabIndex        =   56
            Top             =   2280
            Width           =   1335
         End
      End
      Begin TabDlg.SSTab PatientTab2 
         Height          =   3495
         Left            =   120
         TabIndex        =   41
         Top             =   4320
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6165
         _Version        =   393216
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Garamond"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Patient Database Listing"
         TabPicture(0)   =   "FrmPatients.frx":04AA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Patients Resent to Consultation"
         TabPicture(1)   =   "FrmPatients.frx":04C6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Direct Lab Request"
         TabPicture(2)   =   "FrmPatients.frx":04E2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame10"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame10 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   79
            Top             =   360
            Width           =   9735
            Begin VB.CheckBox ChkEnable 
               Caption         =   "Enable for Typing"
               Height          =   255
               Left            =   8040
               TabIndex        =   81
               Top             =   0
               Width           =   1575
            End
            Begin VB.TextBox TxtLabRequest 
               Height          =   2655
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   9495
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "List in Waiting Room"
            Height          =   3015
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   9735
            Begin VSFlex6DAOCtl.vsFlexGrid Grid 
               Height          =   2655
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   4683
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
         Begin VB.Frame Frame5 
            Caption         =   "List in Waiting Room"
            Height          =   3015
            Left            =   -74880
            TabIndex        =   42
            Top             =   360
            Width           =   9735
            Begin VSFlex6DAOCtl.vsFlexGrid GridConsult 
               Height          =   2655
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   4683
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
      Begin TabDlg.SSTab TabAppointments 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   90
         Top             =   2880
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   8705
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Pending Appointments"
         TabPicture(0)   =   "FrmPatients.frx":04FE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame9"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Missed Appointments"
         TabPicture(1)   =   "FrmPatients.frx":051A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame12"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame11"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Honored Appointments"
         TabPicture(2)   =   "FrmPatients.frx":0536
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame14"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame13"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         Begin VB.Frame Frame14 
            Height          =   735
            Left            =   -74880
            TabIndex        =   103
            Top             =   4080
            Width           =   9735
            Begin VB.TextBox TxtHonoredApp 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   7200
               TabIndex        =   106
               Text            =   "0"
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label Label28 
               Caption         =   "Number of Honored Appointments"
               Height          =   255
               Left            =   7200
               TabIndex        =   107
               Top             =   0
               Width           =   2415
            End
         End
         Begin VB.Frame Frame13 
            Height          =   3735
            Left            =   -74880
            TabIndex        =   101
            Top             =   360
            Width           =   9975
            Begin VSFlex6DAOCtl.vsFlexGrid GridHonored 
               Height          =   3375
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   9735
               _ExtentX        =   17171
               _ExtentY        =   5953
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
            Height          =   735
            Left            =   -74880
            TabIndex        =   99
            Top             =   4080
            Width           =   9975
            Begin VB.CommandButton CmdDeleteMissed 
               Caption         =   "Delete Missed Appointment"
               Height          =   375
               Left            =   120
               TabIndex        =   125
               Top             =   240
               Width           =   2295
            End
            Begin VB.TextBox TxtMissedApp 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   3720
               TabIndex        =   104
               Text            =   "0"
               Top             =   240
               Width           =   2415
            End
            Begin VB.CommandButton CmdBookReshedule 
               Caption         =   "Reschedule Visit"
               Enabled         =   0   'False
               Height          =   375
               Left            =   7560
               TabIndex        =   100
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label Label27 
               Caption         =   "Number of Missed Appointments"
               Height          =   255
               Left            =   3720
               TabIndex        =   105
               Top             =   0
               Width           =   2415
            End
         End
         Begin VB.Frame Frame8 
            Height          =   3735
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   9975
            Begin VSFlex6DAOCtl.vsFlexGrid GridAppointments 
               Height          =   3375
               Left            =   120
               TabIndex        =   98
               Top             =   240
               Width           =   9735
               _ExtentX        =   17171
               _ExtentY        =   5953
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
         Begin VB.Frame Frame11 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   95
            Top             =   360
            Width           =   9975
            Begin VSFlex6DAOCtl.vsFlexGrid GridMissedApp 
               Height          =   3255
               Left            =   120
               TabIndex        =   96
               Top             =   240
               Width           =   9735
               _ExtentX        =   17171
               _ExtentY        =   5741
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
         Begin VB.Frame Frame9 
            Height          =   735
            Left            =   120
            TabIndex        =   91
            Top             =   4080
            Width           =   9735
            Begin VB.CommandButton CmdConvertBooking 
               Caption         =   "Convert Booking to a Visit"
               Height          =   375
               Left            =   7200
               TabIndex        =   93
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox TxtPendingApp 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   120
               TabIndex        =   92
               Text            =   "0"
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label Label26 
               Caption         =   "Number of Pending Appointments"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   0
               Width           =   2415
            End
         End
      End
      Begin VB.CheckBox ChkPreserve 
         Caption         =   "Preserve New Data for Posting"
         Height          =   255
         Left            =   7320
         TabIndex        =   84
         Top             =   3795
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   4560
         TabIndex        =   67
         Top             =   3765
         Width           =   495
      End
      Begin VB.CheckBox ChkConsultation 
         Caption         =   "Charge Consultation Fee"
         Height          =   195
         Left            =   2400
         TabIndex        =   68
         Top             =   3825
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label Label25 
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
         Left            =   -66240
         TabIndex        =   89
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Post Patient"
      Height          =   855
      Left            =   120
      TabIndex        =   29
      Top             =   8160
      Width           =   13215
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
         Index           =   5
         Left            =   10080
         TabIndex        =   78
         Top             =   280
         Width           =   1215
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
         Left            =   11400
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptObservation 
         Caption         =   "TO OBSERVATION"
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
         Left            =   120
         TabIndex        =   39
         Top             =   280
         Value           =   -1  'True
         Width           =   2655
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
         Left            =   3000
         TabIndex        =   21
         Top             =   280
         Width           =   2175
      End
      Begin VB.OptionButton OptCashier 
         Caption         =   "TO CASHIERS"
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
         TabIndex        =   38
         Top             =   280
         Width           =   2295
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
         Left            =   7710
         TabIndex        =   37
         Top             =   280
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      Height          =   8895
      Left            =   13440
      TabIndex        =   23
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear Text"
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
         TabIndex        =   111
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
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
         TabIndex        =   28
         Top             =   8280
         Width           =   1215
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
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
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
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
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
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
         TabIndex        =   26
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
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
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDog 
         Left            =   600
         Top             =   6600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Menu MnuOtherAppointments 
      Caption         =   "Non Patient Appointments"
      Begin VB.Menu MnuOfficial 
         Caption         =   "Official Appointments"
      End
   End
End
Attribute VB_Name = "FrmPatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RsAdd As New ADODB.Recordset
Dim RsGrid As New ADODB.Recordset
Dim RsRecords As New ADODB.Recordset
Dim RsNextOfKin As New ADODB.Recordset
Dim RsImageStore As New ADODB.Recordset

Dim ItemSelected As Integer
Dim Adding As Boolean
Dim Editing As Boolean
Dim lvSaving As Boolean

Dim FromAppointments As Boolean
Dim lvDoNotSearch As Boolean

Dim lvImagepath As String

Public Function CheckDateTimeAvailability(AppDate As Date, AppStartTime As String, AppEndTime As String) As Boolean
    Dim lvLastAllocatedTime As String
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT APPENDTIME as EndTime FROM APPOINTMENTS WHERE APPDATE = '" & AppDate & "' ORDER BY APPENDTIME DESC", Conn, adOpenDynamic, adLockOptimistic
        If RsRecords.EOF = False Then
            lvLastAllocatedTime = RsRecords!EndTime
        End If
    RsRecords.Close
    
''    'MAKE 24 HOUR CLOCK FOR ACCURATE COMPARISON
''    If Right(AppStartTime, 2) = "PM" Then AppStartTime = AppStartTime + 12
''    If Right(AppEndTime, 2) = "PM" Then AppEndTime = Right(AppEndTime.Value, 11) + 12
''        adDBTimeStamp
''        adDBTime
        
    If Mid(AppStartTime, 10, 8) < Left(lvLastAllocatedTime, 8) Then
        MsgBox "Time already allocated to another appointment", vbExclamation
        CheckDateTimeAvailability = False
        Exit Function
    End If
        CheckDateTimeAvailability = True
End Function

Private Sub FillNextOfKin()
    If RsNextOfKin.State = 1 Then Set RsNextOfKin = Nothing
    RsNextOfKin.Open "SELECT * FROM NEXT_OF_KIN where cardnumber = '" & TxtCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsNextOfKin
            If .EOF = False Then
                TxtCardNumber = !CardNumber
                TxtKinFirstName = !FirstName
                TxtKinSecondName = !SECONDNAME
                TxtKinRelationship = !Relationship
                TxtKINAddress = !ADDRESS
                TxtKinTel = !TELEPHONE
                If TxtKinEmail = "" Then TxtKinEmail = O Else TxtKinEmail = !email
            End If
        End With
    RsNextOfKin.Close
End Sub

Private Sub RetrievePatientPicture(PicCardNumber As String)
On Error GoTo Errorhandler
    TxtLabTest = ""
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM PATIENT_PICTURES WHERE CARDNUMBER = '" & PicCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsImageStore
            If .EOF = False Then
                LoadPictureFromDB RsImageStore, "PICTURE", ImgPreview, D
            End If
        End With
    RsImageStore.Close
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub SavePatientPicture()
On Error GoTo Errorhandler
    If lvImagepath = "" Then Exit Sub
    If RsImageStore.State = 1 Then Set RsImageStore = Nothing
    RsImageStore.Open "SELECT * FROM PATIENT_PICTURES", Conn, adOpenStatic, adLockOptimistic
    
   'IF IMAGE ALREADY SAVED, THEN THE ASSUMPTION IS THAT ITS BEING REPLACED.
   If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT CARDNUMBER FROM PATIENT_PICTURES WHERE CARDNUMBER = '" & TxtCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
        If RsRecords.EOF = False Then
            Resp = MsgBox("A Picture has already been saved for this Patient. Do you wish to Replace it?", vbQuestion + vbYesNo)
                If Resp = vbYes Then
                    GoTo Updating
                Else
                    Exit Sub
                End If
        End If
        
        RsImageStore.AddNew
Updating:
            RsImageStore!CardNumber = TxtCardNumber
            LoadImageFromFileToDB lvImagepath, RsImageStore, "PICTURE", FileLen(lvImagepath)
            'RsImageStore!PROCESSNUMBER = "001"
        RsImageStore.Update
        
        'MsgBox "Conversion and Saving Completed Succesfully", vbInformation
        RsRecords.Close
        'POPULATEPreScan
        'POPULATEPostScan
        'ClearText FrmConverter
    Exit Sub
Errorhandler:
   MsgBox Err.Number & " " & Err.Description
   Exit Sub
   Resume
End Sub

Private Sub Check1_Click()
    TxtLabRequest.Enabled = True
End Sub

Private Sub ChkDateFilter_Click()
    If ChkDateFilter.Value = 1 Then
        DtDateFilter.Enabled = True
        CmdDateFilter.Enabled = True
    Else
        DtDateFilter.Enabled = False
        CmdDateFilter.Enabled = False
        FillPendingAppointment
    End If
End Sub

Private Sub ChkEnable_Click()
    If ChkEnable.Value = 1 Then
        TxtLabRequest.Enabled = True
    Else
        TxtLabRequest.Enabled = False
    End If
End Sub

Private Sub CmdBook_Click()
    MsgBox "Are you sure", vbQuestion
    FromAppointments = True
    CmdNew_Click
    ConvertAppointmentToVisit (TxtID)
End Sub
Public Function ConvertAppointmentToVisit(ByVal AppointmentID)
    TxtFirstname = TxtBookFirstName
    TxtSecondName = TxtBookSecondName
    TxtSurname = TxtBookSurname
    TxtTel = TxtBookTelephone
    PatientsTab.Tab = 0
    If Grid.Rows >= 2 Then
        Grid_DblClick
    End If
End Function

Private Sub CmdBookReshedule_Click()
    'LOOP TO SEE IF ANY RECORD IS TICKED
        For i = 1 To GridMissedApp.Rows - 1
            If GridMissedApp.TextMatrix(i, 0) <> 0 Then
                Conn.Execute "UPDATE APPOINTMENTS SET APPDATE = '" & DtAppointment & "' WHERE APPOINTMENTID = '" & GridMissedApp.TextMatrix(GridMissedApp.Row, 8) & "'"
                FillMissedAppointments
                FillPendingAppointment
                ResetMode
                Exit Sub
            Else
            
            End If
        Next
        MsgBox "Please Select a Record by Ticking before Rescheduling Visit", vbInformation
End Sub

Private Sub CmdBookSearch_Click()
    GlbCalledFromAppointments = True
    Load FrmHistory
    FrmHistory.Show 1
    'populate other fields. by reusing the double click code
    If TxtBookCardNumber = "" Then Exit Sub
        loadDetailsToScreen TxtBookCardNumber
End Sub
Private Sub loadDetailsToScreen(ByRef CardNo As String)
On Error GoTo Errorhandler
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT FIRSTNAME,SECONDNAME,SURNAME,Telephone FROM PATIENT_DETAILS WHERE CARDNUMBER = '" & TxtBookCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsRecords
            TxtBookFirstName = !FirstName
            TxtBookSecondName = !SECONDNAME
            TxtBookSurname = !SURNAME
            TxtBookTelephone = !TELEPHONE
            DtAppointment = GlbSysDate
        End With
    RsRecords.Close
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdBrowse_Click()
On Error GoTo Errorhandler
    CommonDog.ShowOpen
    StrImagePath = CommonDog.FileName
   'Picture1. = StrImagePath
   ImgPreview.Container = StrImagePath
   ImgPreview.Picture = LoadPicture(StrImagePath)
   lvImagepath = StrImagePath
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdCardSearch_Click()
On Error GoTo Errorhandler
    GlbCalledFromPatients = True
    Load FrmHistory
    FrmHistory.Show 1
    'populate other fields. by reusing the double click code
    If TxtCardNumber = "" Then Exit Sub
    StrPharmCardNumber = TxtCardNumber
        Grid_DblClick
Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub CmdClear_Click()
        ResetMode
        ClearControls
       ImgPreview.Picture = LoadPicture(App.Path & "\NoImage.bmp")
End Sub

Private Sub CmdConvertBooking_Click()
On Error GoTo Errorhandler
Dim Resp
    Resp = MsgBox("Are you Sure you Wish to Covert Booking to Visit?", vbQuestion + vbYesNo)
    If Resp = vbNo Then Exit Sub
    FromAppointments = True
    'UPDATE STATUS ON APPOINTMENTS LIST
    Conn.Execute "UPDATE APPOINTMENTS SET STATUS = '1' WHERE APPOINTMENTID = '" & TxtBookID & "'"
    PatientsTab.Tab = 0
    CmdNew_Click
    ConvertAppointmentToVisit (TxtBookID)
    FillPendingAppointment
    FillHonoredAppointments
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub CmdDateFilter_Click()
    On Error GoTo Errorhandler
'''    Dim Rcount As Integer
'''    GridAppointments.Clear
'''    GridAppointments.Rows = 1
'''    GridAppointments.Cols = 3
'''    GridAppointments.ColAlignment(1) = flexAlignCenterCenter
'''    GridAppointments.ColWidth(1) = 3105
'''    GridAppointments.ColWidth(2) = 3990
'''    'GridAppointments.FormatString = "FIRST NAME |   SECOND NAME  |   SURNAME     |  PHONE NUMBER |APPOINTMENT DATE|     TIME  |  DOCTOR TO BE SEEN  | ID | CARD NO"
'''    GridAppointments.FormatString = "FIRST NAME |   SECOND NAME  |   SURNAME     |  PHONE NUMBER |APPOINTMENT DATE|    START TIME  | END TIME  |  DOCTOR TO BE SEEN  | ID | CARD NO"
'''        If RsGrid.State = adStateOpen Then RsGrid.Close
'''        RsGrid.Open "SELECT * FROM APPOINTMENTS WHERE APPDATE = '" & Format(DtDateFilter, "DD MMM YYYY") & "' ORDER BY APPDATE DESC", Conn, adOpenStatic, adLockOptimistic
'''            If RsGrid.RecordCount <> 0 Then
'''                With RsGrid
'''                    While Not .EOF
'''                        'GridAppointments.AddItem !FirstName & vbTab & !SECONDNAME & vbTab & !SURNAME & vbTab & !TELEPHONE & vbTab & !AppDate & vbTab & Right(!APPTIME, 11) & vbTab & !DOCTOR & vbTab & !AppointmentID & vbTab & !cardnumber
'''                        GridAppointments.AddItem !FirstName & vbTab & !SECONDNAME & vbTab & !SURNAME & vbTab & !TELEPHONE & vbTab & !AppDate & vbTab & Right(!AppStartTime, 11) + " " + !STARTAMPM & vbTab & Right(!AppEndTime, 11) + " " + !Endampm & vbTab & !DOCTOR & vbTab & !AppointmentID & vbTab & !CardNumber
'''                        .MoveNext
'''                        Rcount = Rcount + 1
'''                    Wend
'''                End With
'''            End If
'''        RsGrid.Close
'''        TxtPendingApp = Rcount

        FillPendingAppointment
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub CmdDelete_Click()
On Error GoTo Errorhandler
   Select Case PatientsTab.Tab
        Case 0
            Select Case PatientTab2.Tab

                Case 1
                    MsgBox "Commented for Now", vbExclamation: Exit Sub
                    'AuditTrail GlbCurrentUser, EnumConsultation, GlbSysDate, Time, "Deleted a returned Visit for - " & GridConsult.TextMatrix(GridConsult.Row, 2) & ""
                    'Conn.Execute "DELETE FROM COMPLAINS WHERE CARDNUMBER = '" & GridConsult.TextMatrix(GridConsult.Row, 0) & "' AND VISITNUMBER = '" & GridConsult.TextMatrix(GridConsult.Row, 1) & "'"
                    FillPostedToConsultant
                    Exit Sub
                Case 2
                    
            End Select
                    
            Resp = MsgBox("Are you sure you wish to delete Card Number " & TxtCardNumber & " ?", vbQuestion + vbYesNo)
            If Resp = vbNo Then Exit Sub
            Conn.Execute "DELETE FROM PATIENT_DETAILS WHERE CARDNUMBER = '" & TxtCardNumber & "'"
            Conn.Execute "DELETE FROM NEXT_OF_KIN WHERE CARDNUMBER = '" & TxtCardNumber & "'"
            AuditTrail GlbCurrentUser, EnumConsultation, GlbSysDate, Time, "Deleted Patient Details For CardNumber -" + " " & TxtCardNumber & "  " + "-" + "  " & TxtFirstname & " " + " " & TxtSecondName & " " + " " & TxtSurname & ""
            FillPatientGrid
        Case 1
            If TxtBookID = "" Then Exit Sub
            Resp = MsgBox("Are you sure you wish to delete Appointment for " & TxtBookFirstName & " " & TxtBookSecondName & "?", vbQuestion + vbYesNo)
            If Resp = vbNo Then Exit Sub
            Conn.Execute "DELETE FROM APPOINTMENTS WHERE APPOINTMENTID = '" & TxtBookID & "'"
            AuditTrail GlbCurrentUser, EnumConsultation, GlbSysDate, Time, "Deleted Appointment for - " + " " & TxtBookFirstName & " " + " " & TxtBookSecondName & " " + " " & TxtBookSurname & " "
            FillPendingAppointment
            FillHonoredAppointments
            FillMissedAppointments
    End Select
    MsgBox "Record Deleted Succesfully", vbInformation
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
    'Resume
End Sub

Private Sub CmdDeleteMissed_Click()
    Conn.Execute "DELETE FROM APPOINTMENTS WHERE APPOINTMENTID = '" & GridMissedApp.TextMatrix(GridMissedApp.Row, 8) & "'"
    FillMissedAppointments
End Sub
Private Sub CmdEdit_Click()
    EditMode
    EnableControls
    Select Case PatientsTab.Tab
        Case 0
        
        Case 1
            CmdBookReshedule.Enabled = True
            Select Case TabAppointments.Tab
                Case 0
                
                Case 1
                    
                Case 2
            
            End Select
    End Select
End Sub

Private Sub CmdExit_Click()
    If CmdExit.Caption = "Cancel" Then
        ResetMode
        ClearControls
        If PatientsTab.Tab = 1 Then FillPendingAppointment: FillHonoredAppointments: FillMissedAppointments
        DisableControls
        
        OptNewClient.Enabled = False
        OptExistingClient.Enabled = False
        CmdBookSearch.Enabled = False
        CmdBookReshedule.Enabled = False
        CmdCardSearch.Enabled = True
    Else
        
        Unload Me
    End If
End Sub

Private Sub CmdExplode_Click()
   RetrievePatientPicture TxtCardNumber
    FrmPicture.Show 1
End Sub

Private Sub CmdNew_Click()
On Error GoTo Errorhandler
    AddMode
    EnableControls
    TxtAge.Enabled = False
    If FromAppointments = False Then ClearControls
    If PatientsTab.Tab = 1 Then FillPendingAppointment: FillHonoredAppointments: FillMissedAppointments
    TxtCardNumber = AllocateRunningNumber + "/" + Right(GlbSysDate, 2)
    CboPaymentMode.Text = "1 - CASH"
    TxtCardNumber.Enabled = False
    FromAppointments = False
    OptExistingClient.Enabled = True
    OptNewClient.Enabled = True
    CmdCardSearch.Enabled = False
    TxtFileNumber = "NP"
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
End Sub
Private Function AllocateRunningNumber()
    AllocateRunningNumber = Trim(FindRecord("GENERALPARAMS", "ITEMVALUE", "ITEMNAME = 'RUNNUMBER'"))
End Function
Private Sub CmdPost_Click()
On Error GoTo Errorhandler
KARI = GlbSysDate
    
    If PatientTab2.Tab = 1 Then
        If GridConsult.Row <= 1 Then Exit Sub
           Resp = MsgBox("Are you sure you wish to send this record back to be processed?", vbInformation + vbYesNo + vbDefaultButton2)
                If Resp = vbYes Then
                    GlbCardNumber = GridConsult.TextMatrix(GridConsult.Row, 0)
                    GoTo Send_Only
                End If
            Exit Sub
    End If
    
    If TxtCardNumber = "" Then MsgBox "Please Input Card Number before Posting Record", vbExclamation: Exit Sub
    If TxtFirstname = "" Then MsgBox "Please fill the First Name before Posting Record", vbExclamation: Exit Sub
    If TxtAge = "" Then MsgBox "Please Select the Date of Birth before Posting Record", vbExclamation: Exit Sub
    
    If ItemSelected = 0 Then MsgBox "Please select where to send the patient.", vbInformation: Exit Sub: Exit Sub
        
    If CboPaymentMode.Text = "" Then MsgBox "Please Select Payment Mode before Posting Patient for Treatment", vbInformation: Exit Sub
    If CboDoctors.Text = "" Then MsgBox "Please Select Doctor to be seen before Posting Patient for Treatment", vbInformation: Exit Sub
    
    'CHECK IF THIS PATIENT HAS ALREADY BEEN SENT TO OBSERVATION TO AVOID DUPLICATION. RECEPTIONISTS DO THAT ALOT.
    Dim RsCheck As New ADODB.Recordset
        RsCheck.Open "SELECT * FROM COMPLAINS WHERE CARDNUMBER = '" & TxtCardNumber & "' AND VISITDATE = '" & Format(KARI, "DD MMM YYYY") & "' AND TOOBSERVATION = '1'", Conn, adOpenStatic, adLockOptimistic
            If RsCheck.EOF = False Then
                MsgBox "" & TxtFirstname & " " + TxtSecondName + " " + TxtSurname & " with Card Number " & TxtCardNumber & " has already been posted to Observation", vbExclamation
                Exit Sub
            End If
    'Conn.Execute "INSERT INTO COMPLAINS (CARDNUMBER,BP,WEIGHT,HEIGHT,VISITDATE,COMPLAINS,DIAGNOSIS,PRESCRIPTION,ADMISSION,REFERRAL,NURSE,DOCTOR,OBSERVED)VALUES " & _
                 "('" & Grid.TextMatrix(1, 0) & "','','','', '" & Format(KARI, "DDMMMYYYY") & "','','','','','','','','0')"
                 
            'INSERT ASSIGNED PROCEEDURE DURING LAST VISIT.
            lvProceedure = FindRecord("COMPLAINS", "NEXTPROCEEDURE", "CARDNUMBER = '" & TxtCardNumber & "' ORDER BY VISITNUMBER DESC")
    If PatientTab2.Tab = 0 And ItemSelected <> 5 Then
        Conn.Execute "INSERT INTO COMPLAINS (CARDNUMBER,VISITDATE,BILLINGCOMPANY,TimeIn,DOCTOR,NEXTPROCEEDURE)VALUES ('" & TxtCardNumber & "', '" & Format(KARI, "DDMMMYYYY") & "','" & GetID_NameFromCombo(CboBillingCompany, 1) & "','" & Time & "','" & CboDoctors & "','" & lvProceedure & "')"
        GlbCardNumber = TxtCardNumber
        VARGRID = GridConsult
    ElseIf ItemSelected = 5 Then
        If TxtLabRequest = "" Then MsgBox "Please enter the lab request before sending to LAB.", vbInformation: Exit Sub
        Conn.Execute "INSERT INTO COMPLAINS (CARDNUMBER,VISITDATE,BILLINGCOMPANY,LABREQUEST,DOCTOR)VALUES ('" & TxtCardNumber & "', '" & Format(KARI, "DDMMMYYYY") & "','" & Mid(Grid.TextMatrix(Grid.Row, 2), 1, 3) & "','" & TxtLabRequest & "','FROM RECEPTION')"
    End If
    
Send_Only:

    Select Case ItemSelected
   Case 1
        'To Obsservation
         SendCardnumber = GlbCardNumber
         DUMMY = SendPatient(EnumObservation, SendCardnumber, KARI)
         'FrmObservation.Show
         'Unload Me
    Case 2
        'To Doctor
        SendCardnumber = GlbCardNumber
        DUMMY = SendPatient(EnumDoctors, SendCardnumber, KARI)
            'FrmWaitingRoom.Show
            'Unload Me
    Case 3
        'ToCashiers
        SendCardnumber = GlbCardNumber
        DUMMY = SendPatient(EnumCashier, SendCardnumber, KARI)
        'FrmCashier.Show
        'Unload Me
    Case 4
        'To Pharmacy
        SendCardnumber = GlbCardNumber
        DUMMY = SendPatient(EnumPharmacy, SendCardnumber, KARI)
        'FrmPharmacy.Show
        'Unload Me
    Case 5
        'To Lab
        SendCardnumber = GlbCardNumber
        DUMMY = SendPatient(EnumLab, SendCardnumber, KARI)
    End Select
    
    'POST CONSULTATION FEES IF CHECKED.
        If ChkConsultation.Value = 1 Then
            VISITNUMBER = FindRecord("COMPLAINS", "VISITNUMBER", "CARDNUMBER = '" & SendCardnumber & "'  ORDER BY VISITNUMBER DESC")
            ConsultationFee = FindRecord("COMPLAINS", "CARDNUMBER", "CARDNUMBER = '" & TxtCardNumber & "'")
            If ConsultationFee = "" Then
                  StrVariable = Mid(CboBillingCompany, 1, 3)
                  ConsultationFee = FindRecord("SERVICE_PROVIDER", "CONSULTATIONFEE", "COMPANYCODE = '" & StrVariable & "'")
                  'CHECK IF ITS A REVISIT AND CHARGE DIFFERENT CONSULTATION FEE
            Else
                    ConsultationFee = 1700
            End If
            Select Case Mid(CboPaymentMode, 1, 1)
                Case "1"
                    Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,VISITDATE,CODE,DESCRIPTION,PAYMENTMODE,CASHAMOUNT,PAYDATE)VALUES ('" & SendCardnumber & "','" & VISITNUMBER & "','" & Format(KARI, "DDMMMYYYY") & "','001','Consultation Fee','" & Mid(CboPaymentMode, 1, InStr(CboPaymentMode, "-") - 1) & "','" & ConsultationFee & "','" & Format(KARI, "DDMMMYYYY") & "')"
                Case Else
                    Conn.Execute "INSERT INTO PRESCRIPTION (CARDNUMBER,VISITNUMBER,VISITDATE,CODE,DESCRIPTION,PAYMENTMODE,CREDITAMOUNT)VALUES ('" & SendCardnumber & "','" & VISITNUMBER & "','" & Format(KARI, "DDMMMYYYY") & "','001','Consultation Fee','" & Mid(CboPaymentMode, 1, InStr(CboPaymentMode, "-") - 1) & "','" & ConsultationFee & "')"
            End Select
        End If
    ResetMode
    FillPostedToConsultant
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
   Exit Sub
   Resume
End Sub

Private Sub CmdRefresh_Click()
    FillPendingAppointment
End Sub

Private Sub CmdSave_Click()
    On Error GoTo Errorhandler
    lvSaving = True
   Select Case PatientsTab.Tab
        Case 0
            'VALIDATE FIELDS BEFORE SAVING
            If TxtCardNumber = "" Then MsgBox "Please fill in the Card Number before Saving Patient Data", vbInformation: Exit Sub
            If CboBillingCompany = "" Then MsgBox "Please Select the Billing Company before Saving Patient Data", vbInformation: Exit Sub
            If GetID_NameFromCombo(CboBillingCompany, 1) <> "001" And TxtInsuranceNumber = "" Then MsgBox "Please enter the Insurance Card Number", vbInformation: Exit Sub
            If TxtSurname = "" Then MsgBox "Please fill in the SurName before Saving Patient Data", vbInformation: Exit Sub
            If TxtFirstname = "" Then MsgBox "Please fill in the First Name before Saving Patient Data", vbInformation: Exit Sub
            'If TxtSecondName = "" Then MsgBox "Please fill in the Second Name before Saving Patient Data", vbInformation: Exit Sub
            If DTDateofBirth = "" Then MsgBox "Please Select the Date of Birth before Saving Patient Data", vbInformation: Exit Sub
            'If TxtID = "" Then MsgBox "Please fill in the Identification Number before Saving Patient Data", vbInformation: Exit Sub
            If TxtAge = "" Then MsgBox "Please fill in the Identification Number before Saving Patient Data", vbInformation: Exit Sub
            If CboGender = "" Then MsgBox "Please Select the Gender before Saving Patient Data", vbInformation: Exit Sub
            If TxtTel = "" Then MsgBox "Please fill in the Telephone Number before Saving Patient Data", vbInformation: Exit Sub
            'If TxtAddress = "" Then MsgBox "Please fill in the Address before Saving Patient Data", vbInformation: Exit Sub
             If RsAdd.State = 1 Then Set RsAdd = Nothing
             RsAdd.Open "SELECT * FROM PATIENT_DETAILS WHERE CARDNUMBER = '" & TxtCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
                 With RsAdd
                     If .EOF = True Then
                         .AddNew
                     End If
                         '!CardNumber = TxtCardNumber
                         !BILLINGCOMPANY = Trim(Mid(CboBillingCompany, 1, InStr(Trim(CboBillingCompany), "-") - 1))
                         If GetID_NameFromCombo(CboBillingCompany, 1) <> "001" Then !InsuranceCardNumber = TxtInsuranceNumber
                         !SURNAME = TxtSurname
                         !FirstName = TxtFirstname
                         If TxtSecondName = "" Then !SECONDNAME = "O" Else !SECONDNAME = TxtSecondName
                         !DATEOFBIRTH = DTDateofBirth
                         !AGE = TxtAge
                         If TxtID = "" Then !IDNUMBER = "0" Else !IDNUMBER = TxtID
                         !GENDER = CboGender
                         !TELEPHONE = TxtTel
                         If TxtAddress = "" Then !ADDRESS = "0" Else !ADDRESS = TxtAddress
                         '!NURSE = StrCurrentUser
                         
                         'Added for Wellness Associates & NCC
                         
                         If Not IsNull(TxtPOBox) Then !PostOfficeBox = TxtPOBox
                         If Not IsNull(TxtPoBoxCode) Then !PoboxCode = TxtPoBoxCode
                         If Not IsNull(TxtTown) Then !Town = TxtTown
                         'If Not IsNull(TxtEmployer) Then !Employer = TxtEmployer
                         If Not IsNull(TxtOfficeLine) Then !OfficeLine = TxtOfficeLine
                         If Not IsNull(TxtEmail) Then !email = TxtEmail
                         If Not IsNull(TxtFileNumber) Then !FileNumber = TxtFileNumber
                         
                         'RUNNING NUMBER RE-SELECTED HERE INCASE THE NUMBER WAS RE ASSIGNED WHILE INPUT WAS IN PROGRESS.
                         'ASSIGNMENT SHOULD BE AT THE LAST SECOND POSSIBLE TO AVOID DUPLICATION. SOLOMON 23042013
                         TxtCardNumber = AllocateRunningNumber + "/" + Right(GlbSysDate, 2)
                         !CardNumber = TxtCardNumber
                     RsAdd.Update
                     NextOFKin
                     '*********
                     
                     Conn.Execute "UPDATE GENERALPARAMS SET ITEMVALUE = ITEMVALUE + 1 WHERE ITEMNAME = 'RUNNUMBER'"
                     If TxtBookID <> "" Then Conn.Execute "UPDATE APPOINTMENTS SET STATUS = 1 WHERE APPOINTMENTID = '" & TxtBookID & "'"
                    FillPatientGrid
                     MsgBox "Patient Record Updated Succesfully", vbInformation
                 End With
                 RsAdd.Close
                 
                 'SAVE PATIENT POTRAIT IMAGE FROM WEBCAM.
                 SavePatientPicture
                 
             ResetMode
             'FEATURE TO ALLOW INPUT OF NEW PATIENT AND POSTING WITHOUT HAVING TO CLEAR FIELDS
             If ChkPreserve.Value = 0 And lvSaving <> True Then
                ClearControls
             End If
        Case 1
            If TxtBookSurname = "" Then MsgBox "Please fill in the SurName before Saving Patient Data", vbInformation: Exit Sub
            If TxtBookFirstName = "" Then MsgBox "Please fill in the First Name before Saving Patient Data", vbInformation: Exit Sub
            If TxtBookSecondName = "" Then MsgBox "Please fill in the Second Name before Saving Patient Data", vbInformation: Exit Sub
            If OptNewClient.Value = 0 And OptExistingClient.Value = 0 And TxtBookID = "" Then MsgBox "Please Select New or Existing Client Option before Saving Patient Data", vbInformation: Exit Sub
            If TxtBookTelephone = "" Then TxtBookTelephone = 0
            If TxtBookID = "" Then TxtBookID = 0
            
            'If CheckDateTimeAvailability(DtAppointment, DTStartTime, DTEndTime) = False Then Exit Sub
            
            If RsAdd.State = 1 Then Set RsAdd = Nothing
            RsAdd.Open "SELECT * FROM APPOINTMENTS WHERE APPOINTMENTID = '" & TxtBookID & "'", Conn, adOpenStatic, adLockOptimistic
                With RsAdd
                     If .EOF = True Then
                         .AddNew
                     End If
                        If Not IsNull(TxtBookCardNumber) Then !CardNumber = TxtBookCardNumber
                        !FirstName = TxtBookFirstName
                        !SECONDNAME = TxtBookSecondName
                        !SURNAME = TxtBookSurname
                        !TELEPHONE = Replace(TxtBookTelephone, "-", "")
                        !AppDate = Format(DtAppointment, "DD MMM YYYY")
                        !AppStartTime = Right(DTStartTime, 8) 'Mid(DTStartTime.Value, 13, 5)
                        'If Mid(DTStartTime.Value, 13, 5) = "" Then !AppStartTime = Mid(DTStartTime.Value, 1, 5)
                        '!STARTAMPM = CboStartAMPM
                        !AppEndTime = Right(DTEndTime, 8) 'Mid(DTEndTime.Value, 13, 5)
                        'If Mid(DTEndTime.Value, 13, 5) = "" Then !AppEndTime = Mid(DTEndTime.Value, 1, 5)
                        '!Endampm = CboEndAMPM
                        !DOCTOR = Replace(CboDoctor, "<", "")
                        !DOCTOR = Replace(!DOCTOR, ">", "")
                        !Status = 0
                    .Update
                    FillPendingAppointment
                    MsgBox "Appointment Details Saved Succesfully!", vbInformation
                    ResetMode
                    ClearControls
                End With
        End Select
        lvSaving = False
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
    Exit Sub
    Resume
End Sub

Private Sub CmdBillingcompanies_Click()
    Load FrmBillingCompanies
    FrmBillingCompanies.Show
End Sub

Private Sub Command2_Click()
    'CboPaymentMode.Text = CboPaymentMode.ListIndex(1)
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command3_Click()
    MsgBox DTStartTime & " " & DTEndTime
End Sub

Private Sub DtAppointment_Click()
    D = D
End Sub

Private Sub DTDateofBirth_Change()
'    TxtAge = DateDiff("M", DTDateofBirth, Date) \ 12
'    If TxtAge = "" Then Exit Sub
'    If Mid(TxtAge, 1, 1) = "-" Or TxtAge = "0" Then MsgBox "Age is not Valid. Pleasse re-enter", vbExclamation: TxtAge = "": Exit Sub
End Sub

'Private Sub DTDateofBirth_CLICK()
'    TxtAge = DateDiff("M", DTDateofBirth, Date) \ 12
'    If Mid(TxtAge, 1, 1) = "-" Then MsgBox "Age is not Valid. Pleasse re-enter", vbExclamation: TxtAge = "": Exit Sub
'End Sub
'
Private Sub DTDateofBirth_Validate(Cancel As Boolean)
    TxtAge = DateDiff("M", DTDateofBirth, GlbSysDate) \ 12
    'If CInt(TxtAge) <= 0 Then MsgBox "Age is not Valid. Pleasse re-enter", vbExclamation: TxtAge = "": Cancel = True: Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo Errorhandler
    If Format(GlbSysDate, "ddmmmyyyy") <> Format(Date, "ddmmmyyyy") Then
        Beep
        FrmDateDiff.Show 1
    End If
    
    'SET THE DATE PICKERS TO SHOW CURRENT DATE AND NOT THE DATE THEY WERE PLACED THERE.
    For Each CTL In Me.Controls
        If TypeOf CTL Is DTPicker Then
            CTL = GlbSysDate
        End If
    Next
    
    ResetMode
    centerform Me
    DisableControls
    FillPatientGrid
    'FillPendingAppointment
    'FillMissedAppointments
    'FillHonoredAppointments
    
    'SET DEFAULT TAB
    PatientTab2.Tab = 0
        
    'POPULATE SIMPLE COMBOS
    CboGender.AddItem "MALE"
    CboGender.AddItem "FEMALE"
    
    'POPULATE GRID
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM SERVICE_PROVIDER", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboBillingCompany.AddItem RsRecords!COMPANYCODE & " - " & RsRecords!SERVICEPROVIDER
            RsRecords.MoveNext
        Wend
    RsRecords.Close
        
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    RsRecords.Open "SELECT * FROM PAYMENT_MODES", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboPaymentMode.AddItem RsRecords!PAYMENTCODE & " - " & RsRecords!PAYMENTDESCRIPTION
            RsRecords.MoveNext
        Wend
        CboPaymentMode.Text = "1 - CASH"
    RsRecords.Close
    
   'POPULATE DOCTORS COMBO BOX
    If RsRecords.State = 1 Then Set RsRecords = Nothing
    CboDoctor.AddItem "<< ANY >>"
    RsRecords.Open "SELECT * FROM PROFILES WHERE RIGHTS = 'DOCTORS' AND ACCESS = 'TRUE'", Conn, adOpenStatic, adLockOptimistic
        While RsRecords.EOF = False
            CboDoctor.AddItem RsRecords!UserName
            CboDoctors.AddItem RsRecords!UserName
            RsRecords.MoveNext
        Wend
    RsRecords.Close
    
    
  'ASSIGN FORM NAME AS CURRENT FORM
  GlbCurrentForm = EnumConsultation
  
  'DISABLE AND ENABLE NAVIGATION BUTTONS
    ManageProcessFlow EnumConsultation
    
    PatientsTab.Tab = 0
    TabAppointments.Tab = 0
    OptObservation_Click (1)
    If OptObservation.Item(1).Value = False Then OptObservation.Item(1).Value = True
    
    DTDateofBirth.Value = GlbSysDate
    DtAppointment.Value = GlbSysDate
    DtDateFilter.Value = GlbSysDate
    
    
    Exit Sub
Errorhandler:
    MsgBox Err.Number & " " & Err.Description
    Exit Sub
    Resume
End Sub
Private Sub AddMode()
    CmdNew.Enabled = False
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = True
    CmdExit.Caption = "Cancel"
    Adding = True
End Sub
Private Sub EditMode()
    CmdNew.Enabled = False
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = True
    CmdExit.Caption = "Cancel"
    Editing = True
End Sub
Private Sub ResetMode()
    CmdNew.Enabled = True
    CmdEdit.Enabled = True
    CmdDelete.Enabled = True
    CmdSave.Enabled = False
    CmdExit.Caption = "Exit"
    Adding = False
    Editing = False
    If ChkPreserve.Value = 1 And lvSaving <> True Then
        ClearText FrmPatients
    End If
End Sub
Private Sub EnableControls()
    Dim CTL As Control
    For Each CTL In Me.Controls
        If TypeOf CTL Is TextBox Then
            CTL.Enabled = True
        ElseIf TypeOf CTL Is ComboBox Then
            CTL.Enabled = True
        ElseIf TypeOf CTL Is DTPicker Then
            CTL.Enabled = True
        End If
    Next
    TxtBookID.Enabled = False
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
Private Sub NextOFKin()
On Error GoTo Errorhandler
    
    If RsNextOfKin.State = 1 Then Set RsNextOfKin = Nothing
    RsNextOfKin.Open "SELECT * FROM NEXT_OF_KIN WHERE CARDNUMBER = '" & TxtCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
        With RsNextOfKin
            If .EOF = True Then
                .AddNew
            End If
                !CardNumber = TxtCardNumber
                If TxtKinFirstName <> "" Then !FirstName = TxtKinFirstName
                If TxtKinSecondName <> "" Then !SECONDNAME = TxtKinSecondName
                If TxtKinRelationship <> "" Then !Relationship = TxtKinRelationship
                If TxtKINAddress = "" Then !ADDRESS = "O" Else !ADDRESS = TxtKINAddress
                If Not IsNumeric(TxtKinTel) Then TxtKinTel = 0
                If TxtKinTel <> "" Then !TELEPHONE = TxtKinTel
                If TxtKinEmail <> "" Then !email = TxtKinEmail
            .Update
        End With
    RsNextOfKin.Close
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Private Sub FillPostedToConsultant()
    On Error GoTo Errorhandler
    GridConsult.Clear
    GridConsult.Rows = 1
    GridConsult.Cols = 3
    GridConsult.ColAlignment(1) = flexAlignCenterCenter
    GridConsult.ColWidth(1) = 3105
    GridConsult.ColWidth(2) = 3990
    GridConsult.FormatString = "CARD NUMBER|  VISIT NUMBER  |PATIENTS FULL NAME              .  |   BILLING COMPANY   "
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'RsGrid.Open "SELECT * FROM PATIENT_DETAILS", Conn, adOpenStatic, adLockOptimistic
        RsGrid.Open "SELECT * FROM PATIENT_DETAILS INNER JOIN COMPLAINS ON PATIENT_DETAILS.CARDNUMBER = COMPLAINS.CARDNUMBER AND COMPLAINS.TOCONSULTATION = '1'", Conn, adOpenStatic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridConsult.AddItem !CardNumber & vbTab & !VISITNUMBER & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                        .MoveNext
                    Wend
                End With
            End If
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
End Sub
Private Sub FillPatientGrid()
    On Error GoTo Errorhandler
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 3
    Grid.ColAlignment(1) = flexAlignCenterCenter
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER|   PATIENTS FULL NAME                 .  |   BILLING COMPANY     |NATIONAL ID NUMBER"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT  top 10 * FROM PATIENT_DETAILS", Conn, adOpenStatic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !CardNumber & vbTab & !FirstName & " " & !SECONDNAME & " " & !SURNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                        .MoveNext
                    Wend
                End With
            End If
        RsGrid.Close
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub FillPatientSearch_FirstName()
    On Error GoTo Errorhandler
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 3
    Grid.ColAlignment(1) = flexAlignCenterCenter
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER|   PATIENTS FULL NAME          .  |   BILLING COMPANY     |NATIONAL ID NUMBER"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        WAMBOH = TxtFirstname + "%"
        If WAMBOH = "%" Then
            RsGrid.Open "SELECT  top 550 * FROM PATIENT_DETAILS", Conn, adOpenStatic, adLockOptimistic
        Else
            RsGrid.Open "SELECT  top 550 * FROM PATIENT_DETAILS WHERE FIRSTNAME LIKE  '" & WAMBOH & "'", Conn, adOpenStatic, adLockOptimistic
        End If
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                        .MoveNext
                    Wend
                End With
            End If
        RsGrid.Close
    Exit Sub
Errorhandler:
    MsgBox Err.Description
   ' Resume
End Sub
Private Sub FillPatientSearch_SecondName()
    On Error GoTo Errorhandler
    Grid.Clear
    Grid.Rows = 1
    Grid.Cols = 3
    Grid.ColAlignment(1) = flexAlignCenterCenter
    Grid.ColWidth(1) = 3105
    Grid.ColWidth(2) = 3990
    Grid.FormatString = "CARD NUMBER|   PATIENTS FULL NAME  |   BILLING COMPANY     |NATIONAL ID NUMBER"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        WAMBOH = TxtSecondName + "%"
        If WAMBOH = "%" Then
            RsGrid.Open "SELECT  top 2 * FROM PATIENT_DETAILS", Conn, adOpenStatic, adLockOptimistic
        Else
            'MsgBox "First and Second Name Search", vbInformation
            RsGrid.Open "SELECT  top 550 * FROM PATIENT_DETAILS WHERE FIRSTNAME = '" & TxtFirstname & "' AND SECONDNAME LIKE  '" & WAMBOH & "'", Conn, adOpenStatic, adLockOptimistic
        End If
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        Grid.AddItem !CardNumber & vbTab & !SURNAME & " " & !FirstName & " " & !SECONDNAME & vbTab & !BILLINGCOMPANY & vbTab & !IDNUMBER
                        .MoveNext
                    Wend
                End With
            End If
        RsGrid.Close
    Exit Sub
Errorhandler:
    MsgBox Err.Description
   ' Resume
End Sub

Private Sub FillPendingAppointment()
    On Error GoTo Errorhandler
    Dim Rcount As Integer
    GridAppointments.Clear
    GridAppointments.Rows = 1
    GridAppointments.Cols = 3
    GridAppointments.ColAlignment(1) = flexAlignCenterCenter
    GridAppointments.ColWidth(1) = 3105
    GridAppointments.ColWidth(2) = 3990
    GridAppointments.FormatString = "FIRST NAME |   SECOND NAME  |   SURNAME     |  PHONE NUMBER |APPOINTMENT DATE|    START TIME  | END TIME  |  DOCTOR TO BE SEEN  | ID | CARD NO"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        'WHERE APPDATE >= '" & Format(GlbSysDate, "DD MMM YYYY") & "'
        RsGrid.Open "SELECT * FROM APPOINTMENTS  ORDER BY APPDATE DESC", Conn, adOpenStatic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridAppointments.AddItem !FirstName & vbTab & !SECONDNAME & vbTab & !SURNAME & vbTab & !TELEPHONE & vbTab & !AppDate & vbTab & !AppStartTime & vbTab & !AppEndTime & vbTab & !DOCTOR & vbTab & !AppointmentID & vbTab & !CardNumber
                        .MoveNext
                        Rcount = Rcount + 1
                    Wend
                End With
            End If
        RsGrid.Close
        TxtPendingApp = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub SearchPendingAppointment()
    On Error GoTo Errorhandler
    Dim Rcount As Integer
    GridAppointments.Clear
    GridAppointments.Rows = 1
    GridAppointments.Cols = 3
    GridAppointments.ColAlignment(1) = flexAlignCenterCenter
    GridAppointments.ColWidth(1) = 3105
    GridAppointments.ColWidth(2) = 3990
    'GridAppointments.FormatString = "FIRST NAME |   SECOND NAME  |   SURNAME     |  PHONE NUMBER |APPOINTMENT DATE| TIME |  DOCTOR TO BE SEEN  | ID | CARD NO"
     GridAppointments.FormatString = "FIRST NAME |   SECOND NAME  |   SURNAME     |  PHONE NUMBER |APPOINTMENT DATE|    START TIME  | END TIME  |  DOCTOR TO BE SEEN  | ID | CARD NO"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        WAMBOH = TxtBookFirstName + "%"
        RsGrid.Open "SELECT * FROM APPOINTMENTS WHERE FIRSTNAME LIKE '" & WAMBOH & "' ORDER BY APPDATE DESC", Conn, adOpenStatic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        'GridAppointments.AddItem !FirstName & vbTab & !SECONDNAME & vbTab & !SURNAME & vbTab & !TELEPHONE & vbTab & !AppDate & vbTab & Right(!AppStartTime, 11) & vbTab & !DOCTOR & vbTab & !AppointmentID & vbTab & !cardnumber
                         GridAppointments.AddItem !FirstName & vbTab & !SECONDNAME & vbTab & !SURNAME & vbTab & !TELEPHONE & vbTab & !AppDate & vbTab & Right(!AppStartTime, 11) + " " + !STARTAMPM & vbTab & Right(!AppEndTime, 11) + " " + !Endampm & vbTab & !DOCTOR & vbTab & !AppointmentID & vbTab & !CardNumber
                        .MoveNext
                        Rcount = Rcount + 1
                    Wend
                End With
            End If
        RsGrid.Close
        TxtPendingApp = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub FillMissedAppointments()
    On Error GoTo Errorhandler
    Dim Rcount As Integer
    GridMissedApp.Clear
    GridMissedApp.Rows = 1
    GridMissedApp.Cols = 3
    GridMissedApp.ColAlignment(1) = flexAlignCenterCenter
    GridMissedApp.ColDataType(0) = flexDTBoolean
    GridMissedApp.Editable = True
    GridMissedApp.ColWidth(1) = 3105
    GridMissedApp.ColWidth(2) = 3990
    GridMissedApp.FormatString = "TICK | FIRST NAME |   SECOND NAME  |   SURNAME     |  PHONE NUMBER |APPOINTMENT DATE| TIME |  DOCTOR TO BE SEEN  | ID | CARD NO"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT * FROM APPOINTMENTS WHERE APPDATE < '" & Format(GlbSysDate, "DD MMM YYYY") & "' ORDER BY APPDATE DESC", Conn, adOpenStatic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridMissedApp.AddItem 0 & vbTab & !FirstName & vbTab & !SECONDNAME & vbTab & !SURNAME & vbTab & !TELEPHONE & vbTab & !AppDate & vbTab & Right(!AppStartTime, 11) & vbTab & !DOCTOR & vbTab & !AppointmentID & vbTab & !CardNumber
                        .MoveNext
                        Rcount = Rcount + 1
                    Wend
                End With
            End If
        RsGrid.Close
        TxtMissedApp = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Private Sub FillHonoredAppointments()
    On Error GoTo Errorhandler
    Dim Rcount As Integer
    GridHonored.Clear
    GridHonored.Rows = 1
    GridHonored.Cols = 3
    GridHonored.ColAlignment(1) = flexAlignCenterCenter
    GridHonored.ColWidth(1) = 3105
    GridHonored.ColWidth(2) = 3990
    GridHonored.FormatString = "FIRST NAME |   SECOND NAME  |   SURNAME     |  PHONE NUMBER |APPOINTMENT DATE| TIME |  DOCTOR SEEN  | ID | CARD NO"
        If RsGrid.State = adStateOpen Then RsGrid.Close
        RsGrid.Open "SELECT * FROM APPOINTMENTS WHERE APPDATE >= '" & Format(GlbSysDate, "DD MMM YYYY") & "' AND STATUS = '1' ORDER BY APPDATE DESC", Conn, adOpenStatic, adLockOptimistic
            If RsGrid.RecordCount <> 0 Then
                With RsGrid
                    While Not .EOF
                        GridHonored.AddItem !FirstName & vbTab & !SECONDNAME & vbTab & !SURNAME & vbTab & !TELEPHONE & vbTab & !AppDate & vbTab & Right(!AppStartTime, 11) & vbTab & !DOCTOR & vbTab & !AppointmentID & vbTab & !CardNumber
                        .MoveNext
                        Rcount = Rcount + 1
                    Wend
                End With
            End If
        RsGrid.Close
        TxtHonoredApp = Rcount
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub
Public Sub ManageProcessFlow(ActiveForm)
    Dim RsControls As New ADODB.Recordset
    RsControls.Open "SELECT * FROM PROCESSFLOW WHERE SCREENID = '" & ActiveForm & "'", Conn, adOpenStatic, adLockOptimistic
        If RsControls.EOF = False Then
            With RsControls
                If !OBSERVATION = 1 Then OptObservation.Item(1).Enabled = True
                If !DOCTORS = 1 Then OptDoctors.Item(2).Enabled = True
                If !CASHIER = 1 Then OptCashier.Item(3).Enabled = True
                If !PHARMACY = 1 Then OptPharmacy.Item(4).Enabled = True
                If !LAB = 1 Then OptLab.Item(5).Enabled = True
            End With
        End If
End Sub

Private Sub Grid_DblClick()
On Error GoTo Errorhandler
Dim rsfill As New ADODB.Recordset
    If Grid.Row = 0 Then Exit Sub
        GlbCardNumber = Grid.TextMatrix(Grid.Row, 0)
        If GlbCalledFromPatients = True Or GlbCalledFromAppointments = True Then
            GlbCardNumber = StrPharmCardNumber
            GlbCalledFromPatients = False
            GlbCalledFromAppointments = False
        Else
            StrPharmCardNumber = Grid.TextMatrix(Grid.Row, 0)
        End If
        
        'StrPharmVisitDate = Grid.TextMatrix(Grid.Row, 7)
        If rsfill.State = 1 Then rsfill.Close
        If StrPharmCardNumber = "" Then Exit Sub
                rsfill.Open "SELECT * FROM PATIENT_DETAILS WHERE CARDNUMBER = '" & StrPharmCardNumber & "'", Conn, adOpenStatic, adLockOptimistic
                    With rsfill
                        
                        TxtCardNumber = StrPharmCardNumber
                        CboBillingCompany = GetBillingCompanyName(!BILLINGCOMPANY)
                        TxtInsuranceNumber = !InsuranceCardNumber
                        TxtSurname = !SURNAME
                        TxtFirstname = !FirstName
                        TxtSecondName = !SECONDNAME
                        TxtTel = !TELEPHONE
                        TxtID = !IDNUMBER
                        'WACHA HII IJAZWE NA RECEPTIONIST JUU IT WILL VARY VISIT FROM THE OTHER.
                        'CboPaymentMode.ListIndex = !PAYMENTMODE - 1
                        If Not IsNull(!ADDRESS) Then
                            TxtAddress = !ADDRESS
                        End If
                        If Not IsNull(!AGE) Then
                            TxtAge = !AGE
                        End If
                        CboGender = !GENDER
                        
                        TxtInsuranceNumber = !InsuranceCardNumber
                        TxtPOBox = !PostOfficeBox
                        TxtPoBoxCode = !PoboxCode
                        TxtTown = !Town
                         'TxtEmployer = !Employer
                        TxtOfficeLine = !OfficeLine
                        TxtEmail = !email
                        If !FileNumber <> "" Then TxtFileNumber = !FileNumber
                        
                rsfill.Close
                FillNextOfKin
                RetrievePatientPicture TxtCardNumber
            End With
        CboPaymentMode.Enabled = True
        CboDoctors.Enabled = True
Exit Sub
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Please contact System Administrator"
   Exit Sub
   Resume
End Sub
Private Function GetBillingCompanyName(ByVal ID)
    GetBillingCompanyName = FindRecord("SERVICE_PROVIDER", "SERVICEPROVIDER", "COMPANYCODE = '" & Trim(ID) & "'")
    GetBillingCompanyName = ID + " - " + GetBillingCompanyName
End Function

Private Sub GridAppointments_dblclick()
On Error GoTo Errorhandler
    If GridAppointments.Row = 0 Then Exit Sub
    lvDoNotSearch = True
    TxtBookID = GridAppointments.TextMatrix(GridAppointments.Row, 8)
    TxtBookCardNumber = GridAppointments.TextMatrix(GridAppointments.Row, 8): If GridAppointments.TextMatrix(GridAppointments.Row, 8) = "" Then TxtBookCardNumber = "NEW CLIENT"
    TxtBookFirstName = GridAppointments.TextMatrix(GridAppointments.Row, 0)
    TxtBookSecondName = GridAppointments.TextMatrix(GridAppointments.Row, 1)
    TxtBookSurname = Trim(GridAppointments.TextMatrix(GridAppointments.Row, 2))
    TxtBookTelephone = GridAppointments.TextMatrix(GridAppointments.Row, 3)
    DtAppointment = GridAppointments.TextMatrix(GridAppointments.Row, 4)
    If GridAppointments.TextMatrix(GridAppointments.Row, 5) <> "" Then DTStartTime = Left(GridAppointments.TextMatrix(GridAppointments.Row, 5), 5): DTEndTime = Left(GridAppointments.TextMatrix(GridAppointments.Row, 6), 5)
    CboStartAMPM = Right(GridAppointments.TextMatrix(GridAppointments.Row, 5), 2)
    CboEndAMPM = Right(GridAppointments.TextMatrix(GridAppointments.Row, 6), 2)
    TxtBookDoctor = GridAppointments.TextMatrix(GridAppointments.Row, 7)
    CboDoctor = GridAppointments.TextMatrix(GridAppointments.Row, 7)
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub OptAdmission_Click(Index As Integer)
    ItemSelected = OptAdmission.Item(Index).Index
End Sub

Private Sub GridMissedApp_Click()
On Error GoTo Errorhandler
    lvDoNotSearch = True
    TxtBookFirstName = GridMissedApp.TextMatrix(GridMissedApp.Row, 1)
    TxtBookSecondName = GridMissedApp.TextMatrix(GridMissedApp.Row, 2)
    TxtBookSurname = GridMissedApp.TextMatrix(GridMissedApp.Row, 3)
    TxtBookTelephone = GridMissedApp.TextMatrix(GridMissedApp.Row, 4)
    If GridMissedApp.TextMatrix(GridMissedApp.Row, 5) <> "" Then DtAppointment = GridMissedApp.TextMatrix(GridMissedApp.Row, 5)
    CboDoctor = GridMissedApp.TextMatrix(GridMissedApp.Row, 7)
    TxtBookCardNumber = GridMissedApp.TextMatrix(GridMissedApp.Row, 9): If GridMissedApp.TextMatrix(GridMissedApp.Row, 9) = "" Then TxtBookCardNumber = "NEW CLIENT"
    DTTime = GridMissedApp.TextMatrix(GridMissedApp.Row, 6)
    Exit Sub
Errorhandler:
    MsgBox Err.Description
End Sub

Private Sub Image1_Click()

End Sub

Private Sub ImgPreview_DblClick()
    FrmPicture.Picture = LoadPicture(Picture1.Picture)
    FrmPicture.Show 1
End Sub

Private Sub MnuOfficial_Click()
    FrmNonpatients.Show 1
End Sub

Private Sub OptCashier_Click(Index As Integer)
        ItemSelected = OptCashier.Item(Index).Index
End Sub

Private Sub OptDoctors_Click(Index As Integer)
    ItemSelected = OptDoctors.Item(Index).Index
End Sub

Private Sub OptExistingClient_Click()
    If OptExistingClient.Value = True Then
        CmdBookSearch.Enabled = True
        TxtBookCardNumber.Locked = False
    Else
        CmdBookSearch.Enabled = False
        TxtBookCardNumber.Locked = True
    End If
End Sub

Private Sub OptLab_Click(Index As Integer)
    ItemSelected = OptLab.Item(Index).Index
End Sub

Private Sub OptNewClient_Click()
    If OptExistingClient.Value = True Then
        CmdBookSearch.Enabled = True
        TxtBookCardNumber.Locked = False
    Else
        CmdBookSearch.Enabled = False
        TxtBookCardNumber.Locked = True
    End If
End Sub

Private Sub OptObservation_Click(Index As Integer)
    ItemSelected = OptObservation.Item(Index).Index
End Sub

Private Sub OptPharmacy_Click(Index As Integer)
    ItemSelected = OptPharmacy.Item(Index).Index
End Sub
    


Private Sub PatientTab2_Click(PreviousTab As Integer)
    FillPostedToConsultant
End Sub

Private Sub Picture1_DblClick()
    FrmPicture.Picture = LoadPicture(Picture1.Picture)
    FrmPicture.Show 1
End Sub

Private Sub TxtBookFirstName_Change()
    If lvDoNotSearch = True Then lvDoNotSearch = False: Exit Sub
    SearchPendingAppointment
End Sub

Private Sub TxtBookFirstName_LostFocus()
    TxtBookFirstName = ConvertToUppercase(TxtBookFirstName)
End Sub

Private Sub TxtBookSecondName_LostFocus()
    TxtBookSecondName = ConvertToUppercase(TxtBookSecondName)
End Sub

Private Sub TxtBookSurname_LostFocus()
    TxtBookSurname = ConvertToUppercase(TxtBookSurname)
End Sub

Private Sub TxtCardNumber_Change()
    'ValidateDataType TxtCardNumber, 1, "FrmPatients", "TxtSurname"
End Sub

Private Sub TxtFirstName_Change()
    ValidateDataType TxtFirstname, 1, "FrmPatients", "TxtFirstName"
    FillPatientSearch_FirstName
End Sub

Private Sub TxtFirstName_LostFocus()
    TxtFirstname = ConvertToUppercase(TxtFirstname)
End Sub

Private Sub TxtID_Change()
    ValidateDataType TxtID, 0, "FrmPatients", "TxtID"
End Sub

Private Sub TxtKinAddress_LostFocus()
    ConvertToUppercase TxtKINAddress
End Sub

Private Sub TxtKinFirstName_Change()
    ValidateDataType TxtKinFirstName, 1, "FrmPatients", "TxtKinFirstName"
End Sub

Private Sub TxtKinFirstName_LostFocus()
    TxtKinFirstName = ConvertToUppercase(TxtKinFirstName)
End Sub

Private Sub TxtKinID_Change()
    'ValidateDataType TxtKinID, 0, "FrmPatients", "TxtKinID"
End Sub

Private Sub TxtKinRelationship_LostFocus()
    ConvertToUppercase TxtKinRelationship
End Sub

Private Sub TxtKinSecondName_Change()
    ValidateDataType TxtKinSecondName, 1, "FrmPatients", "TxtKinSecondName"
End Sub

Private Sub TxtKinSecondName_LostFocus()
    TxtKinSecondName = ConvertToUppercase(TxtKinSecondName)
End Sub

Private Sub TxtKinTel_Change()
    'ValidateDataType TxtKinTel, 0, "FrmPatients", "TxtKinTel"
End Sub

Private Sub TxtSecondName_Change()
    ValidateDataType TxtSecondName, 1, "FrmPatients", "TxtSecondName"
    FillPatientSearch_SecondName
End Sub

Private Sub TxtSecondName_LostFocus()
    TxtSecondName = ConvertToUppercase(TxtSecondName)
End Sub

Private Sub TxtSurname_Change()
'    If ValidateDataType(TxtSurname, 1) = False Then
'        MsgBox "Only Characters are allowed in this field", vbInformation
'    End If
    ValidateDataType TxtSurname, 1, "FrmPatients", "TxtSurname"
End Sub

Private Sub TxtSurname_LostFocus()
    TxtSurname = ConvertToUppercase(TxtSurname)
End Sub

Private Sub TxtTel_Change()
    'Do not Validate to allow foreign numbers with alphanumeric
   ' ValidateDataType TxtTel, 0, "FrmPatients", "TxtTel"
End Sub
Public Function LoadPictureFromDB(ByRef rs As ADODB.Recordset, ByVal fldName As String, ByRef Image1 As Object, Optional ByVal strFileName As String)

    On Error GoTo Errorhandler
    
    'If Recordset is Empty, Then Exit
    If rs Is Nothing Then
        'GoTo procNoPicture
    End If
    
    Dim strTempFileName As String
    'strTempFileName = GetParam("IMAGETEMPFILE")
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    
    strStream.Write rs.Fields(fldName).Value

    If strFileName = "" Then
        strFileName = "c:\Temp.bmp"
    End If
    strStream.SaveToFile strFileName, adSaveCreateOverWrite
    ImgPreview.Picture = LoadPicture(strFileName)
    On Error Resume Next
    'Image1.DisplayBlankImage Image1.Width, Image1.Height, 800, 600, 2
    'Image1.Display
    If Err.Number <> 0 Then
        'Image1.DisplayBlankImage Image1.Width, Image1.Height, 800, 600, 2
        'Image1.Refresh
        ImgPreview.Picture = LoadPicture(App.Path & "\NoImage.bmp")
        LoadPictureFromDB = False
    Else
        LoadPictureFromDB = True
    End If
    On Error GoTo 0
    'Kill ("C:\Temp.tif")
    'LoadPictureFromDB = True

Exit Function
Errorhandler:
    MsgBox Err.Description, vbExclamation, "Image Problem"
    'SystemErrorHandler Err.Number, Err.Description
End Function

