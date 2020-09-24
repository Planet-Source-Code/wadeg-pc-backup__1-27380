VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   Caption         =   "PC Backup ver 1.0"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   210
   ClientWidth     =   9390
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   59000
      Left            =   1920
      Top             =   6360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Config"
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Unload && Exit"
      Height          =   615
      Left            =   7800
      TabIndex        =   25
      Top             =   6360
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Source"
      TabPicture(0)   =   "frmMain.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblSource"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkIncSub"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstSource"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "driveSource"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dirSource"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fileSource"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAddDir"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAddFile"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdClrSoucre"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdExclude"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Destination"
      TabPicture(1)   =   "frmMain.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDest"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkCustom"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtDest"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "driveDest"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "dirDest"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Schedule"
      TabPicture(2)   =   "frmMain.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblSched"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblAlwaysTime(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblAlwaysTime(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblWhen"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label4"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkDays(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "chkDays(3)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "chkDays(4)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "chkDays(5)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "chkDays(6)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "chkDays(7)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "chkDays(1)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "chkLog"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "chkIncr"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "optAlways"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "optEach"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "optDaily"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "MaskEdBox2"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "MaskEdBox1"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "cmdBakNow"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).ControlCount=   23
      TabCaption(3)   =   "Setup / Options"
      TabPicture(3)   =   "frmMain.frx":091E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblAppSetup"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblAppOptions"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "chkServ"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabPicture(4)   =   "frmMain.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "About / Info"
      TabPicture(5)   =   "frmMain.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblAbout"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lblContact"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lblPhone(2)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "lblMe"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Progress"
      TabPicture(6)   =   "frmMain.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label10"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label11"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label12"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Label13"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Label14"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "cmdLog"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).ControlCount=   6
      Begin VB.CommandButton cmdExclude 
         Caption         =   "Exclude"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69600
         TabIndex        =   63
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Custom Directory Options"
         Enabled         =   0   'False
         Height          =   4215
         Left            =   -71280
         TabIndex        =   51
         Top             =   1560
         Width           =   5175
         Begin VB.CommandButton cmdReset 
            Caption         =   "Reset"
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   3240
            Width           =   1215
         End
         Begin VB.ComboBox cmbSep 
            Height          =   315
            ItemData        =   "frmMain.frx":098E
            Left            =   240
            List            =   "frmMain.frx":099B
            TabIndex        =   60
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtCusDir 
            Height          =   285
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Seperator"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   1920
            Width           =   1215
         End
         Begin MSForms.OptionButton optOr 
            Height          =   255
            Left            =   1080
            TabIndex        =   59
            Top             =   960
            Width           =   615
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1085;450"
            Value           =   "0"
            Caption         =   "Or"
            GroupName       =   "CusDir"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton optAnd 
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   960
            Width           =   735
            VariousPropertyBits=   746588185
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "1296;450"
            Value           =   "0"
            Caption         =   "And"
            GroupName       =   "CusDir"
            FontEffects     =   1073750016
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblCustomDir2 
            Caption         =   "e.g. My Projects"
            Height          =   375
            Left            =   2520
            TabIndex        =   56
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblCustomDir 
            Caption         =   "Custom Directory Name"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblDated 
            Caption         =   "e.g. C:\My Dir\MM-DD-YYYY"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   1560
            Width           =   2295
         End
         Begin MSForms.CheckBox chkDated 
            Height          =   375
            Left            =   240
            TabIndex        =   52
            Top             =   1200
            Width           =   1695
            VariousPropertyBits=   746588191
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2990;661"
            Value           =   "0"
            Caption         =   "Dated Directories"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.CommandButton cmdLog 
         Caption         =   "View Log File"
         Height          =   495
         Left            =   -71280
         TabIndex        =   50
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdClrSoucre 
         Caption         =   "Clear"
         Height          =   375
         Left            =   -67560
         TabIndex        =   41
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton cmdBakNow 
         Caption         =   "Backup Now"
         Height          =   495
         Left            =   -69960
         TabIndex        =   22
         Top             =   3360
         Width           =   1455
      End
      Begin VB.DirListBox dirDest 
         Height          =   3690
         Left            =   -74640
         TabIndex        =   11
         Top             =   2040
         Width           =   3015
      End
      Begin VB.DriveListBox driveDest 
         Height          =   315
         Left            =   -74640
         TabIndex        =   10
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtDest 
         Height          =   285
         Left            =   -74640
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   6375
      End
      Begin VB.CommandButton cmdAddFile 
         Caption         =   "+ File(s)"
         Height          =   375
         Left            =   -71160
         TabIndex        =   7
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddDir 
         Caption         =   "+ Directory"
         Height          =   375
         Left            =   -72840
         TabIndex        =   6
         Top             =   5640
         Width           =   1215
      End
      Begin VB.FileListBox fileSource 
         Height          =   1845
         Left            =   -74760
         TabIndex        =   5
         Top             =   3480
         Width           =   3135
      End
      Begin VB.DirListBox dirSource 
         Height          =   2565
         Left            =   -74760
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
      Begin VB.DriveListBox driveSource 
         Height          =   315
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   3135
      End
      Begin VB.ListBox lstSource 
         Height          =   4740
         Left            =   -71160
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   4815
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   -73200
         TabIndex        =   20
         Top             =   1080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   -73200
         TabIndex        =   21
         Top             =   1440
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMe 
         Caption         =   "Redesigned and coded by Wade Grimm"
         Height          =   255
         Left            =   -74760
         TabIndex        =   64
         Top             =   1320
         Width           =   6375
      End
      Begin MSForms.CheckBox chkCustom 
         Height          =   375
         Left            =   -74640
         TabIndex        =   57
         Top             =   1080
         Width           =   2175
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3836;661"
         Value           =   "0"
         Caption         =   "Custom Directory Name"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74400
         TabIndex        =   49
         Top             =   3120
         Width           =   7815
      End
      Begin VB.Label Label13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74400
         TabIndex        =   48
         Top             =   2520
         Width           =   7815
      End
      Begin VB.Label Label12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74400
         TabIndex        =   47
         Top             =   1920
         Width           =   7815
      End
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74400
         TabIndex        =   46
         Top             =   1320
         Width           =   7815
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74400
         TabIndex        =   45
         Top             =   600
         Width           =   7815
      End
      Begin MSForms.OptionButton optDaily 
         Height          =   345
         Left            =   -73560
         TabIndex        =   44
         Top             =   1920
         Width           =   750
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1323;609"
         Value           =   "1"
         Caption         =   "Daily"
         GroupName       =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optEach 
         Height          =   345
         Left            =   -73560
         TabIndex        =   43
         Top             =   1440
         Width           =   405
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "714;609"
         Value           =   "0"
         GroupName       =   "Sch"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optAlways 
         Height          =   345
         Left            =   -73560
         TabIndex        =   42
         Top             =   1080
         Width           =   405
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "714;609"
         Value           =   "1"
         GroupName       =   "Sch"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkServ 
         Height          =   375
         Left            =   720
         TabIndex        =   40
         Top             =   960
         Width           =   2175
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3836;661"
         Value           =   "0"
         Caption         =   "Load with Windows"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkIncSub 
         Height          =   495
         Left            =   -74640
         TabIndex        =   39
         Top             =   5520
         Width           =   1575
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2778;873"
         Value           =   "1"
         Caption         =   "Include Sub Directories"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkIncr 
         Height          =   375
         Left            =   -74280
         TabIndex        =   38
         Top             =   3600
         Width           =   2175
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3836;661"
         Value           =   "0"
         Caption         =   "Incremental backup"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkLog 
         Height          =   375
         Left            =   -74280
         TabIndex        =   37
         Top             =   3240
         Width           =   2175
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3836;661"
         Value           =   "1"
         Caption         =   "Save results to Log"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDays 
         Height          =   345
         Index           =   1
         Left            =   -70080
         TabIndex        =   36
         Top             =   2640
         Width           =   975
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1720;609"
         Value           =   "0"
         Caption         =   "Sunday"
         GroupName       =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDays 
         Height          =   345
         Index           =   7
         Left            =   -71280
         TabIndex        =   35
         Top             =   2640
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;609"
         Value           =   "0"
         Caption         =   "Saturday"
         GroupName       =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDays 
         Height          =   345
         Index           =   6
         Left            =   -72360
         TabIndex        =   34
         Top             =   2640
         Width           =   975
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1720;609"
         Value           =   "0"
         Caption         =   "Friday"
         GroupName       =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDays 
         Height          =   345
         Index           =   5
         Left            =   -73440
         TabIndex        =   33
         Top             =   2640
         Width           =   975
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1720;609"
         Value           =   "0"
         Caption         =   "Thusday"
         GroupName       =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDays 
         Height          =   345
         Index           =   4
         Left            =   -71280
         TabIndex        =   32
         Top             =   2280
         Width           =   1215
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2143;609"
         Value           =   "0"
         Caption         =   "Wednesday"
         GroupName       =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDays 
         Height          =   345
         Index           =   3
         Left            =   -72360
         TabIndex        =   31
         Top             =   2280
         Width           =   975
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1720;609"
         Value           =   "0"
         Caption         =   "Tuesday"
         GroupName       =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkDays 
         Height          =   345
         Index           =   2
         Left            =   -73440
         TabIndex        =   30
         Top             =   2280
         Width           =   975
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1720;609"
         Value           =   "0"
         Caption         =   "Monday"
         GroupName       =   "Daily"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblPhone 
         Caption         =   "wadeg@itekk.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   29
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lblContact 
         Caption         =   "Contact Info:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblAbout 
         Caption         =   "PC Backup Ver 1.0 originally Designed and by Alexandre Moro, under the name AutoBack"
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   960
         Width           =   7575
      End
      Begin VB.Label lblAppOptions 
         Caption         =   "Application Options:"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblAppSetup 
         Caption         =   "Application Setup:"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Or only on:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   19
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblWhen 
         Caption         =   "When:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   18
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Begins on Save or application Start-up"
         Height          =   255
         Left            =   -71280
         TabIndex        =   17
         Top             =   1470
         Width           =   3015
      End
      Begin VB.Label lblAlwaysTime 
         Caption         =   "Hours:Minutes"
         Height          =   195
         Index           =   1
         Left            =   -72480
         TabIndex        =   16
         Top             =   1465
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Or each:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   15
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblAlwaysTime 
         Caption         =   "Hours:Minutes"
         Height          =   195
         Index           =   0
         Left            =   -72480
         TabIndex        =   14
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Always at:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblSched 
         Caption         =   "Schedule"
         Height          =   255
         Left            =   -74760
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDest 
         Caption         =   "Destination"
         Height          =   255
         Left            =   -74640
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblSource 
         Caption         =   "Files / Directories"
         Height          =   255
         Left            =   -71160
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Window"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Now"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PC Backup by Wade Grimm
' Adapted from AutoBack by Alexandre Moro
' Thanks go out to DarkJedi for the Registry Routines
' Thanks to any one I forgot to mention or couldn't remeber 'cuz they didn't put thier by-line in the coode I used...  ;-(

' If you modify this code please let me know:  wadeg@itekk.com or rozsa@telusplanet.net


Option Explicit

Private Sub chkCustom_Click()
    If chkCustom Then
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
        txtDest = dirDest.Path
        txtCusDir = ""
        chkDated = False
    End If
End Sub

Private Sub chkDated_Click()
Dim zx
    If chkDated.Value = True Then
            bDatedDir = True
    Else
            bDatedDir = False
    End If
zx = CustomDir(bCusDir, bDatedDir, bUseBoth)
End Sub

Function CustomDir(bCustName As Boolean, bDated As Boolean, bBoth As Boolean)
Dim strName As String, strDate As String
strName = txtCusDir
strDate = GetDateFormat
BaseDir = dirDest.Path
If Not Right(dirDest.Path, 1) = "\" Then
    BaseDir = BaseDir & "\"
End If
If bBoth Then
   txtDest = BaseDir & strName & " - " & strDate
ElseIf bDated And Not bCustName Then
    txtDest = BaseDir & strDate
Else
    txtDest = BaseDir & strName
End If

End Function
Function GetDateFormat()
Dim strSep As String
    Select Case cmbSep.Text
        Case "Space"
            strSep = " "
        Case "-"
            strSep = "-"
        Case Else
            strSep = ""
        End Select
            
GetDateFormat = Format(Date, "mm" & strSep & "dd" & strSep & "yyyy")
End Function
Private Sub chkDays_Click(Index As Integer)
        optDaily.Value = False
End Sub

Private Sub chkServ_Click()
If chkServ Then
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title, App.Path + "\" + App.EXEName + ".exe"
Else
    DeleteSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title
End If
End Sub

Private Sub cmbSep_Change()
Dim zx5
    zx5 = CustomDir(bCusDir, bDatedDir, bUseBoth)
End Sub

Private Sub cmbSep_Click()
Dim zx5
    zx5 = CustomDir(bCusDir, bDatedDir, bUseBoth)
End Sub

Private Sub cmdAddDir_Click()
If chkIncSub.Value = True Then
    Call AddItem(False, True)
Else
    Call AddItem(False)
End If
End Sub

Private Sub cmdAddFile_Click()
    AddItem (True)
End Sub

Private Sub cmdBakNow_Click()
    bBakNow = True
    Call Backup
    bBakNow = False
End Sub

Private Sub cmdClrSoucre_Click()
    lstSource.Clear
End Sub

Private Sub cmdExAddDir_Click()

End Sub

Private Sub cmdExclude_Click()
On Error GoTo erro

    For NLoops = lstSource.ListCount - 1 To 0 Step -1
        If lstSource.Selected(NLoops) Then lstSource.RemoveItem (NLoops)
    Next NLoops
    
    
Saída:
    Exit Sub
    
erro:
    If Err.Number = 68 Then
        MsgBox "The selected drive is not available.", vbCritical
    Else
        MsgBox Err.Number & vbLf & Err.Description, vbCritical
    End If
    Resume Saída

End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdLog_Click()
    ShellExecute hwnd, "open", WindowsDir & "PCBak.log", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub cmdReset_Click()
    txtCusDir = ""
    cmbSep = ""
    txtDest = dirDest.Path
    chkDated = False
    chkCustom = False
    optAnd = False
    optOr = False
End Sub

Private Sub Command1_Click()
    If Not VerifyErrors Then SaveChanges
End Sub

Sub dirDest_Change()
    txtDest.Text = dirDest.Path
End Sub

Private Sub dirSource_Change()
    fileSource.Path = dirSource.Path
End Sub

Private Sub driveDest_Change()
    dirDest.Path = driveDest.Drive
End Sub

Private Sub driveSource_Change()
On Error GoTo erro

    dirSource.Path = driveSource.Drive
        

    Exit Sub
    
erro:
    If Err.Number = 68 Then
        MsgBox "The selected drive is not available.", vbCritical
        driveSource.Drive = "c:"
    Else
        MsgBox Err.Number & vbLf & Err.Description, vbCritical
    End If
    Resume Next
End Sub

Private Sub Form_Activate()
    driveSource.SetFocus
    
        DoEvents
    
    If Not NoIniArchive Then Me.WindowState = vbMinimized

End Sub

Sub Form_Initialize()
If App.PrevInstance Then
    ActivatePrevInstance
End If
End Sub

Private Sub Form_Load()
Call killClose

dirSource.Path = "C:\"
dirDest.Path = "C:\"
Initialize
SSTab1.Tab = 0
With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "PCBackup" & vbNullChar
End With

Shell_NotifyIcon NIM_ADD, nid
SSTab1.TabVisible(4) = False
If lstSource.ListCount > 0 Then 'If there's item(s) in the list check for changes
    recheckItems
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.ScaleMode = vbPixels Then
        Msg = X
    Else
        Msg = X / Screen.TwipsPerPixelX
    End If

    Select Case Msg
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
        Me.WindowState = vbNormal
        result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_RBUTTONUP
        result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu Me.mnuPop
    End Select
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
        Command1_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lblPhone_Click(Index As Integer)
    ShellExecute hwnd, "open", "mailto:wadeg@itekk.com", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub lstSource_Click()
    If lstSource.SelCount >= 1 Then
        cmdExclude.Enabled = True
    Else
        cmdExclude.Enabled = False
    End If
    
End Sub

Private Sub mnuBackup_Click()
    cmdBakNow_Click
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuRestore_Click()
    Me.WindowState = vbNormal
    result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub

Private Sub optAnd_Click()
Dim zx3
    If optAnd Then
        chkDated = True
        bUseBoth = True
    End If
zx3 = CustomDir(bCusDir, bDatedDir, bUseBoth)
End Sub

Private Sub optDaily_Click()
For NLoops = 1 To 7
        chkDays(NLoops).Value = False
    Next NLoops
    
    optDaily.Value = True
End Sub

Private Sub optOr_Click()
Dim zx4
    If optOr Then
        chkDated = True
        bCusDir = False
        bUseBoth = False
    End If
zx4 = CustomDir(bCusDir, bDatedDir, bUseBoth)
End Sub

Private Sub Timer1_Timer()
If Interval = vbEmpty And IniTime = vbEmpty Then Exit Sub
    If Not optDaily Then
        For NLoopsTimer = 1 To 7
            If frmMain.chkDays(NLoopsTimer).Value = True Then If Format(Date, "w") = NLoopsTimer Then CheckTime
        Next NLoopsTimer
    Else
        CheckTime
    End If
End Sub

Private Sub txtCusDir_Change()
Dim zx2
    If txtCusDir <> "" Then
        bCusDir = True
        optAnd.Enabled = True
    Else
        bCusDir = False
        optAnd.Enabled = False
    End If
zx2 = CustomDir(bCusDir, bDatedDir, bUseBoth)
End Sub

Private Sub txtDest_Change()
    If txtDest <> "" Then
        chkCustom.Enabled = True
    Else
        chkCustom.Enabled = False
    End If
End Sub
