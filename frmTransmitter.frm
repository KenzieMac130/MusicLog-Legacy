VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransmitter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Transmitter Power  -  F3"
   ClientHeight    =   7290
   ClientLeft      =   390
   ClientTop       =   915
   ClientWidth     =   10920
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmTransmitter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCopyFeature 
      Caption         =   "C&heck to enable Copy Readings to Screen feature"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   997
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   $"frmTransmitter.frx":058A
      Top             =   3960
      Width           =   3990
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "Copy the Curent &Readings to Screen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1515
      TabIndex        =   35
      ToolTipText     =   "Copies transmitter readings onto the screen."
      Top             =   4200
      Width           =   2955
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3825
      TabIndex        =   73
      Top             =   5565
      Width           =   1650
      Begin VB.CommandButton cmdRestsore 
         Caption         =   "U&ndo Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   855
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   195
         Width           =   705
      End
      Begin VB.CommandButton cmdClearScreen 
         Caption         =   "C&lear Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   90
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   195
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   510
      TabIndex        =   70
      Top             =   5565
      Width           =   1650
      Begin VB.CommandButton cmdRestoreEntries 
         Caption         =   "&Undo Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   855
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   195
         Width           =   705
      End
      Begin VB.CommandButton cmdClearEntries 
         Caption         =   "&Clear Entries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   90
         TabIndex        =   71
         TabStop         =   0   'False
         ToolTipText     =   "Clears Volts and Amps readings from text entry boxes."
         Top             =   195
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transmitter Readings  (Overwrite Existing Reading && Tab)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4470
      Left            =   510
      TabIndex        =   39
      Top             =   225
      Width           =   4950
      Begin VB.Frame Frame5 
         Caption         =   "24 Hour Time Format"
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   630
         Left            =   2955
         TabIndex        =   43
         Top             =   3000
         Width           =   1635
         Begin VB.TextBox txtTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   255
            MaxLength       =   4
            MouseIcon       =   "frmTransmitter.frx":0618
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   8
            Text            =   "Time"
            ToolTipText     =   $"frmTransmitter.frx":0922
            Top             =   233
            Width           =   435
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   825
            TabIndex        =   44
            Top             =   225
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tower Lights"
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   630
         Left            =   315
         TabIndex        =   40
         Top             =   3000
         Width           =   1635
         Begin VB.OptionButton optTwrOff 
            Caption         =   "O&ff"
            Enabled         =   0   'False
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
            Left            =   945
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   233
            Width           =   555
         End
         Begin VB.OptionButton optTwrOn 
            Caption         =   "&On"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   255
            Width           =   570
         End
         Begin VB.Image imgTower 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   240
            Left            =   690
            Picture         =   "frmTransmitter.frx":09BC
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   165
         End
      End
      Begin VB.TextBox txtVmnr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   975
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   780
         Width           =   650
      End
      Begin VB.TextBox txtVrxc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   975
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1305
         Width           =   650
      End
      Begin VB.TextBox txtVgrs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   975
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1845
         Width           =   650
      End
      Begin VB.TextBox txtVgsk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   975
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2370
         Width           =   650
      End
      Begin VB.TextBox txtAmnr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1830
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   780
         Width           =   650
      End
      Begin VB.TextBox txtArxc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1830
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1305
         Width           =   650
      End
      Begin VB.TextBox txtAgrs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1830
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1845
         Width           =   650
      End
      Begin VB.TextBox txtAgsk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1830
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2370
         Width           =   650
      End
      Begin VB.TextBox txtEmnr 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   " Select ""Settings"" menu to change Efficiency values "
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox txtErxc 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   " Select ""Settings"" menu to change Efficiency values "
         Top             =   1305
         Width           =   495
      End
      Begin VB.TextBox txtEgrs 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   " Select ""Settings"" menu to change Efficiency values "
         Top             =   1845
         Width           =   495
      End
      Begin VB.TextBox txtEgsk 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   " Select ""Settings"" menu to change Efficiency values "
         Top             =   2370
         Width           =   495
      End
      Begin VB.TextBox txtVolt1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   13
         Top             =   780
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtVolt2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1305
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtVolt3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1845
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtAmp3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2475
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1845
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtVolt4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1620
         MaxLength       =   1
         TabIndex        =   16
         Top             =   2370
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtAmp4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2475
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2370
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtAmp2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2475
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1305
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtAmp1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   2475
         MaxLength       =   1
         TabIndex        =   17
         Top             =   780
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblLimitWarn 
         Alignment       =   2  'Center
         Caption         =   "Limit Warn Msg Off"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4065
         TabIndex        =   69
         Top             =   375
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblTwr 
         Alignment       =   2  'Center
         Caption         =   "Tower Lights?"
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2100
         TabIndex        =   68
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblgsk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3285
         TabIndex        =   67
         ToolTipText     =   "  Double-Click a station's call letters to bypass readings for that station "
         Top             =   2370
         Width           =   810
      End
      Begin VB.Label lblHL4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   390
         Left            =   4110
         TabIndex        =   66
         Top             =   2355
         Width           =   510
      End
      Begin VB.Label lblPwrLimit4 
         Alignment       =   2  'Center
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3240
         TabIndex        =   65
         Top             =   2640
         Width           =   825
      End
      Begin VB.Label lblPwrLimit3 
         Alignment       =   2  'Center
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3240
         TabIndex        =   64
         Top             =   2110
         Width           =   825
      End
      Begin VB.Label lblPwrLimit2 
         Alignment       =   2  'Center
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3240
         TabIndex        =   63
         Top             =   1580
         Width           =   825
      End
      Begin VB.Label lblPwrLimit1 
         Alignment       =   2  'Center
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3240
         TabIndex        =   62
         Top             =   1050
         Width           =   825
      End
      Begin VB.Label lblmnr 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3285
         TabIndex        =   61
         ToolTipText     =   " Double-Click a station's call letters to bypass readings for that station "
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lblHL1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   390
         Left            =   4110
         TabIndex        =   60
         Top             =   780
         Width           =   510
      End
      Begin VB.Label lblHL3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   390
         Left            =   4110
         TabIndex        =   59
         Top             =   1830
         Width           =   510
      End
      Begin VB.Label lblHL2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   390
         Left            =   4110
         TabIndex        =   58
         Top             =   1305
         Width           =   510
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "  Max. Entry .999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   2460
         TabIndex        =   57
         ToolTipText     =   "An efficiency entry should be a decimal & 2 digits (example .68)"
         Top             =   375
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblStation1 
         Alignment       =   2  'Center
         Caption         =   "WXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         MouseIcon       =   "frmTransmitter.frx":0DFE
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   765
         Width           =   660
      End
      Begin VB.Label lblStation2 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         MouseIcon       =   "frmTransmitter.frx":0F50
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   1335
         Width           =   660
      End
      Begin VB.Label lblStation3 
         Alignment       =   2  'Center
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         MouseIcon       =   "frmTransmitter.frx":10A2
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   1875
         Width           =   660
      End
      Begin VB.Label lblStation4 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         MouseIcon       =   "frmTransmitter.frx":11F4
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "(Flagship)"
         Height          =   225
         Left            =   225
         TabIndex        =   52
         ToolTipText     =   " Double-Click a station's call letters to bypass readings for that station "
         Top             =   990
         Width           =   690
      End
      Begin VB.Label lblWatts 
         Alignment       =   2  'Center
         Caption         =   "Watts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3420
         TabIndex        =   51
         ToolTipText     =   " Click to refresh Min-Max power warnings "
         Top             =   465
         Width           =   540
      End
      Begin VB.Label lblEfficiency 
         Alignment       =   2  'Center
         Caption         =   "Efficiency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2550
         TabIndex        =   50
         ToolTipText     =   " Double-click to change Efficiency values "
         Top             =   465
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Amps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1950
         TabIndex        =   49
         Top             =   465
         Width           =   405
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Volts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1020
         TabIndex        =   48
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblrxc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3285
         TabIndex        =   47
         ToolTipText     =   " Double-Click a station's call letters to bypass readings for that station "
         Top             =   1305
         Width           =   810
      End
      Begin VB.Label lblgrs 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3285
         TabIndex        =   46
         Top             =   1830
         Width           =   810
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Set number of volt-amp left-side characters NOT highlighted when selected"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   60
         TabIndex        =   45
         Top             =   195
         Visible         =   0   'False
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdRestoreDefaults 
      Caption         =   "Restore  &Default Readings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2490
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   " Restores Volt & Amp values, Transmitter Efficiencies, Power Limits, & Station Call Letters to default values  "
      Top             =   5580
      Width           =   1005
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5805
      TabIndex        =   34
      Top             =   5115
      Width           =   4605
      Begin VB.Label lblEntries 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   195
         TabIndex        =   77
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label2 
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   510
         TabIndex        =   76
         Top             =   210
         Width           =   450
      End
      Begin VB.Label lblTimer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   3315
         TabIndex        =   37
         Top             =   195
         Width           =   1020
      End
      Begin VB.Label lblTimerLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Timer Running"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2130
         TabIndex        =   36
         Top             =   195
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5805
      TabIndex        =   23
      Top             =   5670
      Width           =   4605
      Begin VB.CommandButton cmdPrevious 
         Cancel          =   -1  'True
         Caption         =   "&Return to Previous Page F6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   225
         Width           =   2205
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print Screen"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2475
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   225
         Width           =   1050
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3675
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Returns to Planner Page"
         Top             =   225
         Width           =   825
      End
   End
   Begin MSComctlLib.StatusBar staPower 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   22
      Top             =   6960
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   2012
            MinWidth        =   2012
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Is the Num Lock key on?"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15134
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2012
            MinWidth        =   2012
            TextSave        =   "10/12/2017"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9315
      Top             =   180
   End
   Begin VB.ListBox lstXmitter 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   4740
      Left            =   5865
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   300
      Width           =   4485
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Label lblDataMissing 
      Alignment       =   2  'Center
      Caption         =   "Default Data Missing"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2235
      TabIndex        =   31
      Top             =   6435
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblLimits 
      Caption         =   "Transmitter Power Limits are not set. Select ""Settings"" Menu."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   600
      TabIndex        =   30
      Top             =   4785
      Visible         =   0   'False
      Width           =   4725
   End
   Begin VB.Label Label8 
      Caption         =   "Transmitter Efficiencies are not set. Select 'Settings' Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   600
      TabIndex        =   29
      Top             =   5025
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTransmitter.frx":1346
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   975
      TabIndex        =   28
      Top             =   4785
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   $"frmTransmitter.frx":13F3
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   525
      TabIndex        =   27
      Top             =   4755
      Width           =   4920
   End
   Begin VB.Menu mnuPage 
      Caption         =   "P&age"
      Begin VB.Menu mnuPagePlanner 
         Caption         =   "&Music Planning Page"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuPageTimeRemain 
         Caption         =   "&Time Remain Page..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPreviousPage 
         Caption         =   "&Previous Page..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuPageSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageAddTime 
         Caption         =   "&AddTime Calculator..."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuPageSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageStopWatch 
         Caption         =   "&StopWatch..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPageSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingsPrintPage 
         Caption         =   "Print a &Copy of this Page"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSettingsPower 
         Caption         =   "Set or Change Transmitter Power &Limits..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuPageSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingsEffic 
         Caption         =   "Set or Change Transmitter &Efficiency values..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSettingsSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDigitHighlights 
         Caption         =   "&Set the number of Volts/Amps text-entry-box characters (counting from the"
      End
      Begin VB.Menu mnuDigitHighlights2 
         Caption         =   "      LEFT)  which will NOT be highlighted when the text entry box is selected"
      End
      Begin VB.Menu mnuSettingsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingsAdd 
         Caption         =   "&Advanced"
         Begin VB.Menu mnuAdvancedDefault 
            Caption         =   "Save current transmitter readings and efficiency values as &defaults  (access code required)"
         End
         Begin VB.Menu mnuAdvanceSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAdvancePwWarn 
            Caption         =   "&Turn off power limit warning"
         End
         Begin VB.Menu mnuSettingsSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSettingsCall 
            Caption         =   "Enter or change station &call letters (access code required)..."
         End
         Begin VB.Menu mnuAdvSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuShowDefaultPage 
            Caption         =   "Go to &Defaults Page: Set Transmitter Power Limits, Call Letters, Program Average Times && Signature Line... "
            Shortcut        =   {F5}
         End
      End
   End
   Begin VB.Menu mnuHints 
      Caption         =   "User Hin&ts"
   End
End
Attribute VB_Name = "frmTransmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Transmitter Log
    
    Dim miWmnr As Long
    Dim miWrxc As Long
    Dim miWgrs As Long
    Dim miWgsk As Long
    Dim miTime As Integer
    Dim miStation1 As Integer
    Dim miStation2 As Integer
    Dim miStation3 As Integer
    Dim miStation4 As Integer
    Dim miStation As Integer
    Dim mPwrWarn As Boolean
    Dim dDate As Date
    Dim sDate As Date 'date of default transmitter data
    
Option Explicit

Private Sub chkCopyFeature_Click()

    If chkCopyFeature.Value = 1 Then 'copy feature enabled
        cmdList.Enabled = True
        txtTime.Enabled = True
        Frame3.Enabled = True
        Frame5.Enabled = True
        optTwrOn.Enabled = True
        optTwrOn.Value = False
        optTwrOn.ForeColor = vbBlack
        lblTime.ForeColor = &H80&
        optTwrOff.Enabled = True
        optTwrOff.Value = False
        lblEntries.Visible = True
        Label2.Visible = True
        Frame5.Enabled = True

     ElseIf chkCopyFeature.Value = 0 Then 'copy feature disabled
        cmdList.Enabled = False
        txtTime.Enabled = False
        Frame3.Enabled = False
        optTwrOn.Value = 0
        optTwrOn.Enabled = False
        optTwrOff.Value = 0
        optTwrOff.Enabled = False
        lblEntries.Visible = False
        Label2.Visible = False
        lblTime.ForeColor = &H808080
        lblTwr.Visible = False
        imgTower.Visible = False
        Frame5.Enabled = False

    End If
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub cmdClearEntries_Click()
    'clears but saves volt/amp entries
    If txtVmnr <> "" And txtAmnr <> "" Then
       Open "ReadingsL.dat" For Output As #15
            Write #15, txtVmnr, txtVrxc, txtVgrs, txtVgsk, txtAmnr, txtArxc, _
            txtAgrs, txtAgsk, txtEmnr, txtErxc, txtEgrs, txtEgsk, _
            txtVolt1, txtVolt2, txtVolt3, txtVolt4, txtAmp1, txtAmp2, txtAmp3, txtAmp4
        Close #15
    End If
    cmdRestoreEntries.Caption = "&Undo Clear"
    txtVmnr = ""
    txtVrxc = ""
    txtVgrs = ""
    txtVgsk = ""
    txtAmnr = ""
    txtArxc = ""
    txtAgrs = ""
    txtAgsk = ""
    lblTwr.Visible = False
    imgTower.Visible = False
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub cmdList_click()

    If txtVmnr = "" And txtVrxc = "" And txtVgrs = "" And txtVgsk = "" Then
    Dim iResponse As Integer
      iResponse = MsgBox("There are no readings to copy.", vbOKOnly, "No Data")
      If txtVmnr.Enabled = True Then
          txtVmnr.SetFocus
      End If
      Exit Sub
    End If

    Dim iError As String 'allows 24:00 hours
    iError = txtTime
    Dim iErrorSec As String 'prevents seconds entry greater than 59 seconds
    Dim iErrorMin As String 'prevents hours entry greater than 24 hours
    
    If IsNumeric(txtTime) And txtTime <> "Time" Then
        iErrorSec = Right(txtTime, 2)
        iErrorMin = Left(txtTime, 2)
    End If

    If Val(iErrorSec) > 59 Then
        lblTime = "Error"
        MsgBox "Seconds entry exceeds 59 seconds", vbOKOnly, "Seconds Error"
        txtTime.Visible = True
        txtTime.SetFocus
        txtTime.SelStart = 0 'begin selection at start
        txtTime.SelLength = 4 'Len(txtTime) 'selects # of characters

    ElseIf Len(txtTime) = 4 And iError <> "2400" And Val(iErrorMin) > 23 Then
        lblTime = "Error"
        MsgBox "Hours entry exceeds 24:00 Hours", vbOKOnly, "Hours Error"
        txtTime.Visible = True
        txtTime.Visible = True
        txtTime.SetFocus
        txtTime.SelStart = 0 'begin selection at start
        txtTime.SelLength = 4 'Len(txtTime) 'selects # of characters
    Else
        If optTwrOn = False And optTwrOff = False Then 'tower lights query
           iResponse = MsgBox("Are The Tower Lights OFF?", vbYesNo, "Tower Lights")
               If iResponse = vbYes Then
                   optTwrOff.Value = True
               Else
                   optTwrOn.Value = True
               End If
        End If

        Dim iTwr As String
        If optTwrOn = True Then
        iTwr = "Tower Lights ON"
    End If
    '----------
    Dim mnrError As String
    Dim rxcError As String
    Dim grsError As String
    Dim gskError As String
    
    If lblmnr.ForeColor = vbRed Then
        mnrError = " *"
    Else: mnrError = ""
    End If
    
    If lblrxc.ForeColor = vbRed Then
        rxcError = " *"
    Else: rxcError = ""
    End If
    
    If lblgrs.ForeColor = vbRed Then
        grsError = " *"
    Else: grsError = ""
    End If
    
    If lblgsk.ForeColor = vbRed Then
        gskError = " *"
    Else: gskError = ""
    End If
'--------

    lstXmitter.AddItem lblTime & "       " & iTwr
    If lblmnr <> "" Then
        lstXmitter.AddItem "  " & lblStation1 & "   Volts:  " & txtVmnr & "    Amps:  " & txtAmnr & "    Watts:  " & lblmnr & mnrError
        
        If mnrError = " *" Then
            If miWmnr < Val(gStation1Min) Then
                lstXmitter.AddItem "      ---" & lblStation1 & " power below normal minimum of " & gStation1Min & " watts"
            
            ElseIf miWmnr > Val(gStation1Max) Then
                lstXmitter.AddItem "      ---" & lblStation1 & " power above normal maximum of " & gStation1Max & " watts"
            End If
        End If
    End If
    
    If lblrxc <> "" Then
        lstXmitter.AddItem "  " & lblStation2 & "   Volts:  " & txtVrxc & "    Amps:  " & txtArxc & "    Watts:  " & lblrxc & rxcError
        If rxcError = " *" Then
            If miWrxc < Val(gStation2Min) Then
                lstXmitter.AddItem "      ---" & lblStation2 & " power below normal minimum of " & gStation2Min & " watts"
            
            ElseIf miWrxc > Val(gStation2Max) Then
                lstXmitter.AddItem "      ---" & lblStation2 & " power above normal maximum of " & gStation2Max & " watts"
            End If
        End If
    End If
    
    If lblgrs <> "" Then
        lstXmitter.AddItem "  " & lblStation3 & "   Volts:  " & txtVgrs & "    Amps:  " & txtAgrs & "    Watts:  " & lblgrs & grsError
        If grsError = " *" Then
            If miWgrs < Val(gStation3Min) Then
                lstXmitter.AddItem "      ---" & lblStation3 & " power below normal minimum of " & gStation3Min & " watts"
            
            ElseIf miWgrs > Val(gStation3Max) Then
                lstXmitter.AddItem "      ---" & lblStation3 & " power above normal maximum of " & gStation3Max & " watts"
            End If
        End If
    End If
    
    If lblgsk <> "" Then
        lstXmitter.AddItem "  " & lblStation4 & "   Volts:  " & txtVgsk & "    Amps:  " & txtAgsk & "    Watts:  " & lblgsk & gskError
        If gskError = " *" Then
            If miWgsk < Val(gStation4Min) Then
                lstXmitter.AddItem "      ---" & lblStation4 & " power below normal minimum of " & gStation4Min & " watts"
            
            ElseIf miWgsk > Val(gStation4Max) Then
                lstXmitter.AddItem "      ---" & lblStation4 & " power above normal maximum of " & gStation4Max & " watts"
            End If
        End If
    End If
  
    lstXmitter.AddItem ""
'-----
    Dim iListCount As Integer

    Dim iMnr As Integer
    Dim iRxc As Integer
    Dim iGrs As Integer
    Dim iGsk As Integer
    Dim iStaCount As Integer

    If lblmnr <> "" Then
        iMnr = 1
    Else
        iMnr = 0
    End If
    
    If lblrxc <> "" Then
        iRxc = 1
    Else
        iRxc = 0
    End If
    
    If lblgrs <> "" Then
        iGrs = 1
    Else
        iGrs = 0
    End If

    If lblgsk <> "" Then
        iGsk = 1
    Else
        iGsk = 0
    End If

    iStaCount = iMnr + iRxc + iGrs + iGsk

    Select Case iStaCount
    Case 0
        iListCount = 0
    Case 1
        iListCount = (lstXmitter.ListCount - 1) / 3
    Case 2
        iListCount = (lstXmitter.ListCount - 1) / 4
    Case 3
        iListCount = (lstXmitter.ListCount - 1) / 5
    Case 4
        iListCount = (lstXmitter.ListCount - 1) / 6
    Case Else
        iListCount = 0
    End Select
   
   lblEntries = iListCount 'displays number of entries
   
   If lstXmitter.ListCount > 2 Then
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
    End If
  '----------
 
    If lblEntries = "1" Then
        Label2.Caption = "Entry"
    Else
        Label2.Caption = "Entries"
    End If

    If lstXmitter.ListCount > 20 Then 'scrolls screen
        lstXmitter.TopIndex = lstXmitter.ListCount - 19
    End If

    Open "PwrL.dat" For Output As #14
    Dim i As Integer
    For i = 0 To lstXmitter.ListCount - 1
        Print #14, lstXmitter.List(i)
    Next
    Close #14

    txtTime = "Time?"
    lblTwr.Visible = False
    
    If optTwrOn.Value = False Then
        imgTower.Visible = False
    ElseIf optTwrOn.Value = True Then
        imgTower.Visible = True
    End If
    
    txtTime.Visible = True
    Label1.Visible = False
    Label11.Visible = True
    
    If lstXmitter.ListCount > 4 Then 'option to print transmitter readings with music log
        frmPlanner!chkPrintXmitter.Visible = True
    End If
nosave:
    miTime = 0
        If txtVmnr.Enabled = True Then
            txtVmnr.SetFocus
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim iResponse As Integer
     
    If lstXmitter.ListCount <= 4 Then 'if there are no listings, nothing to print message
      iResponse = MsgBox("There is no transmitter data on the screen to print.", vbOKOnly + vbExclamation, "No Data")
      Exit Sub
    Else
    
    iResponse = MsgBox("Print a copy of the transmitter meter readings?", vbYesNo, "Confirm")
        If iResponse = vbYes Then
            
            Dim iIndex As Integer
            Printer.Print
            Printer.Print
            Printer.FontName = "Arial"
            Printer.FontSize = 12
            Printer.Print Tab(8); "Transmitter Readings"   ' & Format(dDate, "; Short; Date; ")"
            Printer.FontSize = 10
            For iIndex = 0 To lstXmitter.ListCount - 1
                Printer.Print Tab(9); lstXmitter.List(iIndex)
            Next iIndex
            Printer.FontSize = 9
            Printer.Print Tab(12); Format(dDate, "Short Date")
            Printer.EndDoc
        End If
    End If
End Sub

Private Sub CmdClearScreen_Click()
    lstXmitter.Clear 'clears screen
    lblEntries = ""
    cmdPrint.Enabled = False
    Label2 = ""
    lblTwr.Visible = False
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
    giChkPrintXmitter = 1
    frmPlanner!chkPrintXmitter.Value = 0
    frmPlanner!chkPrintXmitter.Visible = False
    
    dDate = Now()
    lstXmitter.AddItem "  " & Format(dDate, "Long Date")
    lstXmitter.AddItem ""
End Sub

Private Sub cmdRestoreEntries_Click()
    Dim Vmnr As String
    Dim Vrxc As String
    Dim Vgrs As String
    Dim Vgsk As String
    Dim Amnr As String
    Dim Arxc As String
    Dim Agrs As String
    Dim Agsk As String
    Dim Emnr As String
    Dim Erxc As String
    Dim Egrs As String
    Dim Egsk As String
    
    Dim Volt1, Volt2, Volt3, Volt4, Amp1, Amp2, Amp3, Amp4 As String
    
On Error GoTo HandleErrors

    Open "ReadingsL.dat" For Input As #15
        Input #15, Vmnr, Vrxc, Vgrs, Vgsk, Amnr, Arxc, Agrs, Agsk, Emnr, Erxc, Egrs, Egsk, _
        Volt1, Volt2, Volt3, Volt4, Amp1, Amp2, Amp3, Amp4
    Close #15
    
    cmdRestoreEntries.Caption = "&Undo Clear"
    txtVolt1 = Volt1
    txtVolt2 = Volt2
    txtVolt3 = Volt3
    txtVolt4 = Volt4
    txtAmp1 = Amp1
    txtAmp2 = Amp2
    txtAmp3 = Amp3
    txtAmp4 = Amp4

    txtVmnr = Vmnr
    txtVrxc = Vrxc
    txtVgrs = Vgrs
    txtVgsk = Vgsk
    txtAmnr = Amnr
    txtArxc = Arxc
    txtAgrs = Agrs
    txtAgsk = Agsk
    txtEmnr = Emnr
    txtErxc = Erxc
    txtEgrs = Egrs
    txtEgsk = Egsk
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
    Exit Sub
    
HandleErrors:
    Close #15
End Sub

Private Sub cmdRestsore_Click()
 
    lstXmitter.Clear
    Dim iResponse As Integer
    Dim sTemp3 As String
    Dim iListCount As Integer
  
On Error GoTo HandleErrors
    Open "PwrL.dat" For Input As #14
        Do Until EOF(14)
        Line Input #14, sTemp3
        lstXmitter.AddItem sTemp3
    Loop
    Close #14
 
    iListCount = (lstXmitter.ListCount) / 6
    lblEntries = iListCount 'counts number of entries
    
    If lstXmitter.ListCount > 2 Then
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
    End If
    
    If lblEntries = "1" Then
        Label2.Caption = "Entry"
    Else
        Label2.Caption = "Entries"
    End If

    If lstXmitter.ListCount > 19 Then 'scrolls screen
        lstXmitter.TopIndex = lstXmitter.ListCount - 18
    End If
    
    If lstXmitter.ListCount > 3 Then
        frmPlanner!chkPrintXmitter.Visible = True
    End If

    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
    Exit Sub
   
HandleErrors:
    iResponse = MsgBox("No transmitter entries have been copied to the screen.", vbOKOnly, "NoData")
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
  Close #14
End Sub

Private Sub cmdExit_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
        mnuDigitHighlights_Click
    End If
    
    giClockShow = 3
    frmPlanner.Show
    frmTransmitter.Hide
End Sub

Private Sub cmdPrevious_Click()

  If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
        mnuDigitHighlights_Click
    End If
 
    If giClockShow <> 0 Then
        Select Case giClockShow
            Case 4
                frmPlanner.Show
            Case 5
                frmTimeRemain.Show
            Case 6
                frmDefaults.Show
            Case Else
                frmPlanner.Show
        End Select
        frmTransmitter.Hide
        giClockShow = 3
    Else
        frmPlanner.Show
        frmTransmitter.Hide
        giClockShow = 3
    End If
End Sub

Private Sub cmdRestoreDefaults_Click()

    Dim Vmnr As String
    Dim Vrxc As String
    Dim Vgrs As String
    Dim Vgsk As String
    Dim Amnr As String
    Dim Arxc As String
    Dim Agrs As String
    Dim Agsk As String
    Dim Emnr As String
    Dim Erxc As String
    Dim Egrs As String
    Dim Egsk As String

    Dim Volt1, Volt2, Volt3, Volt4, Amp1, Amp2, Amp3, Amp4 As String

    Dim Station1, Station2, Station3, Station4 As String
    Dim Station1Min, Station1Max, Station2Min, Station2Max, Station3Min, _
        Station3Max, Station4Min, Station4Max As String
    Dim PlanTime, IntroOut, sClose, Spot As String

    Dim rDate As Date 'records the date default data is current
    rDate = Now

    '-------------------
On Error GoTo HandleErrors:

    Open "DefaultStation.dat" For Input As #20 'default data
        Input #20, Station1, Station2, Station3, Station4
    Close #20

     Open "DefaultMinMax.dat" For Input As #24
        Input #24, Station1Min, Station1Max, Station2Min, Station2Max, _
        Station3Min, Station3Max, Station4Min, Station4Max
    Close #24

    Open "DefaultReadings.dat" For Input As #16
        Input #16, Vmnr, Vrxc, Vgrs, Vgsk, Amnr, Arxc, Agrs, Agsk, Emnr, Erxc, Egrs, Egsk, _
        Volt1, Volt2, Volt3, Volt4, Amp1, Amp2, Amp3, Amp4, Emnr, Erxc, Egrs, Egsk
    Close #16

    Open "DefaultDate.dat" For Input As #17
        Input #17, sDate
    Close #17
    '--------
    Dim iResponse As Integer
    iResponse = MsgBox("If entry data or settings are lost or corrupted, this command will restore Volt, Amp," & _
    vbCrLf & "Transmitter Efficiency, and Power Limit values and Station Call Letters to the readings" _
    & vbCrLf & "current as of " & Format(sDate, "Long Date") & "." _
    & vbCrLf & vbCrLf & "Click 'OK' to set station and transmitter entries to the values of " & _
    Format(sDate, "Long Date") & ".", vbOKCancel + vbInformation, "Set Entries to the Default Values")

    If iResponse = vbCancel Then
        Exit Sub
    Else

        If txtVmnr <> "" And txtAmnr <> "" Then 'to save existing data
        Open "ReadingsL.dat" For Output As #15
            Write #15, txtVmnr, txtVrxc, txtVgrs, txtVgsk, txtAmnr, txtArxc, _
            txtAgrs, txtAgsk, txtEmnr, txtErxc, txtEgrs, txtEgsk, _
            txtVolt1, txtVolt2, txtVolt3, txtVolt4, txtAmp1, txtAmp2, txtAmp3, txtAmp4
        Close #15
        cmdRestoreEntries.Caption = "&Undo Default"
        End If

        'save defaults as standard values
        Open "Stations.dat" For Output As #18
        Write #18, Station1, Station2, Station3, Station4
        Close #18

        Open "MinMax.dat" For Output As #19
        Write #19, Station1Min, Station1Max, Station2Min, Station2Max, _
                Station3Min, Station3Max, Station4Min, Station4Max
        Close #19

        '---------
        lblLimits.Visible = False

        If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
            mnuSettingsEffic_Click
        End If

        If mnuDigitHighlights.Checked = True Then 'closes set highlighted digits
            mnuDigitHighlights_Click
        End If

        frmDefaults!txtStation1Min = Station1Min
        frmDefaults!txtStation1Max = Station1Max
        frmDefaults!txtStation2Min = Station2Min
        frmDefaults!txtStation2Max = Station2Max
        frmDefaults!txtStation3Min = Station3Min
        frmDefaults!txtStation3Max = Station3Max
        frmDefaults!txtStation4Min = Station4Min
        frmDefaults!txtStation4Max = Station4Max

        giStation1 = UCase(Station1) 'use global giStation1 so planner page also gets change

        frmDefaults!txtStation1 = giStation1 'UCase(Station1)
        frmDefaults!txtStation2 = Station2
        frmDefaults!txtStation3 = Station3
        frmDefaults!txtStation4 = Station4

        lblStation1.ToolTipText = lblStation1 & " Minimum power: " & frmDefaults!txtStation1Min & " watts,  Maximum power: " & frmDefaults!txtStation1Max & " watts"
        lblStation2.ToolTipText = lblStation2 & " Minimum power: " & frmDefaults!txtStation2Min & " watts,  Maximum power: " & frmDefaults!txtStation2Max & " watts"
        lblStation3.ToolTipText = lblStation3 & " Minimum power: " & frmDefaults!txtStation3Min & " watts,  Maximum power: " & frmDefaults!txtStation3Max & " watts"
        lblStation4.ToolTipText = lblStation4 & " Minimum power: " & frmDefaults!txtStation4Min & " watts,  Maximum power: " & frmDefaults!txtStation4Max & " watts"

        gStation1Min = Station1Min
        gStation1Max = Station1Max
        gStation2Min = Station2Min
        gStation2Max = Station2Max
        gStation3Min = Station3Min
        gStation3Max = Station3Max
        gStation4Min = Station4Min
        gStation4Max = Station4Max

        lblStation1 = frmDefaults!txtStation1
        lblStation2 = frmDefaults!txtStation2
        lblStation3 = frmDefaults!txtStation3
        lblStation4 = frmDefaults!txtStation4
        
        If lblStation1 <> "" Then
            
            lblStation1.Enabled = True
            lblStation1.ForeColor = &H80000012
            lblmnr.Visible = True
            lblmnr = Format(miWmnr, "#,###")
            lblmnr.BackColor = &H80000018
            
            txtVmnr.Enabled = True
            txtAmnr.Enabled = True
            txtEmnr.Enabled = True

            txtVmnr.BackColor = &H80000005 'white
            txtAmnr.BackColor = &H80000005 'white
            txtEmnr.BackColor = &H80000018 ' yellow
            txtEmnr.ForeColor = &H80000008  'black
        End If

        If lblStation2 <> "" Then
            
            lblStation2.Enabled = True
            lblStation2.ForeColor = &H80000012
            
            lblrxc.Visible = True
            lblrxc = Format(miWrxc, "#,###")
            lblrxc.BackColor = &H80000018
            
            txtVrxc.Enabled = True
            txtArxc.Enabled = True
            txtErxc.Enabled = True
            
            txtVrxc.BackColor = &H80000005 'white
            txtArxc.BackColor = &H80000005 'white
            txtErxc.BackColor = &H80000018  'yellow '&HE0E0E0    'gray
            txtErxc.ForeColor = &H80000008  'black
        End If
        
        If lblStation3 <> "" Then
            
            lblStation3.Enabled = True
            lblStation3.ForeColor = &H80000012
            
            lblgrs.Visible = True
            lblgrs = Format(miWgrs, "#,###")
            lblgrs.BackColor = &H80000018
            
            txtVgrs.BackColor = &H80000005 'white
            txtAgrs.BackColor = &H80000005 'white
            txtEgrs.BackColor = &H80000018  'yellow
            txtEgrs.ForeColor = &H80000008  'black
            
            txtVgrs.Enabled = True
            txtAgrs.Enabled = True
            txtEgrs.Enabled = True
        End If
        
        If lblStation4 <> "" Then
            
            lblStation4.Enabled = True
            lblStation4.ForeColor = &H80000012

            lblgsk.Visible = True
            lblgsk = Format(miWgsk, "#,###")
            lblgsk.BackColor = &H80000018
            
            txtVgsk.BackColor = &H80000005 'white
            txtAgsk.BackColor = &H80000005 'white
            txtEgsk.BackColor = &H80000018  'yellow
            txtEgsk.ForeColor = &H80000008  'black
            
            txtVgsk.Enabled = True
            txtAgsk.Enabled = True
            txtEgsk.Enabled = True
        End If
                      
        '--------

        txtVolt1 = Volt1
        txtVolt2 = Volt2
        txtVolt3 = Volt3
        txtVolt4 = Volt4

        txtAmp1 = Amp1
        txtAmp2 = Amp2
        txtAmp3 = Amp3
        txtAmp4 = Amp4

        txtVmnr = Vmnr
        txtVrxc = Vrxc
        txtVgrs = Vgrs
        txtVgsk = Vgsk

        txtAmnr = Amnr
        txtArxc = Arxc
        txtAgrs = Agrs
        txtAgsk = Agsk
        txtEmnr = Emnr

        txtErxc = Erxc
        txtEgrs = Egrs
        txtEgsk = Egsk

        If txtVmnr.Enabled = True Then
            txtVmnr.SetFocus
        End If

        frmDefaults!ckIdeal1.Value = 0
        frmDefaults!ckIdeal2.Value = 0
        frmDefaults!ckIdeal3.Value = 0
        frmDefaults!ckIdeal4.Value = 0

 '----------
        Dim iMin, iMax, iIdeal1, iIdeal2, iIdeal3, iIdeal4, iCk1, iCk2, iCk3, iCk4, iHourNow As Integer

        Open "DefaultXmitterIdeal.dat" For Input As #26
        Input #26, iMin, iMax, iIdeal1, iIdeal2, iIdeal3, iIdeal4, iCk1, iCk2, iCk3, iCk4
        Close #26

        frmDefaults!txtMin = iMin
        frmDefaults!txtMax = iMax
        frmDefaults!txtIdeal1 = iIdeal1
        frmDefaults!txtIdeal2 = iIdeal2
        frmDefaults!txtIdeal3 = iIdeal3
        frmDefaults!txtIdeal4 = iIdeal4
        frmDefaults!ckIdeal1.Value = iCk1
        frmDefaults!ckIdeal2.Value = iCk2
        frmDefaults!ckIdeal3.Value = iCk3
        frmDefaults!ckIdeal4.Value = iCk4

    End If

    MsgBox "Station and transmitter data have been restored to the values current as of " & Format(sDate, "Long Date") & "." & vbCrLf & vbCrLf _
    & "Check the restored Efficiency and Transmitter Power Limits data with the station transmitter log to be" _
    & vbCrLf & "certain the restored data is still current. Update if necessary. " _
    & vbCrLf & vbCrLf & "Pause the mouse cursor over each station's call letters for a readout of minimum and maximum power limits" _
    & vbCrLf & "for that station. Use the 'Settings' menu to update Efficiency or Power Limit values." & vbCrLf & vbCrLf & _
    "If any stations have been selected for bypass, the call letters must be double-clicked to reactivate." & vbCrLf & vbCrLf & _
    "(Note: To restore transmitter readings and station data to previous values, click the 'Undo Default' button.)", _
    vbOKOnly, "Check restored Efficiency & Power Limits against station transmitter log for currency"

    staPower.Panels(2).Text = "Default values current as of " & Format(rDate, "Long Date") & "."

Exit Sub

HandleErrors:

    MsgBox "Transmitter default data is missing or incomplete.  (1)  Enter at least the Flagship station's Volts, Amps and Effriciency values." & vbCrLf & vbCrLf & _
    "(2)  Set Transmitter Power Limits:" & vbCrLf & _
    "       From the 'Settings' menu select  'Set or Change Transmitter Power Limits...'  to open the Defaults page where Power Limits can be set." & vbCrLf & vbCrLf & _
    "(3)  Save current readings as Defaults:" & vbCrLf & _
    "       From the 'Settings' menu, select  'Advanced',  then select  'Save Station Current Values & Readings as Defaults'." _
    & vbCrLf & vbCrLf & "Transmitter power limits and efficiency values can be obtained from the station's transmitter log.", _
    vbOKOnly, "Default Data Missing or Incomplete"
        
    staPower.Panels(2).Text = "Transmitter default values have not been entered. Select save current values as defaults from Settings menu"
    
    lblDataMissing.Caption = "Default Data Missing:"
    lblDataMissing.Visible = True
    
    Close #15
    Close #16
    Close #17
    Close #18
    Close #19
    Close #20
    Close #24
    Close #26
    
End Sub

Private Sub Form_Activate()

    If lblStation1 <> "" And (txtVmnr <> "" Or txtAmnr <> "") And txtEmnr = "" Then
        MsgBox "The amount of power delivered from the transmitter to the antenna is dependent on the efficiency of the system." & vbCrLf & vbCrLf & _
        "Efficiency ratings normally range from a low of about  .50 (50%) to a high of about .90 (90%)." & vbCrLf & _
        "In all cases, efficiencies will be less than .99 (99%)." & vbCrLf & vbCrLf & "Enter the transmitter efficiency rating as a decimal number." _
        & vbCrLf & vbCrLf & "Efficiency rating information normally can be found on the station's transmitter log.", _
        vbOKOnly, "Enter Transmitter Efficiency Rating"
        
        mnuSettingsEffic_Click
        txtEmnr.SetFocus
        Exit Sub
    End If

    If frmDefaults!txtStation1Min = "" Or frmDefaults!txtStation1Max = "" Then
        lblLimits.Visible = True
        Label1.Visible = False
        Label11.Visible = False
    Else
        lblLimits.Visible = False
    End If
    
    If txtEmnr = "" Then
        Label8.Visible = True
        Label1.Visible = False
        Label11.Visible = False
    Else
        Label8.Visible = False
    End If
    
    miStation4 = 0

    If frmDefaults!txtStation1Min.Text <> "" Or frmDefaults!txtStation1Max.Text <> "" Then
       lblStation1.ToolTipText = lblStation1 & " Minimum power: " & frmDefaults!txtStation1Min & " watts,  Maximum power: " & frmDefaults!txtStation1Max & " watts"
    Else
        If lblStation1 <> "" Then
            lblStation1.ToolTipText = "Power Limits not set"
        Else
            lblStation1.ToolTipText = ""
        End If
        
    End If
    
    If frmDefaults!txtStation2Min.Text <> "" Or frmDefaults!txtStation2Max.Text <> "" Then
        lblStation2.ToolTipText = lblStation2 & " Minimum power: " & frmDefaults!txtStation2Min & " watts,  Maximum power: " & frmDefaults!txtStation2Max & " watts"
    Else
        If lblStation2 <> "" Then
            lblStation2.ToolTipText = "Power Limits not set"
        Else
            lblStation2.ToolTipText = ""
        End If
    End If
        
    If frmDefaults!txtStation3Min.Text <> "" Or frmDefaults!txtStation3Max.Text <> "" Then
        lblStation3.ToolTipText = lblStation3 & " Minimum power: " & frmDefaults!txtStation3Min & " watts,  Maximum power: " & frmDefaults!txtStation3Max & " watts"
     Else
        If lblStation3 <> "" Then
            lblStation3.ToolTipText = "Power Limits not set"
        Else
            lblStation3.ToolTipText = ""
        End If
    End If
        
     If frmDefaults!txtStation4Min.Text <> "" Or frmDefaults!txtStation4Max.Text <> "" Then
        lblStation4.ToolTipText = lblStation4 & " Minimum power: " & frmDefaults!txtStation4Min & " watts,  Maximum power: " & frmDefaults!txtStation4Max & " watts"
     Else
        If lblStation4 <> "" Then
            lblStation4.ToolTipText = "Power Limits not set"
        Else
            lblStation4.ToolTipText = ""
        End If
    End If

End Sub

Private Sub Form_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
   
   If mnuDigitHighlights.Checked = True Then
        mnuDigitHighlights_Click
   End If

   If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
    
End Sub

Private Sub Form_Deactivate()
    If mnuSettingsEffic.Checked = True Then
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
      mnuDigitHighlights_Click
    End If
End Sub

Private Sub Form_Load()
    Dim iResponse As Integer
    Dim Vmnr As String
    Dim Vrxc As String
    Dim Vgrs As String
    Dim Vgsk As String
    Dim Amnr As String
    Dim Arxc As String
    Dim Agrs As String
    Dim Agsk As String
    Dim Emnr As String
    Dim Erxc As String
    Dim Egrs As String
    Dim Egsk As String
    Dim Volt1, Volt2, Volt3, Volt4, Amp1, Amp2, Amp3, Amp4 As String
    
    Dim Station1 As String
    Dim Station2 As String
    Dim Station3 As String
    Dim Station4 As String
    Dim iCopyFeature As Integer
    Dim iTimeNote As Integer
    
     Dim Station1Min, Station1Max, Station2Min, Station2Max, Station3Min, _
        Station3Max, Station4Min, Station4Max As String
    
    dDate = Now()
    lstXmitter.AddItem "  " & Format(dDate, "Long Date")
    lstXmitter.AddItem ""
    
On Error GoTo HandleErrors
    Open "Stations.dat" For Input As #18
    Input #18, Station1, Station2, Station3, Station4
    Close #18
    
    lblStation1 = Station1
    lblStation2 = Station2
    lblStation3 = Station3
    lblStation4 = Station4
    
    Open "ReadingsL.dat" For Input As #15
    Input #15, Vmnr, Vrxc, Vgrs, Vgsk, Amnr, Arxc, Agrs, Agsk, Emnr, Erxc, Egrs, Egsk, _
    Volt1, Volt2, Volt3, Volt4, Amp1, Amp2, Amp3, Amp4
    Close #15
 
 '------station1 (required)-----
 
    txtVolt1 = Volt1
    txtAmp1 = Amp1
    txtVmnr = Vmnr
    txtAmnr = Amnr
    txtEmnr = Emnr
'------------station2-----

    If lblStation2 <> "" Then
        txtVolt2 = Volt2
        txtAmp2 = Amp2
        txtVrxc = Vrxc
        txtArxc = Arxc
        txtErxc = Erxc
    Else
        frmTransmitter!txtVrxc = ""
        frmTransmitter!txtArxc = ""
        frmTransmitter!txtErxc = ""
        frmTransmitter!lblrxc.Visible = False
        frmTransmitter!lblPwrLimit2.Visible = False
        frmDefaults!txtStation2Min = ""
        frmDefaults!txtStation2Max = ""
    
        txtVrxc.Enabled = False
        txtArxc.Enabled = False
        txtVrxc.BackColor = &HE0E0E0 'gray
        txtArxc.BackColor = &HE0E0E0 'gray
    End If

'-----------station3--------

    If lblStation3 <> "" Then
        txtVolt3 = Volt3
        txtAmp3 = Amp3
        txtVgrs = Vgrs
        txtAgrs = Agrs
        txtEgrs = Egrs
    Else
        
        frmTransmitter!txtVgrs = ""
        frmTransmitter!txtAgrs = ""
        frmTransmitter!txtEgrs = ""
        frmTransmitter!lblgrs.Visible = False
        frmTransmitter!lblPwrLimit2.Visible = False
        frmDefaults!txtStation2Min = ""
        frmDefaults!txtStation2Max = ""
    
        txtVgrs.Enabled = False
        txtAgrs.Enabled = False
        txtVgrs.BackColor = &HE0E0E0 'gray
        txtAgrs.BackColor = &HE0E0E0 'gray
    End If

'----------station4---------

    If lblStation4 <> "" Then
        txtVolt4 = Volt4
        txtAmp4 = Amp4
        txtVgsk = Vgsk
        txtAgsk = Agsk
        txtEgsk = Egsk
    Else
        
        frmTransmitter!txtVgsk = ""
        frmTransmitter!txtAgsk = ""
        frmTransmitter!txtEgsk = ""
        frmTransmitter!lblgsk.Visible = False
        frmTransmitter!lblPwrLimit2.Visible = False
        frmDefaults!txtStation2Min = ""
        frmDefaults!txtStation2Max = ""
    
        txtVgsk.Enabled = False
        txtAgsk.Enabled = False
        txtVgsk.BackColor = &HE0E0E0 'gray
        txtAgsk.BackColor = &HE0E0E0 'gray
    End If
    
    Open "DefaultDate.dat" For Input As #17
    Input #17, sDate
    Close #17
    staPower.Panels(2).Text = "Default Entries are the station list and transmitter values current as of " & Format(sDate, "Long Date") & "."
    
    If txtVmnr.Enabled = True Then
        frmTransmitter!txtVmnr.SetFocus
    End If
    Exit Sub
    
HandleErrors:
    Close #15
    Close #18
    Close #17
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'upon closeout, saves settings
On Error GoTo HandleErrors

'if flagship station contains, then saves all volt/amps readings
    If txtVmnr <> "" And txtAmnr <> "" Then
        Open "ReadingsL.dat" For Output As #15
            Write #15, txtVmnr, txtVrxc, txtVgrs, txtVgsk, txtAmnr, txtArxc, _
            txtAgrs, txtAgsk, txtEmnr, txtErxc, txtEgrs, txtEgsk, _
            txtVolt1, txtVolt2, txtVolt3, txtVolt4, txtAmp1, txtAmp2, txtAmp3, txtAmp4
        Close #15
    End If

HandleErrors:
    frmPlanner.Show
End Sub

Private Sub Frame1_Click()
     If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
         mnuSettingsEffic_Click
     End If
    
    If mnuDigitHighlights.Checked = True Then
      mnuDigitHighlights_Click
    End If
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub Frame2_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
   
   If mnuDigitHighlights.Checked = True Then
     mnuDigitHighlights_Click
   End If
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub lblEfficiency_DblClick()
    mnuSettingsEffic_Click
End Sub

Private Sub lblgrs_DblClick()
    If txtVgrs.Enabled = True Then
        txtVgrs.SetFocus
    End If
End Sub

Private Sub lblgsk_DblClick()
    If txtVgsk.Enabled = True Then
        txtVgsk.SetFocus
    End If
End Sub

Private Sub lblLimitWarn_DblClick()
    mnuAdvancePwWarn.Checked = False
    lblLimitWarn.Visible = False
    miStation1 = 0
    miStation2 = 0
    miStation3 = 0
    miStation4 = 0
    miStation = 0 'rearms 'bypass station' message
End Sub

Private Sub lblmnr_DblClick()
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub lblrxc_DblClick()
    If txtVrxc.Enabled = True Then
        txtVrxc.SetFocus
    End If
End Sub

Private Sub lblStation1_Click()
    If lblStation1.Enabled = True And txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub lblStation2_Click()
    If lblStation2.Enabled = True And txtVrxc.Enabled = True Then
        txtVrxc.SetFocus
    End If
End Sub

Private Sub lblStation3_Click()
    If lblStation3.Enabled = True And txtVgrs.Enabled = True Then
        txtVgrs.SetFocus
    End If
End Sub

Private Sub lblStation4_Click()
    If lblStation4.Enabled = True And txtVgsk.Enabled = True Then
        txtVgsk.SetFocus
    End If
End Sub

Private Sub lblWatts_Click()
  
    If txtEmnr <> "" Then
        Dim sStation1Min As String
        Dim sStation1Max As String
        sStation1Min = Val(gStation1Min)
        sStation1Max = Val(gStation1Max)
        
        If frmDefaults!txtStation1Min.Text = "" Or frmDefaults!txtStation1Max.Text = "" Then
                lblPwrLimit1.Caption = "limits not set"
                lblStation1.ToolTipText = "Minimum and Maximum power limits have not been set"
                Exit Sub
        ElseIf miWmnr < Val(gStation1Min) Or miWmnr > Val(gStation1Max) Then
            lblmnr.ForeColor = vbRed
            If miWmnr < Val(gStation1Min) Then
                lblHL1 = Format(sStation1Min, "#,###") & vbCrLf & "(min)"
                lblPwrLimit1.Caption = "below limit"
            ElseIf miWmnr > Val(gStation1Max) Then
                lblHL1 = Format(sStation1Max, "#,###") & vbCrLf & "(max)"
                lblPwrLimit1.Caption = "above limit"
            
            End If
        Else
            lblmnr.ForeColor = vbBlack
            lblHL1 = ""
            lblPwrLimit1.Caption = ""
        End If
    End If
    If lblmnr = "" Then
        lblHL1 = ""
    End If
'----------------
    If txtErxc <> "" Then
        Dim sStation2Min As String
        Dim sStation2Max As String
        sStation2Min = Val(gStation2Min)
        sStation2Max = Val(gStation2Max)
        
         If frmDefaults!txtStation2Min.Text = "" Or frmDefaults!txtStation2Max.Text = "" Then
                lblPwrLimit2.Caption = "limits not set"
                lblStation2.ToolTipText = "Minimum and Maximum power limits have not been set"
                Exit Sub
        ElseIf miWrxc < Val(gStation2Min) Or miWrxc > Val(gStation2Max) Then
            lblrxc.ForeColor = vbRed
            If miWrxc < Val(gStation2Min) Then
                lblHL2 = Format(sStation2Min, "#,###") & vbCrLf & "(min)"
                lblPwrLimit2.Caption = "below limit"
            ElseIf miWrxc > Val(gStation2Max) Then
                lblHL2 = Format(sStation2Max, "#,###") & vbCrLf & "(max)"
                lblPwrLimit2.Caption = "above limit"
            End If
        Else
            lblrxc.ForeColor = vbBlack
            lblHL2 = ""
            lblPwrLimit2.Caption = ""
        End If

    End If
    If lblrxc = "" Then
        lblHL2 = ""
    End If
    '--------------
    If txtEgrs <> "" Then
    
        Dim sStation3Min As String
        Dim sStation3Max As String
        sStation3Min = Val(gStation3Min)
        sStation3Max = Val(gStation3Max)
        
         If frmDefaults!txtStation3Min.Text = "" Or frmDefaults!txtStation3Max.Text = "" Then
                lblPwrLimit3.Caption = "limits not set"
                lblStation3.ToolTipText = "Minimum and Maximum power limits have not been set"
                Exit Sub
        ElseIf miWgrs < Val(gStation3Min) Or miWgrs > Val(gStation3Max) Then
            lblgrs.ForeColor = vbRed
            If miWgrs < Val(gStation3Min) Then
                lblHL3 = Format(sStation3Min, "#,###") & vbCrLf & "(min)"
                lblPwrLimit3.Caption = "below limit"
            ElseIf miWgrs > Val(gStation3Max) Then
                lblHL3 = Format(sStation3Max, "#,###") & vbCrLf & "(max)"
                lblPwrLimit3.Caption = "above limit"
            End If
        Else
            lblgrs.ForeColor = vbBlack
            lblHL3 = ""
            lblPwrLimit3.Caption = ""
        End If
    End If
    If lblgrs = "" Then
        lblHL3 = ""
    End If
    '--------------
     If txtEgsk <> "" Then
    
        Dim sStation4Min As String
        Dim sStation4Max As String
        sStation4Min = Val(gStation4Min)
        sStation4Max = Val(gStation4Max)
        
         If frmDefaults!txtStation4Min.Text = "" Or frmDefaults!txtStation4Max.Text = "" Then
                lblPwrLimit4.Caption = "limits not set"
                lblStation4.ToolTipText = "Minimum and Maximum power limits have not been set"
                Exit Sub
        ElseIf miWgsk < Val(gStation4Min) Or miWgsk > Val(gStation4Max) Then
            lblgsk.ForeColor = vbRed
            If miWgsk < Val(gStation4Min) Then
                lblHL4 = Format(sStation4Min, "#,###") & vbCrLf & "(min)"
                lblPwrLimit4.Caption = "below limit"
            ElseIf miWgsk > Val(gStation4Max) Then
                lblHL4 = Format(sStation4Max, "#,###") & vbCrLf & "(max)"
                lblPwrLimit4.Caption = "above limit"
            End If
        Else
            lblgsk.ForeColor = vbBlack
            lblHL4 = ""
            lblPwrLimit4.Caption = ""
        End If
    End If
    If lblgsk = "" Then
        lblHL4 = ""
    End If
End Sub

Private Sub Label11_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
   
   If mnuDigitHighlights.Checked = True Then
     mnuDigitHighlights_Click
   End If
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub lblgrs_Change()
    If txtEgrs <> "" Then

        Dim sStation3Min As String
        Dim sStation3Max As String
        sStation3Min = Val(gStation3Min)
        sStation3Max = Val(gStation3Max)

         If frmDefaults!txtStation3Min.Text = "" Or frmDefaults!txtStation3Max.Text = "" Then
                lblPwrLimit3.Caption = "limits not set"
                lblStation3.ToolTipText = "Minimum and Maximum power limits have not been set"
                Exit Sub
        ElseIf miWgrs < Val(gStation3Min) Or miWgrs > Val(gStation3Max) Then
            lblgrs.ForeColor = vbRed
            If miWgrs < Val(gStation3Min) Then
                lblHL3 = Format(sStation3Min, "#,###") & vbCrLf & "(min)"
                lblPwrLimit3.Caption = "below limit"
            ElseIf miWgrs > Val(gStation3Max) Then
                lblHL3 = Format(sStation3Max, "#,###") & vbCrLf & "(max)"
                lblPwrLimit3.Caption = "above limit"
            End If
        Else
            lblgrs.ForeColor = vbBlack
            lblHL3 = ""
            lblPwrLimit3.Caption = ""
        End If
    End If
    If lblgrs = "" Then
        lblHL3 = ""
    End If
End Sub

  Private Sub lblgsk_Change()
    If txtEgsk <> "" Then
    
        Dim sStation4Min As String
        Dim sStation4Max As String
        sStation4Min = Val(gStation4Min)
        sStation4Max = Val(gStation4Max)
        
         If frmDefaults!txtStation4Min.Text = "" Or frmDefaults!txtStation4Max.Text = "" Then
                lblPwrLimit4.Caption = "limits not set"
                lblStation4.ToolTipText = "Minimum and Maximum power limits have not been set"
                Exit Sub
        ElseIf miWgsk < Val(gStation4Min) Or miWgsk > Val(gStation4Max) Then
            lblgsk.ForeColor = vbRed
            If miWgsk < Val(gStation4Min) Then
                lblHL4 = Format(sStation4Min, "#,###") & vbCrLf & "(min)"
                lblPwrLimit4.Caption = "below limit"
            ElseIf miWgsk > Val(gStation4Max) Then
                lblHL4 = Format(sStation4Max, "#,###") & vbCrLf & "(max)"
                lblPwrLimit4.Caption = "above limit"
            End If
        Else
            lblgsk.ForeColor = vbBlack
            lblHL4 = ""
            lblPwrLimit4.Caption = ""
        End If
    End If
    If lblgsk = "" Then
        lblHL4 = ""
    End If
End Sub

Private Sub lblHL1_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
   
   If mnuDigitHighlights.Checked = True Then
     mnuDigitHighlights_Click
   End If
   
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub lblHL2_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
   
   If mnuDigitHighlights.Checked = True Then
     mnuDigitHighlights_Click
   End If
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub lblHL3_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
   
   If mnuDigitHighlights.Checked = True Then
     mnuDigitHighlights_Click
   End If
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub lblHL4_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
   
   If mnuDigitHighlights.Checked = True Then
     mnuDigitHighlights_Click
   End If
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
End Sub

Private Sub lblmnr_Change()
    If txtEmnr <> "" Then
        Dim sStation1Min As String
        Dim sStation1Max As String
        sStation1Min = Val(gStation1Min)
        sStation1Max = Val(gStation1Max)
        
        If frmDefaults!txtStation1Min.Text = "" Or frmDefaults!txtStation1Max.Text = "" Then
                lblPwrLimit1.Caption = "limits not set"
                lblStation1.ToolTipText = "Minimum and Maximum power limits have not been set"
                Exit Sub
        ElseIf miWmnr < Val(gStation1Min) Or miWmnr > Val(gStation1Max) Then
            lblmnr.ForeColor = vbRed
            If miWmnr < Val(gStation1Min) Then
                lblHL1 = Format(sStation1Min, "#,###") & vbCrLf & "(min)"
                lblPwrLimit1.Caption = "below limit"
            ElseIf miWmnr > Val(gStation1Max) Then
                lblHL1 = Format(sStation1Max, "#,###") & vbCrLf & "(max)"
                lblPwrLimit1.Caption = "above limit"
            
            End If
        Else
            lblmnr.ForeColor = vbBlack
            lblHL1 = ""
            lblPwrLimit1.Caption = ""
        End If
    End If
    If lblmnr = "" Then
        lblHL1 = ""
    End If
End Sub

Private Sub lblrxc_Change()
    If txtErxc <> "" Then
        Dim sStation2Min As String
        Dim sStation2Max As String
        sStation2Min = Val(gStation2Min)
        sStation2Max = Val(gStation2Max)
        
         If frmDefaults!txtStation2Min.Text = "" Or frmDefaults!txtStation2Max.Text = "" Then
                lblPwrLimit2.Caption = "limits not set"
                lblStation2.ToolTipText = "Minimum and Maximum power limits have not been set"
                Exit Sub
        ElseIf miWrxc < Val(gStation2Min) Or miWrxc > Val(gStation2Max) Then
            lblrxc.ForeColor = vbRed
            If miWrxc < Val(gStation2Min) Then
                lblHL2 = Format(sStation2Min, "#,###") & vbCrLf & "(min)"
                lblPwrLimit2.Caption = "below limit"
            ElseIf miWrxc > Val(gStation2Max) Then
                lblHL2 = Format(sStation2Max, "#,###") & vbCrLf & "(max)"
                lblPwrLimit2.Caption = "above limit"
            End If
        Else
            lblrxc.ForeColor = vbBlack
            lblHL2 = ""
            lblPwrLimit2.Caption = ""
        End If
    End If
    If lblrxc = "" Then
        lblHL2 = ""
    End If
End Sub

Private Sub lblStation1_DblClick()
    If lblStation1.ForeColor <> &H808080 Then
        lblmnr = ""
        lblmnr.BackColor = &H8000000F
        txtVmnr.BackColor = &H8000000F
        txtAmnr.BackColor = &H8000000F
        txtEmnr.BackColor = &H8000000F
        txtEmnr.ForeColor = &H808080     'gray
        lblHL1 = ""
        lblStation1.ForeColor = &H808080   '&HFF&
        txtVmnr.Enabled = False
        txtAmnr.Enabled = False
        lblStation1.ToolTipText = "Double-Click call letters to restore readings function"
    Else
        lblmnr = Format(miWmnr, "#,###")
        lblmnr.BackColor = &H80000018
        
        txtVmnr.BackColor = &H80000005 'white
        txtAmnr.BackColor = &H80000005
        txtEmnr.BackColor = &H80000018 ' yellow
        txtEmnr.ForeColor = &H80000008  'black
        
        lblStation1.ForeColor = &H80000012
        txtVmnr.Enabled = True
        txtAmnr.Enabled = True
        lblStation1.ToolTipText = lblStation1 & " Minimum power: " & frmDefaults!txtStation1Min & " watts,  Maximum power: " & frmDefaults!txtStation1Max & " watts"
    End If
End Sub

Private Sub lblStation2_DblClick()
    If lblStation2.ForeColor <> &H808080 Then
        lblrxc = ""
        lblrxc.BackColor = &H8000000F
        txtVrxc.BackColor = &H8000000F
        txtArxc.BackColor = &H8000000F
        txtErxc.BackColor = &H8000000F
        txtErxc.ForeColor = &H808080     'gray
        lblHL2 = ""
        txtArxc.BackColor = &H8000000F
        lblStation2.ForeColor = &H808080
        txtVrxc.Enabled = False
        txtArxc.Enabled = False
        lblStation2.ToolTipText = "Double-Click call letters to restore readings function"
    Else
        lblrxc = Format(miWrxc, "#,###")
        lblrxc.BackColor = &H80000018
        
        txtVrxc.BackColor = &H80000005 'white
        txtArxc.BackColor = &H80000005
        txtErxc.BackColor = &H80000018  'yellow '&HE0E0E0    'gray
        txtErxc.ForeColor = &H80000008  'black
       
        lblStation2.ForeColor = &H80000012
        txtVrxc.Enabled = True
        txtArxc.Enabled = True
        lblStation2.ToolTipText = lblStation2 & " Minimum power: " & frmDefaults!txtStation2Min & " watts,  Maximum power: " & frmDefaults!txtStation2Max & " watts"
    End If
End Sub

Private Sub lblStation3_DblClick()
    If lblStation3.ForeColor <> &H808080 Then
        lblgrs = ""
        lblgrs.BackColor = &H8000000F
        txtVgrs.BackColor = &H8000000F
        txtAgrs.BackColor = &H8000000F
        txtEgrs.BackColor = &H8000000F
        txtEgrs.ForeColor = &H808080     'gray
        lblHL3 = ""
        lblStation3.ForeColor = &H808080
        txtVgrs.Enabled = False
        txtAgrs.Enabled = False
        lblStation3.ToolTipText = "Double-Click call letters to restore readings function"
    Else
        lblgrs = Format(miWgrs, "#,###")
        lblgrs.BackColor = &H80000018
        
        txtVgrs.BackColor = &H80000005 'white
        txtAgrs.BackColor = &H80000005
        txtEgrs.BackColor = &H80000018  'yellow
        txtEgrs.ForeColor = &H80000008  'black
        
        lblStation3.ForeColor = &H80000012
        txtVgrs.Enabled = True
        txtAgrs.Enabled = True
        lblStation3.ToolTipText = lblStation3 & " Minimum power: " & frmDefaults!txtStation3Min & " watts,  Maximum power: " & frmDefaults!txtStation3Max & " watts"
    End If
End Sub

Private Sub lblStation4_DblClick()
    If lblStation4.ForeColor <> &H808080 Then
        lblgsk = ""
        lblgsk.BackColor = &H8000000F 'gray
        
        txtVgsk.BackColor = &H8000000F
        txtAgsk.BackColor = &H8000000F
        txtEgsk.BackColor = &H8000000F
        txtEgsk.ForeColor = &H808080     'gray
        lblHL4 = ""
        lblStation4.ForeColor = &H808080
        txtVgsk.Enabled = False
        txtAgsk.Enabled = False
        lblStation4.ToolTipText = "Double-Click call letters to restore readings function"
    Else
        lblgsk = Format(miWgsk, "#,###")
        lblgsk.BackColor = &H80000018
        
        txtVgsk.BackColor = &H80000005 'white
        txtAgsk.BackColor = &H80000005
        txtEgsk.BackColor = &H80000018  'yellow
        txtEgsk.ForeColor = &H80000008  'black
        
        lblStation4.ForeColor = &H80000012
        txtVgsk.Enabled = True
        txtAgsk.Enabled = True
        lblStation4.ToolTipText = lblStation4 & " Minimum power: " & frmDefaults!txtStation4Min & " watts,  Maximum power: " & frmDefaults!txtStation4Max & " watts"
    End If
End Sub

Private Sub lblTime_Click()
    If txtTime.Enabled = True Then
        txtTime.Visible = True 'to restore txtTime box for editing time
        txtTime.SetFocus
    End If
End Sub

Private Sub lstXmitter_DblClick()
      
    Dim iTwr As String
    If optTwrOn = True Then
      iTwr = "Tower Lights ON"
    End If
    If mnuDigitHighlights.Checked = True Then
      mnuDigitHighlights_Click
    End If
   
    If lstXmitter.ListCount > 19 Then 'scrolls screen
        lstXmitter.TopIndex = lstXmitter.ListCount - 18
    End If
   
 '---------
    Dim mnrError As String
    Dim rxcError As String
    Dim grsError As String
    Dim gskError As String
    
    If lblmnr.ForeColor = vbRed Then
        mnrError = " *"
    Else: mnrError = ""
    End If
    
    If lblrxc.ForeColor = vbRed Then
        rxcError = " *"
    Else: rxcError = ""
    End If
    
    If lblgrs.ForeColor = vbRed Then
        grsError = " *"
    Else: grsError = ""
    End If
    
    If lblgsk.ForeColor = vbRed Then
        gskError = " *"
    Else: gskError = ""
    End If
  '---------
    lstXmitter.AddItem lblTime & "       " & iTwr
    If lblmnr <> "" Then
        lstXmitter.AddItem "  " & lblStation1 & "   Volts:  " & txtVmnr & "    Amps:  " & txtAmnr & "    Watts:  " & lblmnr & mnrError
        
        If mnrError = " *" Then
            If miWmnr < Val(gStation1Min) Then
                lstXmitter.AddItem "      ---" & lblStation1 & " power below normal minimum of " & gStation1Min & " watts"
            
            ElseIf miWmnr > Val(gStation1Max) Then
                lstXmitter.AddItem "      ---" & lblStation1 & " power above normal maximum of " & gStation1Max & " watts"
            End If
        End If
    End If
    
    If lblrxc <> "" Then
        lstXmitter.AddItem "  " & lblStation2 & "   Volts:  " & txtVrxc & "    Amps:  " & txtArxc & "    Watts:  " & lblrxc & rxcError
        If rxcError = " *" Then
            If miWrxc < Val(gStation2Min) Then
                lstXmitter.AddItem "      ---" & lblStation2 & " power below normal minimum of " & gStation2Min & " watts"
            
            ElseIf miWrxc > Val(gStation2Max) Then
                lstXmitter.AddItem "      ---" & lblStation2 & " power above normal maximum of " & gStation2Max & " watts"
            End If
        End If
    End If
    
    If lblgrs <> "" Then
        lstXmitter.AddItem "  " & lblStation3 & "   Volts:  " & txtVgrs & "    Amps:  " & txtAgrs & "    Watts:  " & lblgrs & grsError
        If grsError = " *" Then
            If miWgrs < Val(gStation3Min) Then
                lstXmitter.AddItem "      ---" & lblStation3 & " power below normal minimum of " & gStation3Min & " watts"
            
            ElseIf miWgrs > Val(gStation3Max) Then
                lstXmitter.AddItem "      ---" & lblStation3 & " power above normal maximum of " & gStation3Max & " watts"
            End If
        End If
    End If
    
    If lblgsk <> "" Then
        lstXmitter.AddItem "  " & lblStation4 & "   Volts:  " & txtVgsk & "    Amps:  " & txtAgsk & "    Watts:  " & lblgsk & gskError
        If gskError = " *" Then
            If miWgsk < Val(gStation4Min) Then
                lstXmitter.AddItem "      ---" & lblStation4 & " power below normal minimum of " & gStation4Min & " watts"
            
            ElseIf miWgsk > Val(gStation4Max) Then
                lstXmitter.AddItem "      ---" & lblStation4 & " power above normal maximum of " & gStation4Max & " watts"
            End If
        End If
    End If
  
    lstXmitter.AddItem ""
  
 '----display count of screen entries
    
    Dim iListCount As Integer

    Dim iMnr As Integer
    Dim iRxc As Integer
    Dim iGrs As Integer
    Dim iGsk As Integer
    Dim iStaCount As Integer

    If lblmnr <> "" Then
        iMnr = 1
    Else
        iMnr = 0
    End If
    
    If lblrxc <> "" Then
        iRxc = 1
    Else
        iRxc = 0
    End If
    
    If lblgrs <> "" Then
        iGrs = 1
    Else
        iGrs = 0
    End If

    If lblgsk <> "" Then
        iGsk = 1
    Else
        iGsk = 0
    End If

    iStaCount = iMnr + iRxc + iGrs + iGsk

    Select Case iStaCount
    Case 0
        iListCount = 0
    Case 1
        iListCount = (lstXmitter.ListCount - 1) / 3
    Case 2
        iListCount = (lstXmitter.ListCount - 1) / 4
    Case 3
        iListCount = (lstXmitter.ListCount - 1) / 5
    Case 4
        iListCount = (lstXmitter.ListCount - 1) / 6
    Case Else
        iListCount = 0
    End Select
   
   lblEntries = iListCount 'displays number of entries
   
   If lstXmitter.ListCount > 2 Then
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
    End If
 '-------

    If lstXmitter.ListCount > 4 Then
        frmPlanner!chkPrintXmitter.Visible = True
    End If
    
          'scrolls list boxes to last entry
    If lstXmitter.ListCount > 20 Then
       lstXmitter.TopIndex = lstXmitter.ListCount - 19
    End If

On Error GoTo HandleErrors
    Open "PwrL.dat" For Output As #14
    Dim i As Integer
    For i = 0 To lstXmitter.ListCount - 1
        Print #14, lstXmitter.List(i)
    Next
    Close #14
HandleErrors:
    
    If lblEntries = "1" Then
        Label2.Caption = "Entry"
    Else
        Label2.Caption = "Entries"
    End If

    txtTime = "Time?"
    lblTwr.Visible = False
    imgTower.Visible = False
    txtTime.Visible = True
    Label1.Visible = False
    Label11.Visible = True
    If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
    End If
    
    If lstXmitter.ListCount > 2 Then
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
    End If
    
 End Sub

Private Sub lstXmitter_LostFocus()
    lstXmitter.ListIndex = -1
End Sub

Private Sub mnuAdvancedDefault_Click()

   If txtVmnr = "" Or lblStation1 = "" Or txtAmnr = "" Or txtEmnr = "" Then
        lblDataMissing.Visible = True
        MsgBox "Saving current readings as default values requires a minimum of Flagship Station call letters and Volts, Amps, and Efficiency entries", _
        vbOKOnly, "Required Data Missing - Values Not Saved as Defaults"
        Exit Sub
    End If

    Dim prompt, AccessCode
    prompt = "Access Code is required to set current transmitter readings and effeciency values as defaults." & vbCrLf & vbCrLf & _
    "Station call letters, transmitter minimum, maximum and ideal power setting defaults are set from the 'Defaults' page." _
    & vbCrLf & vbCrLf & "Enter Access Code."
    AccessCode = InputBox$(prompt, "Access Code Required")

    If AccessCode = giAccess Then
    
        Dim iResponse As Integer
        iResponse = MsgBox("Access code is correct." & vbCrLf & vbCrLf & "Confirm you want to save the current transmitter readings as default values?", vbOKCancel, "Default Values")
        If iResponse = vbCancel Then
            Exit Sub
        ElseIf iResponse = vbOK Then
        End If
        
        Dim rDate As Date
        rDate = Now
    
        lblDataMissing.Visible = False
    
'        Open "DefaultStation.dat" For Output As #20
'            Write #20, frmDefaults!txtStation1, frmDefaults!txtStation2, frmDefaults!txtStation3, frmDefaults!txtStation4
'        Close #20
        
'        If frmDefaults!txtStation1Min <> "" Then
'        Open "DefaultMinMax.dat" For Output As #24
'            Write #24, frmDefaults!txtStation1Min, frmDefaults!txtStation1Max, frmDefaults!txtStation2Min, frmDefaults!txtStation2Max, _
'            frmDefaults!txtStation3Min, frmDefaults!txtStation3Max, frmDefaults!txtStation4Min, frmDefaults!txtStation4Max
'        Close #24
'        End If
        
        Open "DefaultReadings.dat" For Output As #16
            Write #16, txtVmnr, txtVrxc, txtVgrs, txtVgsk, txtAmnr, txtArxc, _
            txtAgrs, txtAgsk, txtEmnr, txtErxc, txtEgrs, txtEgsk, _
            txtVolt1, txtVolt2, txtVolt3, txtVolt4, txtAmp1, txtAmp2, txtAmp3, txtAmp4, _
            txtEmnr, txtErxc, txtEgrs, txtEgsk
        Close #16
                    
        Open "DefaultDate.dat" For Output As #17
            Write #17, rDate
        Close #17
'
'        If frmDefaults!txtMin <> "" And frmDefaults!txtMax <> "" And (frmDefaults!txtIdeal1 <> "" Or frmDefaults!txtIdeal2 <> "" _
'        Or frmDefaults!txtIdeal3 <> "" Or frmDefaults!txtIdeal4 <> "") Then
'
'            Open "DefaultXmitterIdeal.dat" For Output As #26
'                    Write #26, frmDefaults!txtMin, frmDefaults!txtMax, frmDefaults!txtIdeal1, frmDefaults!txtIdeal2, frmDefaults!txtIdeal3, frmDefaults!txtIdeal4, _
'                    frmDefaults!ckIdeal1.Value, frmDefaults!ckIdeal2.Value, frmDefaults!ckIdeal3.Value, frmDefaults!ckIdeal4.Value,
'            Close #26
'        End If
              
    ElseIf AccessCode <> "" Then
        MsgBox "Incorrect Access Code", vbOKOnly, "Incorrect Code"
    Else
    End If

End Sub

Private Sub mnuAdvancePwWarn_Click()

    If mnuAdvancePwWarn.Checked = True Then
        mnuAdvancePwWarn.Checked = False
        lblLimitWarn.Visible = False
        miStation1 = 0
        miStation2 = 0
        miStation3 = 0
        miStation4 = 0
        miStation = 0 'rearms 'bypass station' message
     
    Else
        mnuAdvancePwWarn.Checked = True
        lblLimitWarn.Visible = True
        miStation1 = 4
        miStation2 = 4
        miStation3 = 4
        miStation4 = 4
    End If
    
End Sub

Private Sub mnuDigitHighlights_Click()

   If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
      mnuSettingsEffic_Click
   End If
   
    If lblLimitWarn.Visible = True Then
        mnuAdvancePwWarn_Click
    End If
          
   If mnuDigitHighlights.Checked = True Then
       mnuDigitHighlights.Checked = False
       
       Frame1.Caption = "Transmitter Readings  (Overwrite Existing Reading && Tab)"
       txtVolt1.Visible = False
       txtVolt2.Visible = False
       txtVolt3.Visible = False
       txtVolt4.Visible = False
       txtAmp1.Visible = False
       txtAmp2.Visible = False
       txtAmp3.Visible = False
       txtAmp4.Visible = False
       Label7.Visible = False
   Else
        mnuDigitHighlights.Checked = True
        Beep
        
        Frame1.Caption = ""
        txtVolt1.Visible = True
        txtVolt2.Visible = True
        txtVolt3.Visible = True
        txtVolt4.Visible = True
        txtAmp1.Visible = True
        txtAmp2.Visible = True
        txtAmp3.Visible = True
        txtAmp4.Visible = True
        Label7.Visible = True
       
        If txtVolt1 = "" Then
            txtVolt1 = "0"
        End If
        If txtVolt2 = "" Then
            txtVolt2 = "0"
        End If
        If txtVolt3 = "" Then
            txtVolt3 = "0"
        End If
        If txtVolt4 = "" Then
            txtVolt4 = "0"
        End If
        If txtAmp1 = "" Then
            txtAmp1 = "0"
        End If
        If txtAmp2 = "" Then
            txtAmp2 = "0"
        End If
        If txtAmp3 = "" Then
            txtAmp3 = 0
        End If
        If txtAmp4 = "" Then
            txtAmp4 = "0"
        End If
        txtVolt1.SetFocus
    End If
End Sub

Private Sub mnuDigitHighlights2_Click()
    mnuDigitHighlights_Click
End Sub

Private Sub mnuHints_Click()
    frmTransmitterHints.Show vbModal
End Sub

Private Sub mnuPagePlanner_Click()
   If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    giClockShow = 3
    frmPlanner.Show
    frmTransmitter.Hide
End Sub

Private Sub mnuSettingsCall_Click()
    giClockShow = 3
    If mnuSettingsEffic.Checked = True Then
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
      mnuDigitHighlights_Click
    End If
    
    frmDefaults!txtStation1.Enabled = True
    frmDefaults!txtStation2.Enabled = True
    frmDefaults!txtStation3.Enabled = True
    frmDefaults!txtStation4.Enabled = True
    
    frmDefaults!Label5.Enabled = True
    frmDefaults!Label6.Enabled = True
    frmDefaults!Label7.Enabled = True
    frmDefaults!Label8.Enabled = True
    frmDefaults!Label19.Enabled = True
   '---------------
   'disables all except station call letters
    frmDefaults!Label9 = "Set Station Call Letters"
    frmDefaults!txtStation1Min.Enabled = False
    frmDefaults!txtStation2Min.Enabled = False
    frmDefaults!txtStation3Min.Enabled = False
    frmDefaults!txtStation4Min.Enabled = False
    
    frmDefaults!txtStation1Max.Enabled = False
    frmDefaults!txtStation2Max.Enabled = False
    frmDefaults!txtStation3Max.Enabled = False
    frmDefaults!txtStation4Max.Enabled = False
    
    frmDefaults!txtPlanTime.Enabled = False
    frmDefaults!txtIntroOut.Enabled = False
    frmDefaults!txtClose.Enabled = False
    frmDefaults!txtSpot.Enabled = False
    frmDefaults!Frame2.Enabled = False
    
    frmDefaults!Shape1.Visible = True
    frmDefaults!Label14.Enabled = False
    frmDefaults!Label15.Enabled = False
    frmDefaults!Label16.Enabled = False
    frmDefaults!Label17.Enabled = False
    frmDefaults!Label10.Enabled = False
    frmDefaults!Label11.Enabled = False
    frmDefaults!Label12.Enabled = False
    frmDefaults!Label13.Enabled = False
    frmDefaults!mnuHelp.Enabled = False
    frmDefaults!mnuPage.Enabled = False
    frmDefaults!mnuFile.Enabled = False
    frmDefaults!mnuRestoreDefaults.Enabled = False
    
    frmDefaults!Frame2.Enabled = False
    frmDefaults!Label1.Enabled = False
    frmDefaults!Label18.Enabled = False
    frmDefaults!Label2.Enabled = False
    frmDefaults!Label3.Enabled = False
    frmDefaults!Label4.Enabled = False
    
    frmDefaults!Frame3.Enabled = False 'setting power based on ideal
    frmDefaults!Label20.Enabled = False
    frmDefaults!Label21.Enabled = False
    frmDefaults!Label22.Enabled = False
    frmDefaults!Label23.Enabled = False
    frmDefaults!Label26.Enabled = False
    
    frmDefaults!ckIdeal1.Enabled = False
    frmDefaults!ckIdeal2.Enabled = False
    frmDefaults!ckIdeal3.Enabled = False
    frmDefaults!ckIdeal4.Enabled = False
    frmDefaults!txtIdeal1.Enabled = False
    frmDefaults!txtIdeal2.Enabled = False
    frmDefaults!txtIdeal3.Enabled = False
    frmDefaults!txtIdeal4.Enabled = False
    frmDefaults!txtMax.Enabled = False
    frmDefaults!txtMin.Enabled = False
    frmDefaults!lblIdeal1.Enabled = False
    frmDefaults!lblIdeal2.Enabled = False
    frmDefaults!lblIdeal3.Enabled = False
    frmDefaults!lblIdeal4.Enabled = False
   
    frmDefaults!Frame1.Caption = "Set Station Call Letters"
    frmDefaults.Show
    frmDefaults!cmdDefaults.Enabled = False
    frmTransmitter.Hide

End Sub

Private Sub mnuSettingsEffic_Click()
    txtEmnr.MousePointer = 1 'arrow
    txtErxc.MousePointer = 1
    txtEgrs.MousePointer = 1
    txtEgsk.MousePointer = 1

   If mnuDigitHighlights.Checked = True Then
      mnuDigitHighlights_Click
   End If
   
    If mnuSettingsEffic.Checked = True Then
         mnuSettingsEffic.Checked = False 'procedure not in use
         txtEmnr.Locked = True
         txtErxc.Locked = True
         txtEgrs.Locked = True
         txtEgsk.Locked = True
         
         txtEmnr.TabStop = False
         txtErxc.TabStop = False
         txtEgrs.TabStop = False
         txtEgsk.TabStop = False
         
         txtEmnr.BackColor = &H80000018  'yellow
         txtErxc.BackColor = &H80000018  'yellow
         txtEgrs.BackColor = &H80000018  'yellow
         txtEgsk.BackColor = &H80000018  'yellow
         
        Frame1.ForeColor = &H80&
        Frame1.FontSize = 8
        Frame1.Caption = "Transmitter Readings  (Overwrite Existing Reading && Tab)"
         
        If lblStation1.Caption = "" Then
            txtEmnr.Text = ""
        End If
        
        If lblStation2.Caption = "" Then
            txtErxc.Text = ""
        End If
        
        If lblStation3.Caption = "" Then
            txtEgrs.Text = ""
        End If
        
        If lblStation4.Caption = "" Then
            txtEgsk.Text = ""
        End If
         
        lblEfficiency.Visible = True
        Label6.Visible = False
        mnuHints.Enabled = True
        If txtVmnr.Enabled = True Then
            txtVmnr.SetFocus
        End If
    Else
        Frame1.ForeColor = &HFF0000
        Frame1.FontSize = 10
        Frame1.Caption = "Set Transmitter Efficiency Values" 'procedure in use
        txtEmnr.MousePointer = 3 'I beam
        txtErxc.MousePointer = 3
        txtEgrs.MousePointer = 3
        txtEgsk.MousePointer = 3
        mnuSettingsEffic.Checked = True
        Beep
        txtEmnr.Locked = False
        txtErxc.Locked = False
        txtEgrs.Locked = False
        txtEgsk.Locked = False
        
        txtEmnr.TabStop = True
        txtErxc.TabStop = True
        txtEgrs.TabStop = True
        txtEgsk.TabStop = True
        
        txtEmnr.BackColor = &HE0EFDE 'green
        txtErxc.BackColor = &HE0EFDE
        txtEgrs.BackColor = &HE0EFDE
        txtEgsk.BackColor = &HE0EFDE
     
        txtEmnr.ToolTipText = ""
        txtErxc.ToolTipText = ""
        txtEgrs.ToolTipText = ""
        txtEgsk.ToolTipText = ""
        
        lblEfficiency.Visible = False
        Label6.Visible = True
        mnuHints.Enabled = False
     txtEmnr.SetFocus
    End If
   
End Sub

Private Sub mnuPageAddTime_Click()

    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
   
    frmAddTime.Show
End Sub

Private Sub mnuPreviousPage_Click()

    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
        mnuDigitHighlights_Click
    End If
 
    If giClockShow <> 0 Then
        Select Case giClockShow
            Case 4
                frmPlanner.Show
            Case 5
                frmTimeRemain.Show
            Case 6
                frmDefaults.Show
            Case Else
                frmPlanner.Show
        End Select
        frmTransmitter.Hide
        giClockShow = 3
    Else
        frmPlanner.Show
        frmTransmitter.Hide
        giClockShow = 3
    End If
    
End Sub

Private Sub mnuPageStopWatch_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If

    frmStopWatch.Show
End Sub

Private Sub mnuPageTimeRemain_Click()
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    giClockShow = 3
    frmTimeRemain.Show
    frmTransmitter.Hide
End Sub

Private Sub mnuSettingsPower_Click()
    giClockShow = 3
    frmDefaults.Caption = "Set Transmitter Power Limits"
    frmDefaults!cmdDefaults.Enabled = False
    If mnuSettingsEffic.Checked = True Then
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
      mnuDigitHighlights_Click 'close set digital highlighted position
    End If
    
   ' giDefaultsFocus = 1
    frmDefaults!Label9 = "Set Transmitter Min && Max power Limits. Limits can be set either as a percent of ideal power or directly."
    
    
'    frmDefaults!Label9 = "Set Station Call Letters"
    frmDefaults!txtStation1Min.Enabled = True
    frmDefaults!txtStation2Min.Enabled = True
    frmDefaults!txtStation3Min.Enabled = True
    frmDefaults!txtStation4Min.Enabled = True
    
    frmDefaults!txtStation1Max.Enabled = True
    frmDefaults!txtStation2Max.Enabled = True
    frmDefaults!txtStation3Max.Enabled = True
    frmDefaults!txtStation4Max.Enabled = True
    
    frmDefaults!Label14.Enabled = True
    frmDefaults!Label15.Enabled = True
    frmDefaults!Label16.Enabled = True
    '---------
    frmDefaults!Frame3.Enabled = True 'setting power based on ideal
    frmDefaults!Label20.Enabled = True
    frmDefaults!Label21.Enabled = True
    frmDefaults!Label22.Enabled = True
    frmDefaults!Label23.Enabled = True
    frmDefaults!Label26.Enabled = True
    
    frmDefaults!ckIdeal1.Enabled = True
    frmDefaults!ckIdeal2.Enabled = True
    frmDefaults!ckIdeal3.Enabled = True
    frmDefaults!ckIdeal4.Enabled = True
    frmDefaults!txtIdeal1.Enabled = True
    frmDefaults!txtIdeal2.Enabled = True
    frmDefaults!txtIdeal3.Enabled = True
    frmDefaults!txtIdeal4.Enabled = True
    frmDefaults!txtMax.Enabled = True
    frmDefaults!txtMin.Enabled = True
    frmDefaults!lblIdeal1.Enabled = True
    frmDefaults!lblIdeal2.Enabled = True
    frmDefaults!lblIdeal3.Enabled = True
    frmDefaults!lblIdeal4.Enabled = True
    '----------
        
    frmDefaults!txtStation1.Enabled = False
    frmDefaults!txtStation2.Enabled = False
    frmDefaults!txtStation3.Enabled = False
    frmDefaults!txtStation4.Enabled = False
    
    frmDefaults!txtPlanTime.Enabled = False
    frmDefaults!txtIntroOut.Enabled = False
    frmDefaults!txtClose.Enabled = False
    frmDefaults!txtSpot.Enabled = False
    frmDefaults!Frame2.Enabled = False
    
    frmDefaults!Label1.Enabled = False
    frmDefaults!Label2.Enabled = False
    frmDefaults!Label3.Enabled = False
    frmDefaults!Label4.Enabled = False
    frmDefaults!Label5.Enabled = False
    frmDefaults!Label6.Enabled = False
    frmDefaults!Label7.Enabled = False
    frmDefaults!Label8.Enabled = False
    frmDefaults!Label10.Enabled = False
    frmDefaults!Label11.Enabled = False
    frmDefaults!Label12.Enabled = False
    frmDefaults!Label13.Enabled = False
    frmDefaults!Label18.Enabled = False
    frmDefaults!mnuHelp.Enabled = False
    frmDefaults!mnuPage.Enabled = False
    frmDefaults!mnuFile.Enabled = False
    frmDefaults!mnuRestoreDefaults.Enabled = False
    
    frmDefaults!mnuPage.Enabled = False
    frmDefaults!mnuFile.Enabled = False
    frmDefaults!mnuRestoreDefaults.Enabled = False

    frmDefaults!Shape2(0).Visible = True
    frmDefaults!Shape2(1).Visible = True
    frmDefaults!Label19.Enabled = False
    frmDefaults!Frame1.Caption = "Directly Set Transmitter Power Limits"

    frmDefaults.Show
    frmTransmitter.Hide
End Sub

Private Sub mnuSettingsPrintPage_Click()
On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "Transmitter Page")
    If iResponse = vbNo Then
        Exit Sub
    ElseIf iResponse = vbYes Then
        PrintForm
    End If
    
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

Private Sub mnuShowDefaultPage_Click()
    giClockShow = 3
    frmDefaults.Show
    frmDefaults.Caption = "Set Station Call Letters, Transmitter Power Limits, Program Time Values, and Printout Default Signature Line"
    frmDefaults!cmdCancel.SetFocus
    frmTransmitter.Hide
End Sub

Private Sub optTwrOff_Click()
    lblTwr.Visible = False
    imgTower.Visible = False
    optTwrOn.ForeColor = vbBlack
End Sub

Private Sub optTwrOn_Click()
    lblTwr.Visible = False
    imgTower.Visible = False
    If optTwrOn.Value = True Then
      imgTower.Visible = True
    End If
End Sub

Private Sub Timer1_Timer()
   Dim Today As Variant
    Today = Now
    If miTime = 0 Then
        lblTime.Caption = Format(Today, "HH:MM") '"h:mm:ss ampm")
    Else
        miTime = Val(txtTime)
    End If
End Sub

Private Sub txtAgrs_Change()
    If Not IsNumeric(txtAgrs) And txtAgrs <> "" And txtAgrs <> "." Then
        MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
        txtAgrs = ""
        Exit Sub
    End If
    pAddgrs
End Sub

Private Sub txtAgrs_DblClick()
    txtAgrs.SelStart = 0 'begin selection at start
    txtAgrs.SelLength = Len(txtAgrs) 'selects # of characters
End Sub

Private Sub txtAgrs_GotFocus()
    txtAgrs.SelStart = Val(txtAmp3) 'begin selection at start
    txtAgrs.SelLength = Len(txtAgrs) 'selects # of characters
    If chkCopyFeature.Value = 1 Then
        txtTime.Visible = True
        txtTime = "Time"
    End If
End Sub

Private Sub txtAgrs_LostFocus()

    If lblStation3 <> "" And txtVgrs <> "" And txtEgrs = "" Then
         MsgBox "The amount of power delivered from the transmitter to the antenna is dependent on the efficiency of the system." & vbCrLf & vbCrLf & _
        "Efficiency ratings normally range from a low of about  .50 (50%) to a high of about .90 (90%)." & vbCrLf & _
        "In all cases, efficiencies will be less than .99 (99%)." & vbCrLf & vbCrLf & "Enter the transmitter efficiency rating as a decimal number." _
        & vbCrLf & vbCrLf & "Efficiency rating information normally can be found on the station's transmitter log.", _
        vbOKOnly, "Enter Transmitter Efficiency Rating"
        
        If mnuSettingsEffic.Checked = False Then
            mnuSettingsEffic_Click
        End If
            
        txtEgrs.SetFocus
        Exit Sub
    End If
    
    If gStation3Min <> "" And gStation3Max <> "" And lblgrs <> "" Then
        If (miWgrs < gStation3Min - 2 Or miWgrs > gStation3Max + 2) And miStation3 = 0 Then
            miStation3 = 1
        ElseIf miWgrs >= Val(gStation3Min) And miWgrs <= Val(gStation3Max) Then
            miStation3 = 0
        End If
    End If
    If txtAgrs = "." Then
        txtAgrs = ""
    Else
       txtAgrs = Format$(txtAgrs, "0.00#")
    End If

    If lblStation3 <> "" And txtEgrs <> "" And txtVgrs = "" And txtAgrs = "" And miStation <> 1 Then
        MsgBox "Volt and Amp readings for " & lblStation3 & " have been deleted or not entered." & vbCrLf & vbCrLf & _
       "If you do not want to take transmitter readings for " & lblStation3 & " (or any other station), rather than deleting existing data, the recommended " & _
        "procedure is to Double-Click the station's call letters which will drop the station from the transmitter readings sequence." & vbCrLf & vbCrLf & _
        "The station's call letters and data entry boxes will change color to gray indicating the station is bypassed. The Tab Key will skip " _
        & "the Volts and Amps data entry boxes for that station." _
        & vbCrLf & vbCrLf & "Double-Click the station's call letters a second time " & _
        "to restore normal function and return the station to the Tab Key data entry loop.", _
        vbOKOnly + vbInformation, "Double-Click a station's call letters to bypass readings for this station"
        miStation = 1
    End If
End Sub

Private Sub txtAgsk_Change()
    If Not IsNumeric(txtAgsk) And txtAgsk <> "" And txtAgsk <> "." Then
        MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
        txtAgsk = ""
        Exit Sub
    End If
    pAddgsk
End Sub

Private Sub txtAgsk_DblClick()
    txtAgsk.SelStart = 0 'begin selection at start
    txtAgsk.SelLength = Len(txtAgsk) 'selects # of characters
End Sub

Private Sub txtAgsk_GotFocus()
    txtAgsk.SelStart = Val(txtAmp4) 'begin selection at start
    txtAgsk.SelLength = Len(txtAgsk) 'selects # of characters
    If chkCopyFeature.Value = 1 Then
        txtTime.Visible = True
        txtTime = "Time"
    End If
End Sub

Private Sub txtAgsk_LostFocus()

    If lblStation4 <> "" And txtVgsk <> "" And txtEgsk = "" Then
        MsgBox "The amount of power delivered from the transmitter to the antenna is dependent on the efficiency of the system." & vbCrLf & vbCrLf & _
        "Efficiency ratings normally range from a low of about  .50 (50%) to a high of about .90 (90%)." & vbCrLf & _
        "In all cases, efficiencies will be less than .99 (99%)." & vbCrLf & vbCrLf & "Enter the transmitter efficiency rating as a decimal number." _
        & vbCrLf & vbCrLf & "Efficiency rating information normally can be found on the station's transmitter log.", _
        vbOKOnly, "Enter Transmitter Efficiency Rating"
        
        If mnuSettingsEffic.Checked = False Then
            mnuSettingsEffic_Click
        End If
        
        txtEgsk.SetFocus
        Exit Sub
    End If

    If gStation4Min <> "" And gStation4Max <> "" And lblgsk <> "" Then
        If (miWgsk < gStation4Min - 2 Or miWgsk > gStation4Max + 2) And miStation4 = 0 Then
            miStation4 = 1
        ElseIf miWgsk >= Val(gStation4Min) And miWgsk <= Val(gStation4Max) Then
            miStation4 = 0
        End If
    End If
    
    If txtAgsk = "." Then
        txtAgsk = ""
    Else
       txtAgsk = Format$(txtAgsk, "0.00#")
    End If
    
    If mnuSettingsEffic.Checked = True And txtEgsk <> "" And Val(txtEgsk) < 1 _
    And Val(txtEmnr) < 1 And Val(txtErxc) < 1 And Val(txtEgrs) < 1 Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
     mnuDigitHighlights_Click
    End If
    
    If lblStation4 <> "" And txtEgsk <> "" And txtVgsk = "" And txtAgsk = "" And miStation <> 1 Then
         MsgBox "Volt and Amp readings for " & lblStation4 & " have been deleted or not entered." & vbCrLf & vbCrLf & _
       "If you do not want to take transmitter readings for " & lblStation4 & " (or any other station), rather than deleting existing data, the recommended " & _
        "procedure is to Double-Click the station's call letters which will drop the station from the transmitter readings sequence." & vbCrLf & vbCrLf & _
        "The station's call letters and data entry boxes will change color to gray indicating the station is bypassed. The Tab Key will skip " _
        & "the Volts and Amps data entry boxes for that station." _
        & vbCrLf & vbCrLf & "Double-Click the station's call letters a second time " & _
        "to restore normal function and return the station to the Tab Key data entry loop.", _
        vbOKOnly + vbInformation, "Double-Click a station's call letters to bypass readings for this station"
        miStation = 1
    End If

End Sub

Private Sub txtAmnr_Change()
    If Not IsNumeric(txtAmnr) And txtAmnr <> "" And txtAmnr <> "." Then
      MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
      txtAmnr = ""
      Exit Sub
    End If
    pAddmnr
End Sub

Private Sub txtAmnr_DblClick()
    txtAmnr.SelStart = 0 'begin selection at start
    txtAmnr.SelLength = Len(txtAmnr) 'selects # of characters
End Sub

Private Sub txtAmnr_GotFocus()

    txtAmnr.SelStart = Val(txtAmp1) 'begin selection at start
    txtAmnr.SelLength = Len(txtAmnr) 'selects # of characters
    If chkCopyFeature.Value = 1 Then
        txtTime.Visible = True
        txtTime = "Time"
    End If
End Sub

Private Sub txtAmnr_LostFocus()

    If lblStation1 <> "" And txtVmnr <> "" And txtEmnr = "" Then
         MsgBox "The amount of power delivered from the transmitter to the antenna is dependent on the efficiency of the system." & vbCrLf & vbCrLf & _
        "Efficiency ratings normally range from a low of about  .50 (50%) to a high of about .90 (90%)." & vbCrLf & _
        "In all cases, efficiencies will be less than .99 (99%)." & vbCrLf & vbCrLf & "Enter the transmitter efficiency rating as a decimal number." _
        & vbCrLf & vbCrLf & "Efficiency rating information normally can be found on the station's transmitter log.", _
        vbOKOnly, "Enter Transmitter Efficiency Rating"
        
        If mnuSettingsEffic.Checked = False Then
            mnuSettingsEffic_Click
        End If
        
        txtEmnr.SetFocus
        Exit Sub
    End If

    If gStation1Min <> "" And gStation1Max <> "" And lblmnr <> "" Then
        If (miWmnr < Val(gStation1Min) - 2 Or miWmnr > Val(gStation1Max) + 2) And miStation1 = 0 Then
            miStation1 = 1
        ElseIf miWmnr >= Val(gStation1Min) And miWmnr <= Val(gStation1Max) Then
            miStation1 = 0
        End If
    End If
'------
    
    If txtAmnr = "." Then
        txtAmnr = ""
    Else
       txtAmnr = Format$(txtAmnr, "0.00#")
    End If
    
    If lblStation1 <> "" And txtEmnr <> "" And txtVmnr = "" And txtAmnr = "" And miStation <> 1 Then
        MsgBox "Volt and Amp readings for " & lblStation1 & " have been deleted or not entered." & vbCrLf & vbCrLf & _
        "If you do not want to take transmitter readings for " & lblStation1 & " (or any other station), rather than deleting existing data, the recommended " & _
        "procedure is to Double-Click the station's call letters which will drop the station from the transmitter readings sequence." & vbCrLf & vbCrLf & _
        "The station's call letters and data entry boxes will change color to gray indicating the station is bypassed. The Tab Key will skip " _
        & "the Volts and Amps data entry boxes for that station." _
        & vbCrLf & vbCrLf & "Double-Click the station's call letters a second time " & _
        "to restore normal function and return the station to the Tab Key data entry loop.", _
        vbOKOnly + vbInformation, "Double-Click a station's call letters to bypass readings for this station"
        miStation = 1
    End If
    
End Sub

Private Sub txtAmp1_Change()

    If Not IsNumeric(txtAmp1) And txtAmp1 <> "" Then
        MsgBox "Non-Numeric Entry. Enter a single number of 5 or less.", vbOKOnly, "Entry Error"
        txtAmp1 = ""
        Exit Sub
    End If
    
    If txtAmp1 > "5" Then
        MsgBox "Maximum entry is 5. Choose in the range between 0 and 5." & vbCrLf & vbCrLf & _
        "With 0 all characters will be highlighted. With 5 no character will be highlighted.", vbOKOnly, "Entry Error"
        txtAmp1 = ""
        Exit Sub
    End If

End Sub

Private Sub txtAmp1_GotFocus()
  txtAmp1.SelLength = Len(txtAmp1) 'selects # of characters
End Sub

Private Sub txtAmp2_Change()

    If Not IsNumeric(txtAmp2) And txtAmp2 <> "" Then
        MsgBox "Non-Numeric Entry. Enter a single number of 5 or less.", vbOKOnly, "Entry Error"
        txtAmp2 = ""
        Exit Sub
    End If
    
    If txtAmp2 > "5" Then
        MsgBox "Maximum entry is 5. Choose in the range between 0 and 5." & vbCrLf & vbCrLf & _
        "With 0 all characters will be highlighted. With 5 no character will be highlighted.", vbOKOnly, "Entry Error"
        txtAmp2 = ""
        Exit Sub
    End If
End Sub

Private Sub txtAmp2_GotFocus()
    txtAmp2.SelLength = Len(txtAmp2) 'selects # of characters
End Sub

Private Sub txtAmp3_Change()

    If Not IsNumeric(txtAmp3) And txtAmp3 <> "" Then
        MsgBox "Non-Numeric Entry. Enter a single number of 5 or less.", vbOKOnly, "Entry Error"
        txtAmp3 = ""
        Exit Sub
    End If
    
    If txtAmp3 > "5" Then
        MsgBox "Maximum entry is 5. Choose in the range between 0 and 5." & vbCrLf & vbCrLf & _
        "With 0 all characters will be highlighted. With 5 no character will be highlighted.", vbOKOnly, "Entry Error"
        txtAmp3 = ""
        Exit Sub
    End If
End Sub

Private Sub txtAmp3_GotFocus()
    txtAmp3.SelLength = Len(txtAmp3) 'selects # of characters
End Sub

Private Sub txtAmp4_Change()

    If Not IsNumeric(txtAmp4) And txtAmp4 <> "" Then
        MsgBox "Non-Numeric Entry. Enter a single number of 5 or less.", vbOKOnly, "Entry Error"
        txtAmp4 = ""
        Exit Sub
    End If
    
    If txtAmp4 > "5" Then
        MsgBox "Maximum entry is 5. Choose in the range between 0 and 5." & vbCrLf & vbCrLf & _
        "With 0 all characters will be highlighted. With 5 no character will be highlighted.", vbOKOnly, "Entry Error"
        txtAmp4 = ""
        Exit Sub
    End If
End Sub

Private Sub txtAmp4_GotFocus()
    txtAmp4.SelLength = Len(txtAmp4) 'selects # of characters
End Sub

Private Sub txtAmp4_LostFocus()
    Open "ReadingsL.dat" For Output As #15
        Write #15, txtVmnr, txtVrxc, txtVgrs, txtVgsk, txtAmnr, txtArxc, _
        txtAgrs, txtAgsk, txtEmnr, txtErxc, txtEgrs, txtEgsk, _
        txtVolt1, txtVolt2, txtVolt3, txtVolt4, txtAmp1, txtAmp2, txtAmp3, txtAmp4
    Close #15

    mnuDigitHighlights.Checked = False
    txtVolt1.Visible = False
    txtVolt2.Visible = False
    txtVolt3.Visible = False
    txtVolt4.Visible = False
    txtAmp1.Visible = False
    txtAmp2.Visible = False
    txtAmp3.Visible = False
    txtAmp4.Visible = False
    Label7.Visible = False
    Frame1.Caption = "Transmitter Readings  (Overwrite Existing Reading && Tab)"
End Sub

Private Sub txtArxc_Change()
    If Not IsNumeric(txtArxc) And txtArxc <> "" And txtArxc <> "." Then
        MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
        txtArxc = ""
        Exit Sub
    End If
    pAddrxc
End Sub

Private Sub txtArxc_DblClick()
    txtArxc.SelStart = 0 'begin selection at start
    txtArxc.SelLength = Len(txtArxc) 'selects # of characters
End Sub

Private Sub txtArxc_GotFocus()
    txtArxc.SelStart = Val(txtAmp2) 'begin selection at start
    txtArxc.SelLength = Len(txtArxc) 'selects # of characters
    If chkCopyFeature.Value = 1 Then
        txtTime.Visible = True
        txtTime = "Time"
    End If
End Sub

Private Sub txtArxc_LostFocus()

    If lblStation2 <> "" And txtVrxc <> "" And txtErxc = "" Then
         MsgBox "The amount of power delivered from the transmitter to the antenna is dependent on the efficiency of the system." & vbCrLf & vbCrLf & _
        "Efficiency ratings normally range from a low of about  .50 (50%) to a high of about .90 (90%)." & vbCrLf & _
        "In all cases, efficiencies will be less than .99 (99%)." & vbCrLf & vbCrLf & "Enter the transmitter efficiency rating as a decimal number." _
        & vbCrLf & vbCrLf & "Efficiency rating information normally can be found on the station's transmitter log.", _
        vbOKOnly, "Enter Transmitter Efficiency Rating"
            
        If mnuSettingsEffic.Checked = False Then
            mnuSettingsEffic_Click
        End If
        
        txtErxc.SetFocus
        Exit Sub
    End If

    If gStation2Min <> "" And gStation2Max <> "" And lblrxc <> "" Then
        If (miWrxc < gStation2Min - 2 Or miWrxc > gStation2Max + 2) And miStation2 = 0 Then
            miStation2 = 1
        ElseIf miWrxc >= Val(gStation2Min) And miWrxc <= Val(gStation2Max) Then
            miStation2 = 0
        End If
    End If

'---------
    If txtArxc = "." Then
        txtArxc = ""
    Else
       txtArxc = Format$(txtArxc, "0.00#")
    End If
    
    If lblStation2 <> "" And txtErxc <> "" And txtVrxc = "" And txtArxc = "" And miStation <> 1 Then
    
       MsgBox "Volt and Amp readings for " & lblStation2 & " have been deleted or not entered." & vbCrLf & vbCrLf & _
       "If you do not want to take transmitter readings for " & lblStation2 & " (or any other station), rather than deleting existing data, the recommended " & _
        "procedure is to Double-Click the station's call letters which will drop the station from the transmitter readings sequence." & vbCrLf & vbCrLf & _
        "The station's call letters and data entry boxes will change color to gray indicating the station is bypassed. The Tab Key will skip " _
        & "the Volts and Amps data entry boxes for that station." _
        & vbCrLf & vbCrLf & "Double-Click the station's call letters a second time " & _
        "to restore normal function and return the station to the Tab Key data entry loop.", _
        vbOKOnly + vbInformation, "Double-Click a station's call letters to bypass readings for this station"
        miStation = 1
    End If
End Sub

Private Sub txtEgrs_Click()
    If mnuSettingsEffic.Checked = True Then
        txtEgrs.SelStart = 1 'begin selection at start
        txtEgrs.SelLength = Len(txtEgrs) 'selects # of characters
    End If
End Sub

Private Sub txtEgrs_GotFocus()
    If mnuSettingsEffic.Checked = True Then
        txtEgrs.SelStart = 1 'begin selection at start
        txtEgrs.SelLength = Len(txtEgsk) 'selects # of characters
    End If
End Sub

Private Sub txtEgrs_Change()
    If Val(txtEgrs) <= 0.99 Then
        If IsNumeric(txtEgrs) Or txtEgrs = "" Or txtEgrs = "." Then
            pAddgrs
        Else
            MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
            txtEgrs = ""
        End If
    End If
    
    If Val(txtEgrs) > 0.999 Then
        MsgBox "Transmitter efficiency entry cannot exceed .999. A typical entry would be" _
        & vbCrLf & " .85 (or below), which would mean the transmitter is 85% efficient (or less).", _
        vbOKOnly, "Error:   " & txtEgrs & "   is an incorrect entry.   Check the decimal point."
        txtEgrs = ""
    End If
    
    If txtEgrs = "" Then
        txtVgrs.Enabled = False
        txtAgrs.Enabled = False
        txtVgrs.BackColor = &HE0E0E0   'gray
        txtAgrs.BackColor = &HE0E0E0
    Else
        txtVgrs.Enabled = True
        txtAgrs.Enabled = True
        txtVgrs.BackColor = &H80000005  'white
        txtAgrs.BackColor = &H80000005
    End If
End Sub

Private Sub txtEgrs_LostFocus()

    If txtEgrs.BackColor = &H80000018 Then  'yellow '&HE0E0E0 Then   'gray
        Exit Sub
    End If

    If txtEgrs = "." Then
        txtEgrs = ""
    Else
        txtEgrs = Format(txtEgrs, ".00#")
    End If
End Sub

Private Sub txtEgsk_Click()
    If mnuSettingsEffic.Checked = True Then
        txtEgsk.SelStart = 1 'begin selection at start
        txtEgsk.SelLength = Len(txtEgsk) 'selects # of characters
    End If
End Sub

Private Sub txtEgsk_GotFocus()
    If mnuSettingsEffic.Checked = True Then
        txtEgsk.SelStart = 1 'begin selection at start
        txtEgsk.SelLength = Len(txtEgsk) 'selects # of characters
    End If
End Sub

Private Sub txtEgsk_Change()
    If Val(txtEgsk) <= 0.99 Then
        If IsNumeric(txtEgsk) Or txtEgsk = "" Or txtEgsk = "." Then
            pAddgsk
        Else
            MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
            txtEgsk = ""
        End If
    End If
    
    If Val(txtEgsk) > 0.999 Then
        MsgBox "Transmitter efficiency entry cannot exceed .999. A typical entry would be" _
        & vbCrLf & " .85 (or below), which would mean the transmitter is 85% efficient (or less).", _
        vbOKOnly, "Error:   " & txtEgsk & "   is an incorrect entry.   Check the decimal point."
        txtEgsk = ""
    End If
    
    If txtEgsk = "" Then
        txtVgsk.Enabled = False
        txtAgsk.Enabled = False
        txtVgsk.BackColor = &HE0E0E0   'gray
        txtAgsk.BackColor = &HE0E0E0
    Else
        txtVgsk.Enabled = True
        txtAgsk.Enabled = True
        txtVgsk.BackColor = &H80000005  'white
        txtAgsk.BackColor = &H80000005
    End If

End Sub

Private Sub txtEgsk_LostFocus()

    If txtEgsk.BackColor = &H80000018 Then  'yellow '
        Exit Sub
    End If

    If txtEgsk = "." Then
        txtEgsk = ""
    Else
        txtEgsk = Format(txtEgsk, ".00#")
    End If
    
    If mnuSettingsEffic.Checked = True And txtEmnr <> "" Then 'closes Set Efficiencies
      mnuSettingsEffic_Click
      If txtVmnr.Enabled = True Then
        txtVmnr.SetFocus
      End If
      
    ElseIf mnuSettingsEffic.Checked = True And txtEmnr = "" Then
        txtEmnr.SetFocus
    Else
    End If

End Sub

Private Sub txtEmnr_Click()
    If mnuSettingsEffic.Checked = True Then
        txtEmnr.SelStart = 1 'begin selection at start
        txtEmnr.SelLength = Len(txtEmnr) 'selects # of characters
    End If
End Sub

Private Sub txtEmnr_GotFocus()
    If mnuSettingsEffic.Checked = True Then
        txtEmnr.SelStart = 1 'begin selection at start
        txtEmnr.SelLength = Len(txtEmnr) 'selects # of characters
    End If
End Sub

Private Sub txtEmnr_Change()
    If Val(txtEmnr) <= 0.99 Then
        If IsNumeric(txtEmnr) Or txtEmnr = "" Or txtEmnr = "." Then
            pAddmnr
        Else
            MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
            txtEmnr = ""
        End If
    End If

    If Val(txtEmnr) > 0.999 Then
        MsgBox "Transmitter efficiency entry cannot exceed .999. A typical entry would be" _
        & vbCrLf & " .85 (or below), which would mean the transmitter is 85% efficient (or less).", _
        vbOKOnly, "Error:   " & txtEmnr & "   is an incorrect entry.   Check the decimal point."
        txtEmnr = ""
    End If
    
    If txtEmnr = "" Then
        txtVmnr.Enabled = False
        txtAmnr.Enabled = False
        txtVmnr.BackColor = &HE0E0E0   'gray
        txtAmnr.BackColor = &HE0E0E0
    Else
        txtVmnr.Enabled = True
        txtAmnr.Enabled = True
        txtVmnr.BackColor = &H80000005  'white
        txtAmnr.BackColor = &H80000005
        Label8.Visible = False
    End If
End Sub

Private Sub txtEmnr_LostFocus()

    If txtEmnr.BackColor = &H80000018 Then  'yellow '&HE0E0E0 Then   'gray
        Exit Sub
    End If

    If txtEmnr = "." Then
        txtEmnr = ""
    Else
        txtEmnr = Format(txtEmnr, ".00#")
    End If
    If txtEmnr = "" Then
        Label1.Visible = False
        Label11.Visible = False
        Label8.Visible = True
    End If
End Sub

Private Sub txtErxc_Click()
    If mnuSettingsEffic.Checked = True Then
        txtErxc.SelStart = 1 'begin selection at start
        txtErxc.SelLength = Len(txtErxc) 'selects # of characters
    End If
End Sub

Private Sub txtErxc_GotFocus()
    If mnuSettingsEffic.Checked = True Then
        txtErxc.SelStart = 1 'begin selection at start
        txtErxc.SelLength = Len(txtErxc) 'selects # of characters
    End If
End Sub

Private Sub txtErxc_Change()
    If Val(txtErxc) <= 0.99 Then
        If IsNumeric(txtErxc) Or txtErxc = "" Or txtErxc = "." Then
            pAddrxc
        Else
            MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
            txtErxc = ""
        End If
    End If
    
    If Val(txtErxc) > 0.999 Then
        MsgBox "Transmitter efficiency entry cannot exceed .999. A typical entry would be" _
        & vbCrLf & " .85 (or below), which would mean the transmitter is 85% efficient (or less).", _
        vbOKOnly, "Error:   " & txtErxc & "   is an incorrect entry.   Check the decimal point."
        txtErxc = ""
    End If
    
    If txtErxc = "" Then
        txtVrxc.Enabled = False
        txtArxc.Enabled = False
        txtVrxc.BackColor = &HE0E0E0   'gray
        txtArxc.BackColor = &HE0E0E0
    Else
        txtVrxc.Enabled = True
        txtArxc.Enabled = True
        txtVrxc.BackColor = &H80000005  'white
        txtArxc.BackColor = &H80000005
    End If
End Sub

Private Sub txtErxc_LostFocus()
    If txtErxc.BackColor = &H80000018 Then 'yellow '&HE0E0E0 Then   'gray
        Exit Sub
    End If

    If txtErxc = "." Then
        txtErxc = ""
    Else
        txtErxc = Format(txtErxc, ".00#")
    End If
End Sub

Private Sub txtTime_Change()

    miTime = 1
    
    If IsNumeric(txtTime) Then
        If Len(txtTime) <= 3 Or Left(txtTime, 1) = "0" Then 'allows 155 or 0155 or entry
            lblTime = Format(txtTime, "00:##")  'to read 01:55 or
        Else                                  'or 55 entry to read 00:55
            lblTime = Format(txtTime, "##:##")  'formats 1155 entry to read 11:55
        End If
        
    ElseIf Not IsNumeric(txtTime) And txtTime <> "Time" And txtTime <> "" Then 'And txtTime <> "Time" Then
        miTime = 0
        MsgBox "Enter the Time as a continuous 3 or 4 digits in a 24 hour" & vbCrLf & "format, then click the Tab Key. Do NOT space between" & vbCrLf & "the Hours and Minutes or separate them with a colon.", vbOKOnly, "Enter Digits Only.  Do Not Separate With a Colon"
        txtTime.SelStart = 0 'begin selection at start
        txtTime.SelLength = 4 ' Len(txtTime)
    End If
 End Sub

Private Sub txtTime_Click()
    txtTime = "Time"
    txtTime.SelStart = 0 'begin selection at start
    txtTime.SelLength = 4 'Len(txtTime)
End Sub

Private Sub txtTime_gotFocus()
    txtTime.SelStart = 0 'begin selection at start
    txtTime.SelLength = 4 'Len(txtTime) 'selects # of characters
    lblTwr.Visible = True
    Label1.Visible = True
    Label11.Visible = False
  End Sub

Private Sub txtTime_LostFocus()
    txtTime.Visible = False 'to prevent confusion in seeing non-formatted time entry
    Label1.Visible = False
    Label11.Visible = True
End Sub

Private Sub txtVgrs_Change()

   If Not IsNumeric(txtVgrs) And txtVgrs <> "" Then
        MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
        txtVgrs = ""
        Exit Sub
    End If
    
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
        mnuDigitHighlights_Click
    End If
    pAddgrs
End Sub

Private Sub txtVgrs_DblClick()
    txtVgrs.SelStart = 0 'begin selection at start
    txtVgrs.SelLength = Len(txtVgrs) 'selects # of characters
End Sub

Private Sub txtVgrs_GotFocus()

    Dim sWrxcHeader As String
    Dim sWrxcLegal As String
    If miWrxc < Val(gStation2Min) Then
        sWrxcHeader = lblStation2 & " Power is LOW"
        sWrxcLegal = lblStation2 & " transmitter power is " & Format(miWrxc, "####") & _
        " watts. The minimum normal power is " & gStation2Min & " watts."
    ElseIf miWrxc > Val(gStation2Max) Then
        sWrxcHeader = lblStation2 & " Power is HIGH"
        sWrxcLegal = lblStation2 & " transmitter power is " & Format(miWrxc, "####") & _
        " watts. The maximum normal power is " & gStation2Max & " watts."
    End If
    
    If gStation2Min <> "" And gStation2Max <> "" And lblrxc <> "" And miStation2 = 1 And _
       ((miWrxc < (Val(gStation2Min) - (Val(gStation2Min) * 0.02))) Or (miWrxc > (Val(gStation2Max) + (Val(gStation2Max) * 0.02)))) Then
 
        miStation2 = 2
        
        MsgBox sWrxcLegal & vbCrLf & vbCrLf & _
        "Reminder: Compare efficiency values and power limits with the station" & vbCrLf & _
        "transmitter log to verify that this program is using the most recent data." _
        & vbCrLf & vbCrLf & "To read power limits, pause the mouse cursor over each station's call" _
        & vbCrLf & "letters for a readout of minimum and maximum limits for that station." _
        & vbCrLf & vbCrLf & "Use the 'Settings' menu to update power limits or efficiency values." & vbCrLf, _
        vbExclamation, sWrxcHeader
    End If

    txtVgrs.SelStart = Val(txtVolt3) 'begin selection at start
    txtVgrs.SelLength = Len(txtVgrs) 'selects # of characters
End Sub

Private Sub txtVgrs_LostFocus()
    pAddgrs
End Sub

Private Sub txtVgsk_Change()
   
    If Not IsNumeric(txtVgsk) And txtVgsk <> "" Then
        MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
        txtVgsk = ""
        Exit Sub
    End If
    
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
        mnuDigitHighlights_Click
    End If
    pAddgsk
End Sub

Private Sub txtVgsk_DblClick()
    txtVgsk.SelStart = 0 'begin selection at start
    txtVgsk.SelLength = Len(txtVgsk) 'selects # of characters
End Sub

Private Sub txtVgsk_GotFocus()

    Dim sWgrsHeader As String
    Dim sWgrsLegal As String
    If miWgrs < Val(gStation3Min) Then
        sWgrsHeader = lblStation3 & " Power is LOW"
        sWgrsLegal = lblStation3 & " transmitter power is " & Format(miWgrs, "####") & _
        " watts. The minimum normal power is " & gStation3Min & " watts."
    ElseIf miWgrs > Val(gStation3Max) Then
        sWgrsHeader = lblStation3 & " Power is HIGH"
        sWgrsLegal = lblStation3 & " transmitter power is " & Format(miWgrs, "####") & _
        " watts. The maximum normal power is " & gStation3Max & " watts."
    End If
    
    If gStation3Min <> "" And gStation3Max <> "" And lblgrs <> "" And miStation3 = 1 And _
       ((miWgrs < (Val(gStation3Min) - (Val(gStation3Min) * 0.02))) Or (miWgrs > (Val(gStation3Max) + (Val(gStation3Max) * 0.02)))) Then

        miStation3 = 2
        
        MsgBox sWgrsLegal & vbCrLf & vbCrLf & _
        "Reminder: Compare efficiency values and power limits with the station" & vbCrLf & _
        "transmitter log to verify that this program is using the most recent data." _
        & vbCrLf & vbCrLf & "To read power limits, pause the mouse cursor over each station's call" _
        & vbCrLf & "letters for a readout of minimum and maximum limits for that station." _
        & vbCrLf & vbCrLf & "Use the 'Settings' menu to update power limits or efficiency values." & vbCrLf, _
        vbExclamation, sWgrsHeader
    End If

    txtVgsk.SelStart = Val(txtVolt4) 'begin selection at start
    txtVgsk.SelLength = Len(txtVgsk) 'selects # of characters
End Sub

Private Sub txtVgsk_LostFocus()
    pAddgsk
End Sub

Private Sub txtVmnr_Change()

    If Not IsNumeric(txtVmnr) And txtVmnr <> "" Then
        MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
        txtVmnr = ""
        Exit Sub
    End If
    
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
    mnuSettingsEffic_Click
    End If

    If mnuDigitHighlights.Checked = True Then
        mnuDigitHighlights_Click
    End If
    pAddmnr
    
   End Sub

Private Sub txtVmnr_DblClick()
    txtVmnr.SelStart = 0 'begin selection at start
    txtVmnr.SelLength = Len(txtVmnr) 'selects # of characters
End Sub

Private Sub txtVmnr_GotFocus()

    Dim sWgskHeader As String
    Dim sWgskLegal As String
    If miWgsk < Val(gStation4Min) Then
        sWgskHeader = lblStation4 & " Power is LOW"
        sWgskLegal = lblStation4 & " transmitter power is " & Format(miWgsk, "####") & _
        " watts. The minimum normal power is " & gStation4Min & " watts."
    ElseIf miWgsk > Val(gStation4Max) Then
        sWgskHeader = lblStation4 & " Power is HIGH"
        sWgskLegal = lblStation4 & " transmitter power is " & Format(miWgsk, "####") & _
        " watts. The maximum normal power is " & gStation4Max & " watts."
    End If

    If gStation4Min <> "" And gStation4Max <> "" And lblgsk <> "" And miStation4 = 1 And _
        ((miWgsk < (Val(gStation4Min) - (Val(gStation4Min) * 0.02))) Or (miWgsk > (Val(gStation4Max) + (Val(gStation4Max) * 0.02)))) Then

        miStation4 = 2
        
        MsgBox sWgskLegal & vbCrLf & vbCrLf & _
        "Reminder: Compare efficiency values and power limits with the station" & vbCrLf & _
        "transmitter log to verify that this program is using the most recent data." _
        & vbCrLf & vbCrLf & "To read power limits, pause the mouse cursor over each station's call" _
        & vbCrLf & "letters for a readout of minimum and maximum limits for that station." _
        & vbCrLf & vbCrLf & "Use the 'Settings' menu to update power limits or efficiency values." & vbCrLf, _
        vbExclamation, sWgskHeader
    End If

    If lblStation1 = "" Or lblStation1 = "?" Then
        MsgBox "Station Call Letters required. Use 'Change Call Letters' under menu item 'Settings' to set Call Letters", vbOKOnly, "Call Letters Missing"
        lblStation1 = "?"
        txtVmnr.Text = "0"
        Exit Sub
    End If
    miTime = 0
    txtVmnr.SelStart = Val(txtVolt1) 'begin selection at start
    txtVmnr.SelLength = Len(txtVmnr) 'selects # of characters
End Sub

Private Sub txtVmnr_LostFocus()
    pAddmnr
End Sub

Private Sub txtVolt1_Change()

    If Not IsNumeric(txtVolt1) And txtVolt1 <> "" Then
        MsgBox "Non-Numeric Entry. Enter a single number of 5 or less.", vbOKOnly, "Entry Error"
        txtVolt1 = ""
        Exit Sub
    End If
    
    If txtVolt1 > "5" Then
        MsgBox "Maximum entry is 5. Choose in the range between 0 and 5." & vbCrLf & vbCrLf & _
        "With 0 all characters will be highlighted. With 5 no character will be highlighted.", vbOKOnly, "Entry Error"
        txtVolt1 = ""
        Exit Sub
    End If
End Sub

Private Sub txtVolt1_GotFocus()
    txtVolt1.SelLength = Len(txtVolt1) 'selects # of characters
End Sub

Private Sub txtVolt2_Change()

    If Not IsNumeric(txtVolt2) And txtVolt2 <> "" Then
        MsgBox "Non-Numeric Entry. Enter a single number of 5 or less.", vbOKOnly, "Entry Error"
        txtVolt2 = ""
        Exit Sub
    End If
    
    If txtVolt2 > "5" Then
        MsgBox "Maximum entry is 5. Choose in the range between 0 and 5." & vbCrLf & vbCrLf & _
        "With 0 all characters will be highlighted. With 5 no character will be highlighted.", vbOKOnly, "Entry Error"
        txtVolt2 = ""
        Exit Sub
    End If
End Sub

Private Sub txtVolt2_GotFocus()
    txtVolt2.SelLength = Len(txtVolt2) 'selects # of characters
End Sub

Private Sub txtVolt3_Change()

    If Not IsNumeric(txtVolt3) And txtVolt3 <> "" Then
        MsgBox "Non-Numeric Entry. Enter a single number of 5 or less.", vbOKOnly, "Entry Error"
        txtVolt3 = ""
        Exit Sub
    End If
    
    If txtVolt3 > "5" Then
        MsgBox "Maximum entry is 5. Choose in the range between 0 and 5." & vbCrLf & vbCrLf & _
        "With 0 all characters will be highlighted. With 5 no character will be highlighted.", vbOKOnly, "Entry Error"
        txtVolt3 = ""
        Exit Sub
    End If
End Sub

Private Sub txtVolt3_GotFocus()
    txtVolt3.SelLength = Len(txtVolt3) 'selects # of characters
End Sub

Private Sub txtVolt4_Change()

    If Not IsNumeric(txtVolt4) And txtVolt4 <> "" Then
        MsgBox "Non-Numeric Entry. Enter a single number of 5 or less.", vbOKOnly, "Entry Error"
        txtVolt4 = ""
        Exit Sub
    End If
    
    If txtVolt4 > "5" Then
        MsgBox "Maximum entry is 5. Choose in the range between 0 and 5." & vbCrLf & vbCrLf & _
        "With 0 all characters will be highlighted. With 5 no character will be highlighted.", vbOKOnly, "Entry Error"
        txtVolt4 = ""
        Exit Sub
    End If
End Sub

Private Sub txtVolt4_GotFocus()
    txtVolt4.SelLength = Len(txtVolt4) 'selects # of characters
End Sub

Private Sub txtVrxc_Change()

    If Not IsNumeric(txtVrxc) And txtVrxc <> "" Then
        MsgBox "Non-Numeric Entry", vbOKOnly, "Entry Error"
        txtVrxc = ""
        Exit Sub
    End If
    
    If mnuSettingsEffic.Checked = True Then 'closes Set Efficiencies
        mnuSettingsEffic_Click
    End If
    
    If mnuDigitHighlights.Checked = True Then
        mnuDigitHighlights_Click
    End If
    pAddrxc
    
End Sub

Private Sub pAddmnr()
    Dim iVmnr As Currency
    Dim iAmnr As Currency
    Dim iEmnr As Currency
    
    iVmnr = Val(txtVmnr)
    iAmnr = Val(txtAmnr)
    iEmnr = Val(txtEmnr)
    miWmnr = iVmnr * iAmnr * iEmnr
    
    If lblStation1.ForeColor <> &H808080 Then
        lblmnr.Caption = Format(miWmnr, "#,###")
    End If
    
    If Val(miWmnr) > 99999 Then
        MsgBox "Power exceeds 99,999 watts. Recheck data.", vbCritical, "Excessive Power"
        txtAmnr = ""
        Exit Sub
    End If

End Sub

Private Sub pAddrxc()
    Dim iVrxc As Currency
    Dim iArxc As Currency
    Dim iErxc As Currency
    iVrxc = Val(txtVrxc)
    iArxc = Val(txtArxc)
    iErxc = Val(txtErxc)
    miWrxc = iVrxc * iArxc * iErxc
    
    If lblStation2.ForeColor <> &H808080 Then
        lblrxc.Caption = Format(miWrxc, "#,###") 'miWrxc
    End If
    
    If Val(miWrxc) > 99999 Then
        MsgBox "Power exceeds 99,999 watts. Recheck data.", vbCritical, "Excessive Power"
        txtArxc = ""
        Exit Sub
    End If

End Sub

Private Sub pAddgrs()
    Dim iVgrs As Currency
    Dim iAgrs As Currency
    Dim iEgrs As Currency
    iVgrs = Val(txtVgrs)
    iAgrs = Val(txtAgrs)
    iEgrs = Val(txtEgrs)
    miWgrs = iVgrs * iAgrs * iEgrs
    
    If lblStation3.ForeColor <> &H808080 Then
        lblgrs.Caption = Format(miWgrs, "#,###")
    End If
    
    If Val(miWgrs) > 99999 Then
        MsgBox "Power exceeds 99,999 watts. Recheck data.", vbCritical, "Excessive Power"
        txtAgrs = ""
        Exit Sub
    End If
End Sub

Private Sub pAddgsk()
    Dim iVgsk As Currency
    Dim iAgsk As Currency
    Dim iEgsk As Currency
    iVgsk = Val(txtVgsk)
    iAgsk = Val(txtAgsk)
    iEgsk = Val(txtEgsk)
    miWgsk = iVgsk * iAgsk * iEgsk
    
    If lblStation4.ForeColor <> &H808080 Then
      lblgsk.Caption = Format(miWgsk, "#,###") 'miWgsk
    End If
    
    If Val(miWgsk) > 99999 Then
        MsgBox "Power exceeds 99,999 watts. Recheck data.", vbCritical, "Excessive Power"
        txtAgsk = ""
        Exit Sub
    End If
End Sub

Private Sub txtVrxc_DblClick()
    txtVrxc.SelStart = 0 'begin selection at start
    txtVrxc.SelLength = Len(txtVrxc) 'selects # of characters
End Sub

Private Sub txtVrxc_GotFocus()

    Dim sWmnrHeader As String
    Dim sWmnrLegal As String
    If miWmnr < Val(gStation1Min) Then
        sWmnrHeader = lblStation1 & " Power is LOW"
        sWmnrLegal = lblStation1 & " transmitter power is " & Format(miWmnr, "####") & _
        " watts. The minimum normal power is " & gStation1Min & " watts."
    ElseIf miWmnr > Val(gStation1Max) Then
        sWmnrHeader = lblStation1 & " Transmitter Power is HIGH"
        sWmnrLegal = lblStation1 & " power is " & Format(miWmnr, "####") & _
        " watts. The maximum normal power is " & gStation1Max & " watts."
    End If
    

    If gStation1Min <> "" And gStation1Max <> "" And lblmnr <> "" And miStation1 = 1 And _
        ((miWmnr < (Val(gStation1Min) - (Val(gStation1Min) * 0.02))) Or (miWmnr > (Val(gStation1Max) + (Val(gStation1Max) * 0.02)))) Then

        miStation1 = 2
        
        MsgBox sWmnrLegal & vbCrLf & vbCrLf & _
        "Reminder: Compare efficiency values and power limits with the station" & vbCrLf & _
        "transmitter log to verify that this program is using the most recent data." _
        & vbCrLf & vbCrLf & "To read power limits, pause the mouse cursor over each station's call" _
        & vbCrLf & "letters for a readout of minimum and maximum limits for that station." _
        & vbCrLf & vbCrLf & "Use the 'Settings' menu to update power limits or efficiency values." & vbCrLf, _
        vbExclamation, sWmnrHeader
    End If
    
    txtVrxc.SelStart = Val(txtVolt2) 'begin selection at start
    txtVrxc.SelLength = Len(txtVrxc) 'selects # of characters
End Sub

Private Sub txtVrxc_LostFocus()
    pAddrxc
End Sub
