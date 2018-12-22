VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimeRemain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estimates of On-Air-Time Remaining & Program Total Time (music + announce time + spots)   -   F2"
   ClientHeight    =   9780
   ClientLeft      =   165
   ClientTop       =   645
   ClientWidth     =   11820
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TimeRemain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "TimeRemain.frx":0442
   ScaleHeight     =   9780
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame11 
      Caption         =   "Number"
      ForeColor       =   &H00000080&
      Height          =   720
      Left            =   7965
      TabIndex        =   182
      Top             =   3135
      Width           =   780
      Begin VB.TextBox txtSpotsS 
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
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   135
         MaxLength       =   3
         TabIndex        =   183
         TabStop         =   0   'False
         ToolTipText     =   " Enter the number of PSA's and spots inserted in the time period "
         Top             =   175
         Width           =   510
      End
      Begin VB.Label lblSpotSecs 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   120
         TabIndex        =   184
         Top             =   495
         Width           =   570
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Seconds"
      Height          =   1215
      Left            =   7965
      TabIndex        =   179
      Top             =   3900
      Width           =   780
      Begin VB.TextBox txtSpotLength 
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
         Height          =   285
         Left            =   120
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   195
         Width           =   510
      End
      Begin VB.TextBox txtIntro 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   180
         TabStop         =   0   'False
         ToolTipText     =   "Overwrite to change temporarily. Select 'Announce-Times' menu to save the change. Doutle-Click to reduce value by ten seconds."
         Top             =   720
         Width           =   510
      End
   End
   Begin VB.Frame fraAdjustedTime 
      Height          =   345
      Left            =   9975
      TabIndex        =   171
      Top             =   1755
      Visible         =   0   'False
      Width           =   1050
      Begin VB.Label lblHourAdj 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   45
         TabIndex        =   174
         ToolTipText     =   " Double-click to adjust hour from current hour to hour forward or hour backward "
         Top             =   105
         Width           =   300
      End
      Begin VB.Label lblSecondsLeft 
         Caption         =   ": 00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   645
         TabIndex        =   173
         Top             =   105
         Width           =   330
      End
      Begin VB.Label lblMinutesLeft 
         Alignment       =   1  'Right Justify
         Caption         =   ":00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   315
         TabIndex        =   172
         Top             =   105
         Width           =   300
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lineup"
      ForeColor       =   &H00000080&
      Height          =   1155
      Left            =   7245
      TabIndex        =   167
      Top             =   1890
      Width           =   2280
      Begin VB.CommandButton cmdRestoreEntries 
         Caption         =   "&Restore Lineup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   265
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   495
         Width           =   2055
      End
      Begin VB.CommandButton cmdClearLineUp 
         Caption         =   "Clear &Lineup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   265
         Left            =   105
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton cmdClearAnncTimes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear Announce Times"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   265
         Left            =   105
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   810
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdAdjTime 
      BackColor       =   &H00000000&
      Caption         =   "If Computer Clock Time is Incorrect, Adjust Program Time to Compensate"
      Height          =   870
      Left            =   9660
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   2085
      Width           =   1670
   End
   Begin VB.Frame fraAdjustTime 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   9565
      TabIndex        =   154
      ToolTipText     =   " Double-Click for Instructions "
      Top             =   1890
      Visible         =   0   'False
      Width           =   1860
      Begin VB.CommandButton cmdSecMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1500
         TabIndex        =   161
         TabStop         =   0   'False
         ToolTipText     =   "Subtract 5"
         Top             =   585
         Width           =   210
      End
      Begin VB.CommandButton cmdSecPlus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1500
         TabIndex        =   160
         TabStop         =   0   'False
         ToolTipText     =   "Add 5"
         Top             =   293
         Width           =   210
      End
      Begin VB.CommandButton cmdMinMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   159
         TabStop         =   0   'False
         ToolTipText     =   "Subtract 1"
         Top             =   585
         Width           =   210
      End
      Begin VB.CommandButton cmdMinPlus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   158
         TabStop         =   0   'False
         ToolTipText     =   "Add 1"
         Top             =   293
         Width           =   210
      End
      Begin VB.TextBox txtSecAdj 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   915
         MaxLength       =   3
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   375
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtMinAdj 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   450
         MaxLength       =   3
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   375
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "clear"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   45
         TabIndex        =   178
         Top             =   870
         Width           =   435
      End
      Begin VB.Label lblAdjustTimeExit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "cancel"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1320
         TabIndex        =   165
         Top             =   870
         Width           =   465
      End
      Begin VB.Label lblMinutes 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Caption         =   "min"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   164
         Top             =   195
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblSeconds 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Caption         =   "sec"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   975
         TabIndex        =   163
         Top             =   195
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Caption         =   "clock error adjustment"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   495
         TabIndex        =   162
         ToolTipText     =   " Double-Click for Instructions "
         Top             =   705
         Visible         =   0   'False
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9930
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   7830
      Width           =   1050
   End
   Begin VB.Frame fraAddTime 
      Caption         =   "Pages"
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
      Height          =   990
      Left            =   9495
      TabIndex        =   149
      Top             =   6735
      Width           =   1935
      Begin VB.CommandButton cmdPower 
         Caption         =   "Transmitter Log  F3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   585
         Width           =   1575
      End
      Begin VB.CommandButton cmdPlanner 
         Caption         =   "Music &Planner F1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   210
         Width           =   1575
      End
   End
   Begin VB.Frame fraFrame3 
      Caption         =   "Planned Program Time"
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
      Height          =   645
      Left            =   9450
      TabIndex        =   145
      ToolTipText     =   "Enter in minutes the total program time (music plus talk plus ID) planned for the hour, normally 60 minutes."
      Top             =   5190
      Width           =   1980
      Begin VB.TextBox txtBlock 
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
         Height          =   285
         Left            =   135
         MaxLength       =   4
         MultiLine       =   -1  'True
         TabIndex        =   147
         TabStop         =   0   'False
         ToolTipText     =   " Enter in minutes the total program time (music plus talk plus ID) planned for the hour, normally 60 minutes. "
         Top             =   240
         Width           =   510
      End
      Begin VB.CommandButton cmdClearBlock 
         Caption         =   "Cl&ear"
         Height          =   255
         Left            =   1365
         TabIndex        =   146
         TabStop         =   0   'False
         ToolTipText     =   "Clears 'Planned Time' entry only "
         Top             =   255
         Width           =   450
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Minutes"
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
         Height          =   210
         Left            =   660
         TabIndex        =   148
         ToolTipText     =   " Enter in minutes the total program time (music plus talk plus ID) planned for the hour, normally 60 minutes. "
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Pri&nt Music Lineup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9690
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   6300
      Width           =   1545
   End
   Begin VB.CommandButton cmdAddTime 
      BackColor       =   &H80000018&
      Caption         =   "AddTi&me Calculator F9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   143
      TabStop         =   0   'False
      ToolTipText     =   " Calculator for adding times entered as minutes & seconds "
      Top             =   5925
      Width           =   1845
   End
   Begin VB.TextBox txtBackAnnc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5820
      MaxLength       =   3
      TabIndex        =   141
      TabStop         =   0   'False
      ToolTipText     =   $"TimeRemain.frx":074C
      Top             =   1590
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CheckBox chkBackAnnc 
      Caption         =   "Check if CD will NOT be back-announced"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   6495
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Frame fraIntro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   7395
      Left            =   -2730
      TabIndex        =   47
      Top             =   -1365
      Visible         =   0   'False
      Width           =   2910
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5520
         Left            =   135
         TabIndex        =   67
         Top             =   1740
         Width           =   2640
         Begin VB.CommandButton cmdDefaults 
            Caption         =   "Use Defaults"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   668
            TabIndex        =   125
            TabStop         =   0   'False
            ToolTipText     =   "Sets estimated times to default values."
            Top             =   4500
            Width           =   1305
         End
         Begin VB.CommandButton cmdCloseTimeSet 
            Caption         =   " Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   915
            TabIndex        =   124
            Top             =   4995
            Width           =   810
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save Your Changes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   390
            TabIndex        =   70
            Top             =   4005
            Width           =   1860
         End
         Begin VB.TextBox txtSpotLengthSetting 
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
            Height          =   300
            Left            =   1065
            MaxLength       =   3
            TabIndex        =   69
            Top             =   2130
            Width           =   510
         End
         Begin VB.TextBox txtIntroSetting 
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
            Height          =   300
            Left            =   1065
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   68
            Top             =   1035
            Width           =   510
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00C0C000&
            Index           =   1
            X1              =   180
            X2              =   2460
            Y1              =   3765
            Y2              =   3765
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Seconds"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1650
            TabIndex        =   83
            Top             =   2190
            Width           =   570
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            Caption         =   "What is an average length in seconds of spot announcement and weather INSERTS, etc.?"
            Height          =   525
            Left            =   120
            TabIndex        =   82
            Top             =   1545
            Width           =   2400
         End
         Begin VB.Label lblSelectionTime 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            Caption         =   "What is the length in seconds of the average total time needed to introduce and back-announce a MUSIC SELECTION?"
            Height          =   735
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   2400
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            Caption         =   $"TimeRemain.frx":07D9
            Height          =   1185
            Left            =   150
            TabIndex        =   72
            Top             =   2580
            Width           =   2340
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C000&
            Index           =   0
            X1              =   180
            X2              =   2445
            Y1              =   1410
            Y2              =   1410
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00C0C000&
            Index           =   0
            X1              =   173
            X2              =   2453
            Y1              =   2490
            Y2              =   2490
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Seconds"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1650
            TabIndex        =   71
            Top             =   1095
            Width           =   570
         End
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "(Note: Select 'Use Defaults' to replace current values with default values.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   360
         TabIndex        =   56
         Top             =   930
         Width           =   2205
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "To compute a total Estimated Announce Time, the following estimated times are needed:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   135
         TabIndex        =   48
         Top             =   210
         Width           =   2640
      End
   End
   Begin VB.TextBox txtMinute1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5745
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   360
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   990
      Left            =   375
      TabIndex        =   86
      Top             =   8355
      Width           =   6690
      Begin VB.TextBox txtSecond3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5085
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         ToolTipText     =   " Use the Tab key to advance as necessary thru text boxes "
         Top             =   525
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtMinute3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4590
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         ToolTipText     =   " Use the Tab key to advance as necessary thru text boxes "
         Top             =   525
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblAnncSec 
         Alignment       =   2  'Center
         Caption         =   "sec"
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
         Height          =   225
         Left            =   5115
         TabIndex        =   127
         Top             =   255
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblAnncMin 
         Alignment       =   2  'Center
         Caption         =   "min"
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
         Height          =   225
         Left            =   4680
         TabIndex        =   126
         Top             =   255
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image imgMic 
         Appearance      =   0  'Flat
         Height          =   690
         Left            =   135
         Picture         =   "TimeRemain.frx":08A1
         Stretch         =   -1  'True
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lblS 
         Caption         =   "10 spots"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5670
         TabIndex        =   123
         Top             =   510
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "You can replace the program's estimated announce time with your estimate of the total announce time you will need:"
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
         Left            =   1245
         TabIndex        =   88
         ToolTipText     =   " Double-Click to clear Min/Sec data "
         Top             =   255
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Label lblMinSec3Div 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4980
         TabIndex        =   87
         Top             =   525
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape shpTime3 
         Height          =   300
         Left            =   4545
         Top             =   510
         Visible         =   0   'False
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Time Remaining"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1965
      Left            =   3210
      TabIndex        =   84
      Top             =   6240
      Width           =   6090
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   " • 30 secs will remain for show closeout && station ID    "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   735
         TabIndex        =   190
         Top             =   735
         Width           =   4665
      End
      Begin VB.Label lblSspot 
         Alignment       =   2  'Center
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   5760
         TabIndex        =   185
         Top             =   345
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape shpRunOver 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Height          =   270
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   1575
         Visible         =   0   'False
         Width           =   4200
      End
      Begin VB.Label lblRemain60 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblRemain60"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2520
         TabIndex        =   153
         Top             =   1590
         Width           =   1080
      End
      Begin VB.Label lblRemain30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblRemain30"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2527
         TabIndex        =   129
         Top             =   1200
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblProgramRemain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblProgramRemain"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   420
         TabIndex        =   128
         Top             =   390
         Width           =   5295
      End
      Begin VB.Image imgMusic 
         Height          =   615
         Left            =   75
         Picture         =   "TimeRemain.frx":3D8C2
         Stretch         =   -1  'True
         Top             =   285
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin VB.Frame fraAnnounce 
      Caption         =   "Estimated Announce Time Option"
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
      Height          =   1080
      Left            =   375
      TabIndex        =   78
      Top             =   7125
      Width           =   2715
      Begin VB.CheckBox chkAnnounce 
         Alignment       =   1  'Right Justify
         Caption         =   "C&heck if you do NOT want to include estimated announce times of 50 sec for each selecton and 30 sec for program closeout."
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   135
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   255
         Width           =   2430
      End
   End
   Begin MSComctlLib.StatusBar StaRemain 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   77
      Top             =   9435
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   900
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   14552
            MinWidth        =   14552
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3440
            MinWidth        =   3440
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1711
            MinWidth        =   1711
            TextSave        =   "1/10/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSecond2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6150
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   630
      Width           =   420
   End
   Begin VB.TextBox txtMinute2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5715
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   630
      Width           =   420
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10920
      Top             =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Music Lineup"
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
      Height          =   3375
      Left            =   1920
      TabIndex        =   46
      Top             =   1950
      Width           =   4770
      Begin VB.TextBox txtAnnc10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   2840
         MaxLength       =   3
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   2358
         Width           =   425
      End
      Begin VB.TextBox txtAnnc9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   2840
         MaxLength       =   3
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   2015
         Width           =   425
      End
      Begin VB.TextBox txtAnnc8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   2840
         MaxLength       =   3
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1672
         Width           =   425
      End
      Begin VB.TextBox txtAnnc7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   2840
         MaxLength       =   3
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   1329
         Width           =   425
      End
      Begin VB.TextBox txtAnnc6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   2840
         MaxLength       =   3
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   986
         Width           =   425
      End
      Begin VB.TextBox txtAnnc5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   2840
         MaxLength       =   3
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   643
         Width           =   425
      End
      Begin VB.TextBox txtAnnc4 
         Alignment       =   2  'Center
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
         Height          =   210
         Left            =   2840
         MaxLength       =   3
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   300
         Width           =   425
      End
      Begin VB.TextBox txtAnnc11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   210
         Left            =   2840
         MaxLength       =   3
         TabIndex        =   131
         TabStop         =   0   'False
         ToolTipText     =   "Annc Time boxes 7 & 8 accept temporary negative numbers"
         Top             =   2707
         Width           =   425
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Index           =   7
         Left            =   3375
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   2715
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Index           =   6
         Left            =   3375
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   2370
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   3375
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   2025
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   3375
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   1680
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   3375
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   1335
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   3375
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   990
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   3375
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   645
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   3375
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   300
         Width           =   180
      End
      Begin VB.TextBox txtCD11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   5
         TabIndex        =   33
         Top             =   2655
         Width           =   630
      End
      Begin VB.TextBox txtCD10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   5
         TabIndex        =   29
         Top             =   2310
         Width           =   630
      End
      Begin VB.TextBox txtCD9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   5
         TabIndex        =   25
         Top             =   1965
         Width           =   630
      End
      Begin VB.TextBox txtCD8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   5
         TabIndex        =   21
         Top             =   1620
         Width           =   630
      End
      Begin VB.TextBox txtCD7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   5
         TabIndex        =   17
         Top             =   1275
         Width           =   630
      End
      Begin VB.TextBox txtCD6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   5
         TabIndex        =   13
         Top             =   930
         Width           =   630
      End
      Begin VB.TextBox txtCD5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         MaxLength       =   5
         TabIndex        =   9
         Top             =   585
         Width           =   630
      End
      Begin VB.TextBox txtCD4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2100
         MaxLength       =   5
         TabIndex        =   5
         Top             =   255
         Width           =   630
      End
      Begin VB.TextBox txtSecond11 
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
         Height          =   300
         Left            =   4125
         MaxLength       =   2
         TabIndex        =   35
         Top             =   2655
         Width           =   465
      End
      Begin VB.TextBox txtMinute11 
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
         Height          =   300
         Left            =   3660
         MaxLength       =   2
         TabIndex        =   34
         Top             =   2655
         Width           =   465
      End
      Begin VB.TextBox txtComposer11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   25
         TabIndex        =   32
         Top             =   2655
         Width           =   1935
      End
      Begin VB.TextBox txtComposer4 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   90
         MaxLength       =   25
         TabIndex        =   4
         Top             =   255
         Width           =   1935
      End
      Begin VB.TextBox txtSecond10 
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
         Height          =   300
         Left            =   4125
         MaxLength       =   2
         TabIndex        =   31
         Top             =   2307
         Width           =   465
      End
      Begin VB.TextBox txtSecond9 
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
         Height          =   300
         Left            =   4125
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1965
         Width           =   465
      End
      Begin VB.TextBox txtSecond5 
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
         Height          =   300
         Left            =   4125
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Use the Tab key to advance as necessary thru text boxes"
         Top             =   597
         Width           =   465
      End
      Begin VB.TextBox txtSecond4 
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
         Height          =   300
         Left            =   4125
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Enter seconds."
         Top             =   255
         Width           =   465
      End
      Begin VB.TextBox txtSecond6 
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
         Height          =   300
         Left            =   4125
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   939
         Width           =   465
      End
      Begin VB.TextBox txtSecond7 
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
         Height          =   300
         Left            =   4125
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1281
         Width           =   465
      End
      Begin VB.TextBox txtSecond8 
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
         Height          =   300
         Left            =   4125
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   1623
         Width           =   465
      End
      Begin VB.TextBox txtMinute8 
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
         Height          =   300
         Left            =   3660
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "Enter CD/LP planing time (minutes)"
         Top             =   1620
         Width           =   465
      End
      Begin VB.TextBox txtMinute7 
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
         Height          =   300
         Left            =   3660
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   1275
         Width           =   465
      End
      Begin VB.TextBox txtMinute4 
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
         Height          =   300
         Left            =   3660
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Enter CD/LP planing time (minutes)"
         Top             =   255
         Width           =   465
      End
      Begin VB.TextBox txtMinute5 
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
         Height          =   300
         Left            =   3660
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   465
      End
      Begin VB.TextBox txtMinute6 
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
         Height          =   300
         Left            =   3660
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Enter CD/LP planing time (minutes)"
         Top             =   945
         Width           =   465
      End
      Begin VB.TextBox txtMinute9 
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
         Height          =   300
         Left            =   3660
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1965
         Width           =   465
      End
      Begin VB.TextBox txtMinute10 
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
         Height          =   300
         Left            =   3660
         MaxLength       =   2
         TabIndex        =   30
         ToolTipText     =   "Enter CD/LP planing time (minutes)"
         Top             =   2310
         Width           =   465
      End
      Begin VB.TextBox txtComposer9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   25
         TabIndex        =   24
         Top             =   1965
         Width           =   1935
      End
      Begin VB.TextBox txtComposer8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   25
         TabIndex        =   20
         Top             =   1620
         Width           =   1935
      End
      Begin VB.TextBox txtComposer7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   25
         TabIndex        =   16
         Top             =   1275
         Width           =   1935
      End
      Begin VB.TextBox txtComposer6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   25
         TabIndex        =   12
         Top             =   930
         Width           =   1935
      End
      Begin VB.TextBox txtComposer5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   25
         TabIndex        =   8
         Top             =   585
         Width           =   1935
      End
      Begin VB.TextBox txtComposer10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   25
         TabIndex        =   28
         Top             =   2310
         Width           =   1935
      End
      Begin VB.Frame Frame10 
         Height          =   2790
         Left            =   2055
         TabIndex        =   187
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Sec"
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   2895
         TabIndex        =   139
         Top             =   2940
         Width           =   315
      End
      Begin VB.Label lblAnncTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   " Annc Time"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2700
         MouseIcon       =   "TimeRemain.frx":3EB86
         MousePointer    =   1  'Arrow
         TabIndex        =   85
         Top             =   15
         Width           =   720
      End
      Begin VB.Label Label28 
         Caption         =   "Sec"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4252
         TabIndex        =   102
         Top             =   15
         Width           =   240
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Min"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3772
         TabIndex        =   101
         Top             =   15
         Width           =   240
      End
      Begin VB.Shape shpLink 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         Height          =   120
         Left            =   1125
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Overwrite Annc Times to change"
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
         Height          =   195
         Left            =   825
         TabIndex        =   49
         Top             =   3120
         Width           =   2565
      End
      Begin VB.Label lblTotalS 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "lblTotalS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   3750
         TabIndex        =   58
         Top             =   3045
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "CD#"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   2245
         MouseIcon       =   "TimeRemain.frx":3EE90
         MousePointer    =   99  'Custom
         TabIndex        =   75
         ToolTipText     =   "Double-Click to Deactivate/Activate CD# entry boxes Tab Stop"
         Top             =   15
         Width           =   330
      End
   End
   Begin VB.CheckBox ck10Played 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   180
      Left            =   1635
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4320
      Width           =   165
   End
   Begin VB.CheckBox ck9Played 
      Caption         =   "Check6"
      Enabled         =   0   'False
      Height          =   180
      Left            =   1635
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3975
      Width           =   165
   End
   Begin VB.CheckBox ck8Played 
      Caption         =   "Check5"
      Enabled         =   0   'False
      Height          =   180
      Left            =   1635
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3630
      Width           =   165
   End
   Begin VB.CheckBox ck7Played 
      Caption         =   "Check4"
      Enabled         =   0   'False
      Height          =   180
      Left            =   1635
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3285
      Width           =   165
   End
   Begin VB.CheckBox ck6Played 
      Caption         =   "Check3"
      Enabled         =   0   'False
      Height          =   180
      Left            =   1635
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2955
      Width           =   165
   End
   Begin VB.CheckBox ck5Played 
      Caption         =   "Check2"
      Enabled         =   0   'False
      Height          =   180
      Left            =   1635
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2610
      Width           =   165
   End
   Begin VB.CheckBox ck4Played 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   180
      Left            =   1635
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2250
      Width           =   165
   End
   Begin VB.TextBox txtSecond1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6180
      MaxLength       =   3
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   360
   End
   Begin VB.CheckBox ck11Played 
      Enabled         =   0   'False
      Height          =   180
      Left            =   1635
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4665
      Width           =   165
   End
   Begin VB.CommandButton cmdSetTime 
      BackColor       =   &H80000018&
      Caption         =   "Note the time remaining on current CD then click space bar or Enter key"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   250
      Visible         =   0   'False
      Width           =   5355
   End
   Begin VB.CommandButton cmdClearPads 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear Checks"
      Height          =   240
      Left            =   840
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   104
      TabStop         =   0   'False
      ToolTipText     =   "Clear CD played check boxes"
      Top             =   4935
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame8 
      Caption         =   "Seconds for Show Closeout && ID"
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
      Height          =   825
      Left            =   375
      TabIndex        =   80
      Top             =   6247
      Width           =   2715
      Begin VB.TextBox txtCloseOut 
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1755
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "sec"
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2295
         TabIndex        =   100
         Top             =   375
         Width           =   330
      End
      Begin VB.Label Label24 
         Caption         =   "Overwrite to change time allocated for show closeout && station ID"
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   150
         TabIndex        =   99
         Top             =   195
         Width           =   1545
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Stopwatch: Alternate Source for Program Timing"
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
      Height          =   990
      Left            =   7200
      MouseIcon       =   "TimeRemain.frx":3EFE2
      TabIndex        =   90
      ToolTipText     =   " Is this time accurate? Click for additional information. "
      Top             =   8340
      Width           =   4230
      Begin VB.CommandButton cmdReadMe 
         Caption         =   "Read Me"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   615
         Width           =   1020
      End
      Begin VB.CommandButton cmdStopwatch 
         Caption         =   "Stopwatch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1400
         MousePointer    =   1  'Arrow
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   615
         Width           =   1020
      End
      Begin VB.CheckBox chkStopWatch 
         Caption         =   "Use Stopwatch for Program Timing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   1200
         MousePointer    =   1  'Arrow
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   270
         Width           =   2940
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   270
         MousePointer    =   1  'Arrow
         TabIndex        =   93
         Top             =   315
         Width           =   750
      End
      Begin VB.Label lblClock 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2940
         MouseIcon       =   "TimeRemain.frx":3F2EC
         TabIndex        =   92
         ToolTipText     =   "Is this time accurate? Click for additional information."
         Top             =   570
         Width           =   45
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Convert min to sec"
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
      Height          =   840
      Left            =   375
      TabIndex        =   191
      Top             =   5325
      Width           =   1800
      Begin VB.CommandButton cmdConvert 
         BackColor       =   &H00404040&
         Height          =   240
         Left            =   1215
         MaskColor       =   &H00808080&
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   465
         Width           =   255
      End
      Begin VB.TextBox txtConvert 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   180
         MaxLength       =   3
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   420
         Width           =   530
      End
      Begin VB.Label lblMinSec 
         Caption         =   "min"
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   780
         TabIndex        =   195
         Top             =   495
         Width           =   270
      End
      Begin VB.Label Label8 
         Caption         =   "enter min ---- then click"
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   195
         TabIndex        =   194
         Top             =   225
         Width           =   1485
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Set Current Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1185
      Left            =   6975
      MousePointer    =   1  'Arrow
      TabIndex        =   196
      Top             =   30
      Width           =   4455
      Begin VB.CommandButton cmdRestoreTimes 
         Caption         =   "Rest&ore  Times"
         Height          =   280
         Left            =   2910
         MousePointer    =   1  'Arrow
         TabIndex        =   198
         TabStop         =   0   'False
         Top             =   270
         Width           =   1380
      End
      Begin VB.CommandButton cmdClearTimes 
         Caption         =   "Clear Times"
         Height          =   280
         Left            =   2910
         MousePointer    =   1  'Arrow
         TabIndex        =   197
         TabStop         =   0   'False
         Top             =   660
         Width           =   1380
      End
      Begin VB.CommandButton cmdSystemTime 
         BackColor       =   &H80000018&
         Caption         =   "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
         Height          =   870
         Left            =   105
         MouseIcon       =   "TimeRemain.frx":3F5F6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   225
         Width           =   1830
      End
      Begin VB.Image imgClockSetTime 
         Height          =   765
         Left            =   2025
         Picture         =   "TimeRemain.frx":3F900
         Stretch         =   -1  'True
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame fraTimeAdjust2 
      Caption         =   "Clock Error Adjustment"
      ForeColor       =   &H00000080&
      Height          =   1155
      Left            =   9570
      TabIndex        =   199
      Top             =   1890
      Width           =   1860
   End
   Begin VB.Shape Shape1 
      Height          =   300
      Left            =   5715
      Top             =   270
      Width           =   840
   End
   Begin VB.Label lblLinked 
      Caption         =   "Lineup linked to Planning Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   1995
      MouseIcon       =   "TimeRemain.frx":4598F
      MousePointer    =   99  'Custom
      TabIndex        =   186
      ToolTipText     =   "Double-click to break link to Planner page."
      Top             =   1755
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Caption         =   "10 spots"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   7290
      TabIndex        =   177
      Top             =   3390
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label30 
      Caption         =   "Label30"
      Height          =   30
      Left            =   7980
      TabIndex        =   176
      Top             =   5610
      Width           =   15
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Height          =   495
      Left            =   5145
      TabIndex        =   175
      Top             =   4530
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "• Seconds allotted to back-announce current CD:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1800
      TabIndex        =   142
      Top             =   1590
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Image imgCheck2 
      Height          =   195
      Left            =   1650
      Picture         =   "TimeRemain.frx":45AE1
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgCheck 
      Height          =   270
      Left            =   480
      Picture         =   "TimeRemain.frx":45F23
      Stretch         =   -1  'True
      Top             =   930
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHour 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5460
      TabIndex        =   130
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgOnAirSign 
      Height          =   255
      Left            =   2145
      Picture         =   "TimeRemain.frx":46365
      Stretch         =   -1  'True
      Top             =   1245
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   225
      Picture         =   "TimeRemain.frx":46A43
      Top             =   -45
      Width           =   480
   End
   Begin VB.Image imgClock 
      Height          =   480
      Left            =   2370
      Picture         =   "TimeRemain.frx":46E85
      Top             =   5535
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDisc 
      Height          =   405
      Left            =   375
      Picture         =   "TimeRemain.frx":472C7
      Stretch         =   -1  'True
      Top             =   525
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   4980
      Top             =   4455
      Width           =   1215
   End
   Begin VB.Label lblCurrentTime 
      Caption         =   " 1. Click the 'Set Current Time' button to enter the current time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   870
      TabIndex        =   122
      Top             =   360
      Width           =   4470
   End
   Begin VB.Label Label32 
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
      Height          =   225
      Index           =   7
      Left            =   6720
      TabIndex        =   113
      Top             =   4695
      Width           =   390
   End
   Begin VB.Label Label32 
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
      Height          =   225
      Index           =   6
      Left            =   6720
      TabIndex        =   112
      Top             =   4350
      Width           =   390
   End
   Begin VB.Label Label32 
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
      Height          =   225
      Index           =   5
      Left            =   6720
      TabIndex        =   111
      Top             =   4005
      Width           =   390
   End
   Begin VB.Label Label32 
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
      Height          =   225
      Index           =   4
      Left            =   6720
      TabIndex        =   110
      Top             =   3660
      Width           =   390
   End
   Begin VB.Label Label32 
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
      Height          =   225
      Index           =   3
      Left            =   6720
      TabIndex        =   109
      Top             =   3300
      Width           =   390
   End
   Begin VB.Label Label32 
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
      Height          =   225
      Index           =   2
      Left            =   6720
      TabIndex        =   108
      Top             =   2955
      Width           =   390
   End
   Begin VB.Label Label32 
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
      Height          =   225
      Index           =   1
      Left            =   6720
      TabIndex        =   107
      Top             =   2610
      Width           =   390
   End
   Begin VB.Label Label32 
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
      Height          =   225
      Index           =   0
      Left            =   6720
      TabIndex        =   106
      Top             =   2265
      Width           =   390
   End
   Begin VB.Label lblAddTime 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   450
      TabIndex        =   105
      Top             =   8190
      Width           =   45
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "1"
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
      Height          =   225
      Left            =   1335
      MouseIcon       =   "TimeRemain.frx":6C1A9
      TabIndex        =   103
      Top             =   2250
      Width           =   240
   End
   Begin VB.Label lblEndTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblEndTime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   5385
      TabIndex        =   52
      Top             =   1230
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label13 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6900
      TabIndex        =   95
      Top             =   1260
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Currently playing CD ends at"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3240
      TabIndex        =   94
      Top             =   1260
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblProgramInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblProgramInfo  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5500
      TabIndex        =   89
      Top             =   5790
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "Average time in seconds allotted for each spot announcement, promo, PSA, weather insert, etc. (overwite to change)"
      ForeColor       =   &H00404040&
      Height          =   555
      Left            =   8820
      TabIndex        =   81
      Top             =   3945
      Width           =   2610
   End
   Begin VB.Label lblExport 
      Caption         =   "Double-Click the LINE NUMBER of the line to be copied to the  Planning Page lineup"
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   375
      MouseIcon       =   "TimeRemain.frx":6C2FB
      MousePointer    =   1  'Arrow
      TabIndex        =   76
      ToolTipText     =   "Double-click to close copy feature."
      Top             =   2895
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   1335
      MouseIcon       =   "TimeRemain.frx":6C44D
      TabIndex        =   66
      Top             =   4650
      Width           =   240
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "7"
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
      Height          =   225
      Left            =   1335
      MouseIcon       =   "TimeRemain.frx":6C59F
      TabIndex        =   64
      Top             =   4290
      Width           =   240
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "6"
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
      Height          =   225
      Left            =   1335
      MouseIcon       =   "TimeRemain.frx":6C6F1
      TabIndex        =   63
      Top             =   3960
      Width           =   240
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "5"
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
      Height          =   225
      Left            =   1335
      MouseIcon       =   "TimeRemain.frx":6C843
      TabIndex        =   62
      Top             =   3615
      Width           =   240
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "4"
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
      Height          =   225
      Left            =   1335
      MouseIcon       =   "TimeRemain.frx":6C995
      TabIndex        =   61
      Top             =   3270
      Width           =   240
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "3"
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
      Height          =   225
      Left            =   1335
      MouseIcon       =   "TimeRemain.frx":6CAE7
      TabIndex        =   60
      Top             =   2925
      Width           =   240
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "2"
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
      Height          =   225
      Left            =   1335
      MouseIcon       =   "TimeRemain.frx":6CC39
      TabIndex        =   59
      Top             =   2580
      Width           =   240
   End
   Begin VB.Label lblAnnounceTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "lblAnnounceTime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   5447
      TabIndex        =   57
      ToolTipText     =   " Estimated announce time consists of music introductions and back announcements, spots, and the hour's closeout remarks. "
      Top             =   5490
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5745
      TabIndex        =   55
      Top             =   45
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6105
      TabIndex        =   54
      Top             =   300
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Sec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6150
      TabIndex        =   53
      Top             =   45
      Width           =   375
   End
   Begin VB.Label lblAverageTime 
      Caption         =   "Average time in seconds allotted to introduce plus back-announce each music selection  (overwrite to change)"
      ForeColor       =   &H00404040&
      Height          =   585
      Left            =   8820
      TabIndex        =   51
      ToolTipText     =   " Overwrite to change the seconds allotted for Introduction & Back-Announce "
      Top             =   4575
      Width           =   2490
   End
   Begin VB.Label lblSpots 
      Caption         =   $"TimeRemain.frx":6CD8B
      ForeColor       =   &H00000080&
      Height          =   765
      Left            =   8820
      TabIndex        =   50
      ToolTipText     =   " Select 'Announce-Times' menu to change seconds allotted for announcements "
      Top             =   3135
      Width           =   2220
   End
   Begin VB.Label Label14 
      Caption         =   "To Determine the Time the Current CD Ends && Air Time Remaining:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   765
      TabIndex        =   45
      Top             =   45
      Width           =   4785
   End
   Begin VB.Label Label10 
      Caption         =   " 3. Check lineup box below by any CD played or currently playing:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   870
      TabIndex        =   43
      Top             =   930
      Width           =   4785
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   " 2. Enter the time remaining on the current CD at the above time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   870
      TabIndex        =   36
      Top             =   660
      Width           =   4635
   End
   Begin VB.Menu mnuPage 
      Caption         =   "P&age"
      Begin VB.Menu mnuPagePlanner 
         Caption         =   "&Music Planning Page..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuPageXmitter 
         Caption         =   "&Tramsmitter Log Page..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuPagePrevious 
         Caption         =   "&Previous Page..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuPageSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageAddTime 
         Caption         =   "&AddTime Calculator..."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuPageSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageStopWatch 
         Caption         =   "&StopWatch..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPageSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsPrintPage 
         Caption         =   "Print a &Copy of this Page"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSetCurrentTime 
         Caption         =   "Set Current &Time Past the Hour (Step 1 in Determining  the Time the Current CD Ends && Air Time Remaining)   - "
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuToolsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsCD 
         Caption         =   "Include C&D Numbers with Entries"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuToolsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsMemos 
         Caption         =   "&Memos (access code required)..."
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Music Lineup"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Anno&unce Times"
      Begin VB.Menu mnuOptionsTime 
         Caption         =   "&Set Estimated Announce Times..."
      End
   End
   Begin VB.Menu mnuImport 
      Caption         =   "Import Lineup"
      Begin VB.Menu mnuCopyMusicLogList 
         Caption         =   "&Import Lineup from Music Planning Page    -"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuExport 
      Caption         =   "Export Lineup"
      Begin VB.Menu mnuToolsExportLineCopy 
         Caption         =   "Copy &SELECTED Line to Music Planning Page Lineup"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuToolsSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsExportLineups 
         Caption         =   "Copy &ENTIRE Lineup to Music Planning Page Lineup"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmTimeRemain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TIME REMAIN Sept 4, 2011
    Dim miTotal As String
    Dim sRemain As String
    Dim iiMinute1 As Integer

    Dim miCalSec As Integer
    Dim miCalSec1 As Integer 'Frame2 caption, estimated sec announce time
    Dim miCalMin As Integer
    Dim miCalMin1 As Integer 'Frame2 caption, estimated min announce time
    Dim cTotalRemain As Currency 'formerly pAdd declaration only. changed to be used in txtSpot announcement changes

    Dim miBackAnnc As Integer
    Dim miAnncTime As Integer
    Dim iAnnounce1 As Integer
    Dim iAnnounce2 As Currency
    Dim iAnnounce3 As Integer
    Dim iAnnounce4 As Currency
    Dim iAnnounce5 As Integer
    Dim miCloseOut As String
    Dim miBackTime As Currency
    
    Dim iManSec As Integer
    Dim iManMin As Integer
    
    Dim BackAnnc As String
    Dim iBackAnnc As String
    Dim iAnnc4 As String 'this series is used in saving annc times to file
    Dim iAnnc5 As String
    Dim iAnnc6 As String
    Dim iAnnc7 As String
    Dim iAnnc8 As String
    Dim iAnnc9 As String
    Dim iAnnc10 As String
    Dim iAnnc11 As String
 
    Dim iiBackAnnc As String
    Dim iiAnnc4 As String 'this series is used in computing annc time
    Dim iiAnnc5 As String
    Dim iiAnnc6 As String
    Dim iiAnnc7 As String
    Dim iiAnnc8 As String
    Dim iiAnnc9 As String
    Dim iiAnnc10 As String
    Dim iiAnnc11 As String
 
    Dim AnncControl4 As Integer
    Dim AnncControl5 As Integer
    Dim AnncControl6 As Integer
    Dim AnncControl7 As Integer
    Dim AnncControl8 As Integer
    Dim AnncControl9 As Integer
    Dim AnncControl10 As Integer
    Dim AnncControl11 As Integer
    Dim cMinAx As Currency
    
    Dim mRAnncSum As Integer
    Dim RecallControl As Integer
    Dim iAnncSum As Integer
    Dim mCurrentTime As Integer 'control used if current time before or after current hour
    Dim mcAnncTimez As Currency
    Dim miRestoreSave As Integer 'prevents saving restsored file when "added" --this is copied raw from planner
    Dim miTotalRemain As Integer
    Dim iRandomNumber As Integer 'random msgbox reminder
    Dim sSpotRem As String
    Dim iCurrentTime As Integer
    Dim mtSec As Integer
    Dim mtMin As Integer
    Dim miHourAdjust As Integer
    Dim miCurrentHour As Integer 'prevents TimeRemain double messages if hour between 57 & 59 min
    Dim iHourNow As Integer
    Dim vMinAdj As Integer
    Dim Composer4, Composer5, Composer6, Composer7, Composer8, Composer9, Composer10, Composer11 As String
    Dim Minute1, Minute2, Minute3, Minute4, Minute5, Minute6, Minute7, Minute8, Minute9, Minute10, Minute11, _
        Second1, Second2, Second3, Second4, Second5, Second6, Second7, Second8, Second9, Second10, Second11, _
        spotsS, anncSum As String
        
    Dim CD4, CD5, CD6, CD7, CD8, CD9, CD10, CD11 As String
        
    Dim iMinute4, iMinute5, iMinute6, iMinute7, iMinute8, iMinute9, iMinute10, iMinute11, _
        iSecond4, iSecond5, iSecond6, iSecond7, iSecond8, iSecond9, iSecond10, iSecond11 As String
'----test
    Dim iTMin4, iTMin5, iTMin6, iTMin7, iTMin8, iTMin9, iTMin10, iTMin11 As Integer
    Dim iTSec4, iTSec5, iTSec6, iTSec7, iTSec8, iTSec9, iTSec10, iTSec11 As Integer
    
    Dim mtxtComposer, mtxtMin, mtxtSec As String
    Dim cMusicMin As Integer
    Dim msgSelect As Integer
Option Explicit

Private Sub pAdd()

    'this is the "motor" that drives the add functions
    Dim cBlock As Currency
    Dim cRemain2 As Currency
    Dim cRemain3 As Currency
    Dim cCombined As Currency
    Dim aiMinute1 As Integer
    Dim aiMinute2 As Integer
    Dim aiMinute3 As Integer
    Dim aiMinute4 As Integer
    Dim aiMinute5 As Integer
    Dim aiMinute6 As Integer
    Dim aiMinute7 As Integer
    Dim aiMinute8 As Integer
    Dim aiMinute9 As Integer
    Dim aiMinute10 As Integer
    Dim aiMinute11 As Integer
    
    Dim cTotalMin As Currency
    Dim cTotalSec As Currency
    Dim aiSecond1 As Integer
    Dim aiSecond2 As Integer
    Dim aiSecond3 As Integer
    Dim aiSecond4 As Integer
    Dim aiSecond5 As Integer
    Dim aiSecond6 As Integer
    Dim aiSecond7 As Integer
    Dim aiSecond8 As Integer
    Dim aiSecond9 As Integer
    Dim aiSecond10 As Integer
    Dim aiSecond11 As Integer
    
    Dim cHours As Currency
    Dim cSecAdd As Integer
    Dim cSecAddS As Integer
    Dim cMinAdd As Integer
    Dim cMinAddS As Integer
    Dim cTotalMinA As Currency
    Dim cMCal2 As Currency
    Dim cMCal2S As Currency
    Dim cMinA As Currency
    Dim cSecA As Currency
    Dim cSecNA As Currency
    Dim cHCal1 As Currency
    Dim cHCal3 As Currency
    Dim cHCal4 As Currency
    Dim iCal1 As Currency 'Frame2 caption, estimated announce time
         
    '--------------

    'MINUTES, set "Min" values & formulate as numeric & Entry Error message
    
    If IsNumeric(txtMinute1) Or txtMinute1 = "" Then
        If mCurrentTime = 0 Then
             aiMinute1 = Val(txtMinute1) 'current show timing begins in the current hour
        Else
            aiMinute1 = (Val(txtMinute1) - 60) 'current show timing begins in previous hour
        End If
    Else
        MsgBox "The current time minutes is blank or includes a non-numeric character.", vbOKOnly, "Time Minutes Entry " & txtMinute1
        txtMinute1 = "0"
        txtMinute1.SetFocus
    End If

    If IsNumeric(txtMinute2) Or txtMinute2 = "" Then
        aiMinute2 = Val(txtMinute2)
    Else
        txtMinute2.ForeColor = vbRed
        MsgBox "The current CD minutes is blank or includes a non-numeric character.", vbOKOnly, "Current CD Minutes Entry " & txtMinute2
        txtMinute2 = "0"
        txtMinute2.ForeColor = &H80000008
        txtMinute2.SetFocus
    End If

    If IsNumeric(txtMinute3) Or txtMinute3 = "" Then
        aiMinute3 = Val(txtMinute3)
    Else
        txtMinute3.ForeColor = vbRed
        MsgBox "The estimated announce time is blank or includes a non-numeric character.", vbOKOnly, "Estimated Announce Time Minutes Entry " & txtMinute3
        txtMinute3 = "0"
        txtMinute3.ForeColor = &H80000008
        txtMinute3.SetFocus
    End If

    If IsNumeric(txtMinute4) Or txtMinute4 = "" Then
         aiMinute4 = Val(txtMinute4)
    Else
        txtMinute4.ForeColor = vbRed
        MsgBox "Music lineup minute 1, enter playing time up to a maximum of 99 minutes.", vbOKOnly, "Non-Numeric Entry"
        txtMinute4 = ""
        txtMinute4.ForeColor = &H80000008 'black
        txtMinute4.SetFocus
    End If

    If IsNumeric(txtMinute5) Or txtMinute5 = "" Then
        aiMinute5 = Val(txtMinute5)
    Else
        txtMinute5.ForeColor = vbRed
        MsgBox "Music lineup minute 2, enter playing time up to a maximum of 99 minutes.", vbOKOnly, "Non-Numeric Entry"
        txtMinute5 = ""
        txtMinute5.ForeColor = &H80000008 'black
        txtMinute5.SetFocus
    End If

    If IsNumeric(txtMinute6) Or txtMinute6 = "" Then
        aiMinute6 = Val(txtMinute6)
    Else
        txtMinute6.ForeColor = vbRed
         MsgBox "Music lineup minute 3, enter playing time up to a maximum of 99 minutes.", vbOKOnly, "Non-Numeric Entry"
         txtMinute6 = ""
         txtMinute6.ForeColor = &H80000008 'black
         txtMinute6.SetFocus
    End If

    If IsNumeric(txtMinute7) Or txtMinute7 = "" Then
        aiMinute7 = Val(txtMinute7)
    Else
        txtMinute7.ForeColor = vbRed
         MsgBox "Music lineup minute 4, enter playing time up to a maximum of 99 minutes.", vbOKOnly, "Non-Numeric Entry"
         txtMinute7 = ""
         txtMinute7.ForeColor = &H80000008 'black
         txtMinute7.SetFocus
    End If

    If IsNumeric(txtMinute8) Or txtMinute8 = "" Then
        aiMinute8 = Val(txtMinute8)
    Else
        txtMinute8.ForeColor = vbRed
         MsgBox "Music lineup minute 5, enter playing time up to a maximum of 99 minutes.", vbOKOnly, "Non-Numeric Entry"
         txtMinute8 = ""
         txtMinute8.ForeColor = &H80000008 'black
         txtMinute8.SetFocus
    End If

    If IsNumeric(txtMinute9) Or txtMinute9 = "" Then
        aiMinute9 = Val(txtMinute9)
    Else
        txtMinute9.ForeColor = vbRed
         MsgBox "Music lineup minute 6, enter playing time up to a maximum of 99 minutes.", vbOKOnly, "Non-Numeric Entry"
         txtMinute9 = ""
         txtMinute9.ForeColor = &H80000008 'black
         txtMinute9.SetFocus
    End If

    If IsNumeric(txtMinute10) Or txtMinute10 = "" Then
        aiMinute10 = Val(txtMinute10)
    Else
         txtMinute10.ForeColor = vbRed
         MsgBox "Music lineup minute 7, enter playing time up to a maximum of 99 minutes.", vbOKOnly, "Non-Numeric Entry"
         txtMinute10 = ""
         txtMinute10.ForeColor = &H80000008 'black
         txtMinute10.SetFocus
    End If

    If IsNumeric(txtMinute11) Or txtMinute11 = "" Then 'Or txtMinute11 = "-" Then
        aiMinute11 = Val(txtMinute11)
    Else
        txtMinute11.ForeColor = vbRed
        MsgBox "Music lineup minute 8, enter playing time up to a maximum of 99 minutes.", vbOKOnly, "Non-Numeric Entry"
        txtMinute11 = ""
        txtMinute11.ForeColor = &H80000008 'black
        txtMinute11.SetFocus
    End If

'-----End MINUTES values & validation---Begin SECONDS, values and validation---

'-------1
    If IsNumeric(txtSecond1) Or txtSecond1 = "" Then
        aiSecond1 = Val(txtSecond1)
    Else
         MsgBox "The current time seconds entry of  [ " & txtSecond1 & " ]  is blank or includes a non-numeric character.", vbOKOnly, "Current Time Seconds Entry Non-Numeric:  " & txtSecond1
         txtSecond1 = "0"
         txtSecond1.SetFocus
    End If
'test removal, keyword TimeCorrectionTest
    If Val(txtSecond1) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Current Time Seconds, Entry Error"
        txtSecond1 = ""
        txtSecond1.SetFocus
        Exit Sub
    End If
'-------2
    If IsNumeric(txtSecond2) Or txtSecond2 = "" Then
         aiSecond2 = Val(txtSecond2)
    Else
        txtSecond2.ForeColor = vbRed
        
        'txtMinute2 = ""
         MsgBox "The current CD seconds is blank or includes a non-numeric character.", vbOKOnly, "Current CD Seconds Entry " & txtSecond2
        ' MsgBox "The current CD seconds entry of  [ " & txtSecond2 & " ]  is or includes a non-numeric character.", vbOKOnly, "Current CD Seconds Entry Non-Numeric:  " & txtSecond2
         txtSecond2 = "0"
         txtSecond2.ForeColor = &H80000008 'black
         txtSecond2.SetFocus
    End If

    If Val(txtSecond2) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Current CD Seconds, Entry Error"
        txtSecond2 = ""
        txtSecond2.SetFocus
        Exit Sub
    End If
'-------3
    If IsNumeric(txtSecond3) Or txtSecond3 = "" Then
         aiSecond3 = Val(txtSecond3)
    Else
        txtSecond3.ForeColor = vbRed
        MsgBox "The estimated announce time seconds is blank or includes a non-numeric character.", vbOKOnly, "Estimated Announce Time Seconds " & txtSecond3
         txtSecond3 = "0"
         txtSecond3.ForeColor = &H80000008 'black
         txtSecond3.SetFocus
    End If

    If Val(txtSecond3) > 59 Then
        MsgBox "The estimated announce time seconds entry may not exceed 59 seconds", 0, "Estimated Announce Time Seconds Entry Error"
        txtSecond3 = ""
        txtSecond3.SetFocus
        Exit Sub
    End If
'-------4
    If IsNumeric(txtSecond4) Or txtSecond4 = "" Then
        aiSecond4 = Val(txtSecond4)
    Else
        txtSecond4.ForeColor = vbRed
        MsgBox "Enter number, 59 seconds or less", vbOKOnly, "Music Lineup Seconds 1, Non-Numeric Entry"
        txtSecond4 = ""
        txtSecond4.ForeColor = &H80000008 'black
        txtSecond4.SetFocus
    End If

    If Val(txtSecond4) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Music Lineup Seconds 1, Entry Error"
        txtSecond4 = ""
        txtSecond4.SetFocus
        Exit Sub
    End If
'-------5
    If IsNumeric(txtSecond5) Or txtSecond5 = "" Then
        aiSecond5 = Val(txtSecond5)
    Else
        txtSecond5.ForeColor = vbRed
        MsgBox "Enter number, 59 seconds or less", vbOKOnly, "Music Lineup Seconds 2, Non-Numeric Entry"
        txtSecond5 = ""
        txtSecond5.ForeColor = &H80000008 'black
        txtSecond5.SetFocus
    End If

    If Val(txtSecond5) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Music Lineup Seconds 2, Error"
        txtSecond5 = ""
        txtSecond5.SetFocus
        Exit Sub
    End If
'-------6
    If IsNumeric(txtSecond6) Or txtSecond6 = "" Then
        aiSecond6 = Val(txtSecond6)
    Else
        txtSecond6.ForeColor = vbRed
        MsgBox "Enter number, 59 seconds or less", vbOKOnly, "Music Lineup Seconds 3, Non-Numeric Entry"
        txtSecond6 = ""
        txtSecond6.ForeColor = &H80000008 'black
        txtSecond6.SetFocus
    End If

    If Val(txtSecond6) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Music Lineup Seconds 3, Error"
        txtSecond6 = ""
        txtSecond6.SetFocus
        Exit Sub
    End If
'-------7
      If IsNumeric(txtSecond7) Or txtSecond7 = "" Then
        aiSecond7 = Val(txtSecond7)
    Else
        txtSecond7.ForeColor = vbRed
        MsgBox "Enter number, 59 seconds or less", vbOKOnly, "Music Lineup Seconds 4, Non-Numeric Entry"
        txtSecond7 = ""
        txtSecond7.ForeColor = &H80000008 'black
        txtSecond7.SetFocus
    End If

    If Val(txtSecond7) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Music Lineup Seconds 4, Error"
        txtSecond7 = ""
        txtSecond7.SetFocus
        Exit Sub
    End If
 '-------8
    If IsNumeric(txtSecond8) Or txtSecond8 = "" Then
        aiSecond8 = Val(txtSecond8)
    Else
        txtSecond8.ForeColor = vbRed
        MsgBox "Enter number, 59 seconds or less", vbOKOnly, "Music Lineup Seconds 5, Non-Numeric Entry"
        txtSecond8 = ""
        txtSecond8.ForeColor = &H80000008 'black
        txtSecond8.SetFocus
    End If

    If Val(txtSecond8) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Music Lineup Seconds 5, Error"
        txtSecond8 = ""
        txtSecond8.SetFocus
        Exit Sub
    End If
'-------9
    If IsNumeric(txtSecond9) Or txtSecond9 = "" Then
        aiSecond9 = Val(txtSecond9)
    Else
        txtSecond9.ForeColor = vbRed
        MsgBox "Enter number, 59 seconds or less", vbOKOnly, "Music Lineup Seconds 6, Non-Numeric Entry"
        txtSecond9 = ""
        txtSecond9.ForeColor = &H80000008 'black
        txtSecond9.SetFocus
    End If

    If Val(txtSecond9) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Music Lineup Seconds 6, Error"
        txtSecond9 = ""
        txtSecond9.SetFocus
        Exit Sub
    End If
'------10
    If IsNumeric(txtSecond10) Or txtSecond10 = "" Then
        aiSecond10 = Val(txtSecond10)
    Else
        txtSecond10.ForeColor = vbRed
        MsgBox "Enter number, 59 seconds or less", vbOKOnly, "Music Lineup Seconds 7, Non-Numeric Entry"
        txtSecond10 = ""
        txtSecond10.ForeColor = &H80000008 'black
        txtSecond10.SetFocus
    End If

    If Val(txtSecond10) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Music Lineup Seconds 7, Error"
        txtSecond10 = ""
        txtSecond10.SetFocus
        Exit Sub
    End If

'-------11
    If IsNumeric(txtSecond11) Or txtSecond11 = "" Then 'Or txtSecond11 = "-" Then
        aiSecond11 = Val(txtSecond11)
    Else
        txtSecond11.ForeColor = vbRed
        MsgBox "Enter number, 59 seconds or less", vbOKOnly, "Music Lineup Seconds 8, Non-Numeric Entry"
        txtSecond11 = ""
        txtSecond11.ForeColor = &H80000008 'black
        txtSecond11.SetFocus
    End If

    If Val(txtSecond11) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Music Lineup Seconds 8, Error"
        txtSecond11 = ""
        txtSecond11.SetFocus
        Exit Sub
    End If
'================End SECONDS values and validation--------

    Dim iSpotRS As Integer

    If txtSpotsS <> "" And txtSpotsS <> "0" Then
        iSpotRS = (Val(txtSpotsS) * Val(txtSpotLength))
    ElseIf txtSpotsS = "" Then
        iSpotRS = 0
    End If

'===========Total Time, miTotal

   If chkAnnounce.Value = 0 Then

        If txtMinute3 = "" And txtSecond3 = "" Then
            iAnncSum = Val(iiAnnc4) + Val(iiAnnc5) + Val(iiAnnc6) + Val(iiAnnc7) + Val(iiAnnc8) + Val(iiAnnc9) + Val(iiAnnc10) + Val(iiAnnc11) + iSpotRS + Val(txtBackAnnc)
        Else
            iAnncSum = aiSecond3 + (aiMinute3 * 60) + iSpotRS
        End If

    ElseIf chkAnnounce.Value = 1 Then
        iAnncSum = (Val(txtMinute3) * 60) + Val(txtSecond3) + iSpotRS
    End If
    
    
'============

    cMinAdd = aiMinute1 + aiMinute2 + aiMinute4 + aiMinute5 + aiMinute6 + aiMinute7 + aiMinute8 + aiMinute9 + aiMinute10 + aiMinute11
    cSecAdd = iAnncSum + aiSecond1 + aiSecond2 + aiSecond4 + aiSecond5 + aiSecond6 + aiSecond7 + aiSecond8 + aiSecond9 + aiSecond10 + aiSecond11

'=======================
    cTotalMinA = cMinAdd + cSecAdd / 60

    cMCal2 = cTotalMinA - Int(cTotalMinA) 'sec (fraction)
    cMinA = cTotalMinA - cMCal2 'whole min (removes fraction)
    cSecA = cMCal2 * 60 'sec, multiplies fraction by 60 producing seconds

'---ADDING minute & second entries, used in total program time printout

    Dim cMinAddz As Integer
    Dim cSecAddz As Integer
    Dim cTotalMinAz As Currency
    Dim cMCal2z As Currency
    Dim cMinAz As Currency
    Dim cSecAz As Currency

 '---(1) Used for printing program time on Planner page printout

    cMinAddz = aiMinute3 + aiMinute4 + aiMinute5 + aiMinute6 + aiMinute7 + aiMinute8 + aiMinute9 + aiMinute10 + aiMinute11
    cSecAddz = aiSecond3 + aiSecond4 + aiSecond5 + aiSecond6 + aiSecond7 + aiSecond8 + aiSecond9 + aiSecond10 + aiSecond11

    cTotalMinAz = cMinAddz + cSecAddz / 60

    cMCal2z = cTotalMinAz - Int(cTotalMinAz) 'sec (fraction)
    gcMinA = cTotalMinAz - cMCal2z 'whole min (removes fraction)
    gcSecA = cMCal2z * 60 'sec, multiplies fraction by 60 producing seconds

'---(2) Used for lblTotalS and Panels(2) on Planner page

    cMinAddS = aiMinute1 + aiMinute2 + aiMinute4 + aiMinute5 + aiMinute6 + aiMinute7 + aiMinute8 + aiMinute9 + aiMinute10 + aiMinute11
    cSecAddS = aiSecond1 + aiSecond2 + aiSecond4 + aiSecond5 + aiSecond6 + aiSecond7 + aiSecond8 + aiSecond9 + aiSecond10 + aiSecond11

    gcTotalMinNA = cMinAddS + cSecAddS / 60 'TOTAL, adds min & sec/60 to = min & decimal min.

    cMCal2S = gcTotalMinNA - Int(gcTotalMinNA) 'sec (fraction)
    gcMinNA = gcTotalMinNA - cMCal2S 'whole min (removes fraction)
    cSecNA = cMCal2S * 60 'sec, multiplies fraction by 60 producing seconds

'---(3) Used for Panels 3 TimeRemain page

    Dim cMinAddT As Integer
    Dim cSecAddT As Integer
    Dim cMCal2T As Currency
    Dim cMinNT As Currency
    Dim cSecNT As Currency

    cMinAddT = aiMinute4 + aiMinute5 + aiMinute6 + aiMinute7 + aiMinute8 + aiMinute9 + aiMinute10 + aiMinute11
    cSecAddT = aiSecond4 + aiSecond5 + aiSecond6 + aiSecond7 + aiSecond8 + aiSecond9 + aiSecond10 + aiSecond11

    gcTotalMinNT = cMinAddT + cSecAddT / 60

    cMCal2T = gcTotalMinNT - Int(gcTotalMinNT) 'sec (fraction)
    cMinNT = gcTotalMinNT - cMCal2T 'whole min (removes fraction)
    cSecNT = cMCal2T * 60 'sec, multiplies fraction by 60 producing seconds

'========Planner Page program line time computations

    Dim cProgramMinAdds As Integer
    Dim cProgramSecAdds As Integer

    Dim cProgramSecAdd As Currency
    Dim cMinPAdd As Currency
    Dim cAnncTimez As Currency

     cAnncTimez = iAnncSum

    cProgramMinAdds = aiMinute4 + aiMinute5 + aiMinute6 + aiMinute7 + aiMinute8 + aiMinute9 + aiMinute10 + aiMinute11
    cProgramSecAdds = cAnncTimez + aiSecond4 + aiSecond5 + aiSecond6 + aiSecond7 + aiSecond8 + aiSecond9 + aiSecond10 + aiSecond11

    cMinPAdd = cProgramMinAdds + cProgramSecAdds / 60 'converts seconds into decimal minutes and add to minutes
    '--------
    cProgramSecAdd = cMinPAdd - Int(cMinPAdd) 'sec (fraction)
    gcProgramMinAdd = cMinPAdd - cProgramSecAdd
    gcSecPAdd = cProgramSecAdd * 60
    '------------

    If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
    And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then

'--------Begin Planner Page program lines (sPlannerProgramTime reads out on planner page)

    Dim sSpotNum As String 'converts digit to word in Planner program message

        If Val(frmPlanner!txtSpots) <> 0 Then
            Select Case Val(frmPlanner!txtSpots)
                Case Is < 1
                    sSpotNum = Val(frmPlanner!txtSpots) & " "
                Case 1
                    sSpotNum = "one"
                Case 2
                    sSpotNum = "two"
                Case 3
                    sSpotNum = "three"
                Case 4
                    sSpotNum = "four"
                Case 5
                    sSpotNum = "five"
                Case Else
                    sSpotNum = Val(frmPlanner!txtSpots) & " "
            End Select
        Else
            sSpotNum = ""
        End If

        If giPlannerTxtSpot = 1 Then

            If Val(frmPlanner!txtSpots) <= 1 Then 'single spot ••

                If txtMinute1 = "" And txtSecond1 = "" Then
                    sPlannerProgramTime = "music + " & miCalMin & " min " & Format$(miCalSec, "#0") & " sec announce (includes " & sSpotNum & " " & Val(txtSpotLength) & "-sec" & _
                    " spot) =" & Format$(gcProgramMinAdd, " 0") & " min " & Format$(gcSecPAdd, "#0") & " sec program time"

                 ElseIf txtMinute1 <> "" Or txtSecond1 <> "" Then '* is a reminder computer time adjustment in effect
                    sPlannerProgramTime = "* music + " & miCalMin & " min " & Format$(miCalSec, "#0") & " sec announce (includes " & sSpotNum & " " & Val(txtSpotLength) & "-sec" & _
                    " spot) =" & Format$(gcProgramMinAdd, " 0") & " min " & Format$(gcSecPAdd, "#0") & " sec program time *"
                 End If

            ElseIf Val(frmPlanner!txtSpots) > 1 Then 'multiple spots

                If txtMinute1 = "" And txtSecond1 = "" Then
                    sPlannerProgramTime = "music + " & miCalMin & " min " & Format$(miCalSec, "#0") & " sec announce (includes " & sSpotNum & " " & Val(txtSpotLength) & "-sec" & _
                     " spots) =" & Format$(gcProgramMinAdd, " 0") & " min " & Format$(gcSecPAdd, "#0") & " sec program time"

                ElseIf txtMinute1 <> "" Or txtSecond1 <> "" Then
                    sPlannerProgramTime = "* music + " & miCalMin & " min " & Format$(miCalSec, "#0") & " sec announce (includes " & sSpotNum & " " & Val(txtSpotLength) & "-sec" & _
                    " spots) =" & Format$(gcProgramMinAdd, " 0") & " min " & Format$(gcSecPAdd, "#0") & " sec program time *"
                End If

            End If

       ElseIf giPlannerTxtSpot = 0 Then 'no spots

            If txtMinute1 = "" And txtSecond1 = "" Then

                sPlannerProgramTime = "music + " & miCalMin & " min " & Format$(miCalSec, "#0") & " sec estimated announce time =" & Format$(gcProgramMinAdd, " 0") & " min " & Format$(gcSecPAdd, "#0") & " sec program time"

            ElseIf txtMinute1 <> "" Or txtSecond1 <> "" Then
                sPlannerProgramTime = "* music + " & miCalMin & " min " & Format$(miCalSec, "#0") & " sec estimated announce time =" & Format$(gcProgramMinAdd, " 0") & " min " & Format$(gcSecPAdd, "#0") & " sec program time *"
            End If

        End If
  '----------------

    Else
        frmPlanner!lblProgramTime = ""
    End If
'============ End Planner Page program time line

'----------------- Panel 2 display

    Dim cTotalMinNA1 As Currency
    Dim cMCal2S1 As Currency
    Dim cMinNA1 As Currency
    Dim cSecNA1 As Currency

    If txtBlock <> "" And ((cMinAddS + cSecAddS / 60) <= Val(txtBlock)) Then
        cTotalMinNA1 = Val(txtBlock) - (cMinAddS + cSecAddS / 60) 'Block time less TOTAL (adds min & sec/60 to = min & decimal min)
        cMCal2S1 = cTotalMinNA1 - Int(cTotalMinNA1) 'sec (fraction)
        cMinNA1 = cTotalMinNA1 - cMCal2S1 'Minutes, whole min (removes fraction)
        cSecNA1 = cMCal2S1 * 60 'Seconds, multiplies fraction by 60 producing seconds

        If lblCurrentTime.ForeColor <> vbRed Then
        
            If txtMinute2 <> "" Or txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Then

                StaRemain.Panels(2) = ""
                
                StaRemain.Panels(2) = gcMinNA & " min " & Format(cSecNA, "00") & " sec programmed. " _
               & "Excluding announce time, " & cMinNA1 & " min " & Format(cSecNA1, "#0") & " sec remain of " & txtBlock & " min planned"
    
                StaRemain.Panels(3) = "Music Lineup: " & cMinNT & ":" & Format(cSecNT, "00")
            Else
                StaRemain.Panels(2) = ""
            End If
        Else
            StaRemain.Panels(2) = ""
        End If
    Else
        cTotalMinNA1 = (Val(txtBlock) - (cMinAddS + cSecAddS / 60)) * (-1) 'Block time less TOTAL (adds min & sec/60 to = min & decimal min)
        cMCal2S1 = cTotalMinNA1 - Int(cTotalMinNA1) 'sec (fraction)
        cMinNA1 = cTotalMinNA1 - cMCal2S1 'Minutes, whole min (removes fraction)
        cSecNA1 = cMCal2S1 * 60 'Seconds, multiplies fraction by 60 producing seconds

        If lblCurrentTime.ForeColor <> vbRed Then
        
            If txtMinute2 <> "" Or txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Then
            
                StaRemain.Panels(2) = ""
                StaRemain.Panels(2) = "Excluding announce time, " & gcMinNA & " min " & Format(cSecNA, "00") & " sec programmed, " _
                & "Time overrun: " & cMinNA1 & " min " & Format(cSecNA1, "#0") & " sec"
    
                StaRemain.Panels(3) = "Music Lineup: " & cMinNT & ":" & Format(cSecNT, "00")
                
            Else
                StaRemain.Panels(2) = ""
            End If
        Else
            StaRemain.Panels(2) = ""
        End If
    End If

''--------------Program Information (lblProgramInfo) text

'----------determining TIME REMAINING by subtracting total time from block time
    cBlock = Val(txtBlock) 'planned time
    cTotalRemain = cBlock - cTotalMinA 'time remaining, assigned time subtracted from planned time
    cRemain2 = 60 - cTotalMinA '60 minus min & decimal min.
    cRemain3 = 30 - cTotalMinA '30 minus min & decimal min.
'------------

    Dim cRemainFraction As Currency
    Dim cRemainMin As Currency
    Dim cRemainSec As Currency
    Dim cMusicRemain As Currency
    Dim sRemain30 As String

    Dim cMinConvert2 As Currency
    Dim cMusicMin2 As Currency
    Dim cMusicSec2 As Currency
    Dim cConvertRemain As Currency

    Dim cSecExtract As Currency
    Dim cMusicSec As Integer

    Dim test21 As Currency
    Dim test22 As Currency
    Dim test23 As Currency
    Dim test24 As Currency
    Dim test25 As Currency
    Dim test26 As Currency
    Dim testRemain2 As Currency

    Dim test31 As Currency
    Dim test32 As Currency
    Dim test33 As Currency
    Dim test34 As Currency
    Dim test35 As Currency
    Dim test36 As Currency
    Dim testRemain3 As Currency
  '----
  Dim sRemain60 As String

  'values for actual time remaining (sRemain60) display
    If cTotalRemain > 0 Then 'convert to min & seconds
        cRemainFraction = cTotalRemain - Int(cTotalRemain) 'removes integer,
                                          'remaining time fraction remains
        cRemainMin = cTotalRemain - cRemainFraction 'minutes
        cRemainSec = cRemainFraction * 60 'seconds
        sRemain60 = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec"

    Else 'converts from a negative time-remaining display
        cConvertRemain = cTotalRemain * (-1)
        cMinConvert2 = cConvertRemain - Int(cConvertRemain)
        cMusicMin2 = cConvertRemain - cMinConvert2 'if 60 minutes exceeded, shows time over 60
        cMusicSec2 = cMinConvert2 * 60
        sRemain60 = Format$(cMusicMin2, "0") & " min " & Format$(cMusicSec2, "#0") & " sec"
     End If

'----------------values for Displaying PROGRAM TIME REMAINING (lblProgramRemain)-----------------------------
    'cTotalRemain is uncorrected (for intro, closeout & spots) remaining seconds
    'cMusicRemain is seconds remaining for music (before separation into min & sec
    'if Announce Time is manually entered, there is no Intro/BachkAnnounce correction

    If txtMinute3 = "" And txtSecond3 = "" Then 'announce adjustments

        If txtMinute1 <> "" And txtMinute2 <> "" And lblTotalS.Visible = False Then
            cMusicRemain = cTotalRemain - (Val(txtIntro) / 60) - (iSpotRS / 60) - (Val(txtCloseOut) / 60)
        Else
            cMusicRemain = cTotalRemain - (Val(txtIntro) / 60) - (Val(txtCloseOut) / 60)
        End If

    ElseIf txtMinute3 <> "" Or txtSecond3 <> "" Or chkAnnounce.Value = 1 Then
            cMusicRemain = cTotalRemain
    End If

    '---------------Conversion music remain time into minutes & seconds (cMusicMin & cMusicSec)
    cSecExtract = cMusicRemain - Int(cMusicRemain) 'extracts seconds fraction from min & fraction
    cMusicMin = cMusicRemain - cSecExtract 'minutes
    cMusicSec = cSecExtract * 60 'secondsformat

    '-------Time Remain statement-------------
    miTotalRemain = cTotalRemain

    If cTotalRemain = 0 Then

        If txtSpotsS = "" Then
            sRemain = "No Time Remaining"
            gsRemain1 = "No Time Remaining"
        Else
            sRemain = "No Time Remaining " & " (" & sSpotRem & ")"
            gsRemain1 = "No Time Remaining " & " (" & sSpotRem & ")"
        End If

        Label20.Visible = False
    ElseIf cTotalRemain < 0 And cTotalRemain > -60 Then

        If Val(txtBlock) > 0.99 Then
            sRemain = txtBlock & " Min Planned Program Time Exceeded"
            gsRemain1 = txtBlock & " Min Planned Program Time Exceeded"

        ElseIf Val(txtBlock) > 0 And Val(txtBlock) <= 0.99 Then

            sRemain = txtBlock & " is not a valid planned time entry"
            gsRemain1 = txtBlock & " is not a valid planned time entry"

        Else
            sRemain = "Planned Program Time Missing"
            gsRemain1 = "Planned Program Time Missing"
        End If

        Label20.Visible = False
    ElseIf cTotalRemain <= -60 Then
        lblProgramRemain.ForeColor = vbRed
        sRemain = "Time Runs Beyond the Next Hour"
        gsRemain1 = "Time Runs Beyond the Next Hour"
        Label20.Visible = False
    Else

'----------------------------------------------

    Dim iSpots As String

    If txtSpotsS <> "" Then
        iSpots = Val(txtSpotsS)
    Else
        iSpots = ""
    End If

    If iSpots <> "" And iSpots <> "0" Then
        Select Case iSpots

            Case Is < 1
                iSpots = "(" & ((txtSpotsS) * txtSpotLengthSetting) & " sec spot)"
            Case 1
                iSpots = "one spot"
            Case 2
                iSpots = "two spots"
            Case 3
                iSpots = "three spots"
            Case 4
                iSpots = "four spots"
            Case 5
                iSpots = "five spots" '
            Case Else
                iSpots = (txtSpotsS) & " spots"
        End Select
    Else
        iSpots = ""
    End If
    sSpotRem = iSpots
'=======================

        '(1) time remaining is 3:15 minute or greater

        If cTotalRemain >= 3.16 Then

            If sSpotRem = "" Then 'no spots
            
                If Label20.Visible = True Then 'label is 30 sec remain for closeout & ID
                    sRemain = Format$(cMusicMin, "0") & " min " & Format$(cMusicSec, "#0") & " secs are available for music •"
                    gsRemain1 = Format$(cMusicMin, "0") & " min " & Format$(cMusicSec, "#0") & " secs are available for music"
                
                Else
                    sRemain = Format$(cMusicMin, "0") & " min " & Format$(cMusicSec, "#0") & " secs are available for music"
                    gsRemain1 = Format$(cMusicMin, "0") & " min " & Format$(cMusicSec, "#0") & " secs are available for music"
                End If
                
            Else 'spots
                sRemain = Format$(cMusicMin, "0") & " min " & Format$(cMusicSec, "#0") & " secs are available for music " & "(" & sSpotRem & " unplayed)"
                gsRemain1 = Format$(cMusicMin, "0") & " min " & Format$(cMusicSec, "#0") & " secs are available for music " & "(" & sSpotRem & " unplayed)"

            End If
            lblProgramRemain.ForeColor = &H80& 'light rust

        '(2) time remaining is less than 3:15 min
        ElseIf cTotalRemain < 3.16 And cTotalRemain > 0 Then

            If cTotalRemain < 3.16 And cTotalRemain > 2 Then
                lblProgramRemain.ForeColor = &HC00000 'blue '&H8000& 'green
            ElseIf cTotalRemain <= 2 And cTotalRemain > 0 Then
                lblProgramRemain.ForeColor = &HC00000      'Blue
            End If

            If txtSpotsS = "" Then

                If cTotalRemain < 3.16 And cTotalRemain > 1.5 Then
                
'                    If txtMinute3 = "" And txtSecond3 = "" Then
'                        sRemain = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain for filler, closeout && ID"
'                        gsRemain1 = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain for filler, closeout & ID"
'
'                    ElseIf txtMinute3 <> "" Or txtSecond3 <> "" Then
                        
                        sRemain = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain for filler, closeout && ID"
                        gsRemain1 = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain for filler, closeout && ID"
                 '   End If
                                
                ElseIf cTotalRemain <= 1.5 And cTotalRemain > 0.5 Then
                
'                    If txtMinute3 = "" And txtSecond3 = "" Then
'                        sRemain = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain for closeout && ID"
'                        gsRemain1 = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain for closeout & ID"
'
'                    ElseIf txtMinute3 <> "" Or txtSecond3 <> "" Then
                        
                        sRemain = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain for closeout && ID"
                        gsRemain1 = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain for closeout && ID"
                   ' End If

                ElseIf cTotalRemain <= 0.5 And cTotalRemain > 0 Then
                
                   sRemain = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain"
                   gsRemain1 = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain"
                    
                End If

           ElseIf txtSpotsS <> "" And txtSpotsS <> "0" Then

                If txtMinute3 = "" And txtSecond3 = "" Then

                    sRemain = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec for program, " & iSpotRS & " sec for " & sSpotRem
                              
                If Val(txtSpotsS) = 1 Then
                    gsRemain1 = iSpotRS & " sec for " & (txtSpotsS) & " spot + " & Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec for closeout"

                ElseIf Val(txtSpotsS) > 1 Then
                    gsRemain1 = iSpotRS & " sec for " & (txtSpotsS) & " spots + " & Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec for closeout"
                End If

            Else
                If cMusicRemain > 0 Then

                    If txtSpotsS <> "" And txtSpotsS <> "0" Then
                        sRemain = iSpotRS & " sec for " & sSpotRem & " + " & Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec for closeout"
                    Else
                        sRemain = Format$(cMusicMin, "0") & " min " & Format$(cMusicSec, "#0") & " sec closeout"
                    End If

                    ElseIf cMusicRemain <= 0 And cMusicRemain > -1 Then
                        sRemain = "No Time Remaining " & " (" & sSpotRem & ")"
                    ElseIf cMusicRemain <= -0.5 Then
                        sRemain = "Program time exceeded by 1/2 minute or more"
                    End If
            End If

        Else
            sRemain = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain"
            gsRemain1 = Format$(cRemainMin, "0") & " min " & Format$(cRemainSec, "#0") & " sec remain"
        End If

            Label20.Visible = False

        End If
    End If
'------------

    iCal1 = (miAnncTime + Val(txtIntro)) / 60
    miCalSec1 = (iCal1 - Int(iCal1)) * 60 'miCalSec, announce time seconds
    miCalMin1 = Int(iCal1) 'miCalMin1, announce time minutes

    If lblAnnounceTime.Visible = True And lblProgramRemain.ForeColor = &H80& And txtMinute3 = "" And txtSecond3 = "" Then 'light rust
     
        Frame2.Caption = " • Including a PROJECTED total announce time of " & miCalMin1 & " min " & Format$(miCalSec1, "#0") & " sec"
        Frame4.Caption = " • Projected program announce time of " & miCalMin1 & " min " & Format$(miCalSec1, "#0") & " sec excluding closeout && ID"
    ElseIf lblAnnounceTime.Visible = True And lblProgramRemain.ForeColor = &H80& And (txtMinute3 <> "" Or txtSecond3 <> "") Then
       
    '-----------------------------spells out txtSpots number
        Dim sSpotsS As String

        If txtSpotsS <> "" And txtSpotsS <> "0" Then
            Select Case txtSpotsS

                Case Is < 1
                    sSpotsS = Val(txtSpotsS)
                Case 1
                    sSpotsS = "one"
                Case 2
                    sSpotsS = "two"
                Case 3
                    sSpotsS = "three"
                Case 4
                    sSpotsS = "four"
                Case 5
                    sSpotsS = "five"
                Case Else
                    sSpotsS = txtSpotsS
            End Select
        Else
            sSpotsS = ""
        End If
    '------------------

        If txtSpotsS.Visible = True And txtSpotsS <> "" And txtSpotsS <> "0" Then
            Frame2.Caption = " • Including " & sSpotsS & " unplayed " & txtSpotLength & "-sec spot announcement(s)"
            Frame4.Caption = ""
        Else
           ' Frame2.ForeColor = &H80&
            Frame2.Caption = " • With no allowance for additional announce time"
            Frame4.Caption = ""
        End If

    Else
       ' Frame2.ForeColor = &H80&
        Frame2.Caption = "Time Available"
        Frame4.Caption = ""
    End If
'----------------Remain of 30 minutes (lblRemain30)--------------------------------------

    Dim cRemain30 As Currency
    Dim cMusicRemain30 As Currency
    Dim cMusicMin30 As Currency
    Dim cMusicSec30 As Currency
    Dim cSecExtract30 As Currency

    Dim cRemain31 As Currency
    Dim cMusicRemain31 As Currency
    Dim cMusicMin31 As Currency
    Dim cMusicSec31 As Currency
    Dim cSecExtract31 As Currency

    cMusicRemain30 = (cTotalRemain - 30) - (Val(txtIntro) / 60) 'does not include show closeout
    cMusicRemain31 = (cTotalRemain - 30)

'---------------Conversion music remain time into minutes & seconds (cMusicMin & cMusicSec)
    cSecExtract30 = cMusicRemain30 - Int(cMusicRemain30) 'extracts seconds fraction from min & fraction
    cMusicMin30 = cMusicRemain30 - cSecExtract30 'minutes
    cMusicSec30 = cSecExtract30 * 60 'secondsformat
'-------------
    cSecExtract31 = cMusicRemain31 - Int(cMusicRemain31) 'extracts seconds fraction from min & fraction
    cMusicMin31 = cMusicRemain31 - cSecExtract31 'minutes
    cMusicSec31 = cSecExtract31 * 60 'secondsformat
'-------------

    cRemain30 = (cMusicMin30 - 30) 'this position follows the computation of cMusicMin

'------------DISPLAY of the above values

    lblProgramRemain.Caption = sRemain
   
    test21 = cRemain2 - Int(cRemain2)

    If cRemain2 >= 0 Then
        test22 = cRemain2 - test21
        test23 = test21 * 60
    Else
        testRemain2 = cRemain2 * (-1)
        test24 = testRemain2 - Int(testRemain2)
        test25 = testRemain2 - test24
        test26 = test24 * 60
    End If
'-----

    If cRemain2 > 25 Then
    
        test31 = cRemain3 - Int(cRemain3)
    
         If cRemain3 >= 0 Then
            test32 = cRemain3 - test31
            test33 = test31 * 60
            lblRemain30.Caption = "  " & Format$(test32, "0") & " min " & Format$(test33, "#0") & " sec remain of the first 30 minutes  "
            lblRemain30.BackColor = &H80000005  'white
          
        Else
            testRemain3 = cRemain3 * (-1)
            test34 = testRemain3 - Int(testRemain3)
            test35 = testRemain3 - test34
            test36 = test34 * 60
            lblRemain30.Caption = "  " & Format$(test35, "0") & " min " & Format$(test36, "#0") & " sec in excess of 30 minutes  "
            lblRemain30.BackColor = &H8000000F   'gray
        End If
        
    Else

        test21 = cRemain2 - Int(cRemain2)
    
        If cRemain2 >= 0 Then
            test22 = cRemain2 - test21
            test23 = test21 * 60
            lblRemain30.Caption = Format$(test22, "0") & " min " & Format$(test23, "#0") & " sec remain of 60 minutes"
            lblRemain30.BackColor = &H8000000F   'gray
        Else
            testRemain2 = cRemain2 * (-1)
            test24 = testRemain2 - Int(testRemain2)
            test25 = testRemain2 - test24
            test26 = test24 * 60
            lblRemain30.Caption = Format$(test25, "0") & " min " & Format$(test26, "#0") & " sec in excess of 60 minutes"
        End If
    
    End If
    
'--------------------------------
    Dim iBlock As Integer
    iBlock = Val(txtBlock)

    If txtBlock <> "" Then 'if planned time block contains an entry

        If cTotalRemain >= 0 Then
            lblRemain60.ForeColor = &H80&    'rust
            shpRunOver.Visible = False      'blue rectangle around text
            lblRemain60.Caption = sRemain60 & " remain of the " & txtBlock & " minutes Planned Program Time"

        ElseIf cTotalRemain < 0 And cTotalRemain > -60 Then 'exceeds 60 minutes
            shpRunOver.Visible = True 'red rectangle around text

            lblRemain60.ForeColor = &HC00000   'Blue

            If txtMinute1 = "" Then
                lblRemain60.Caption = sRemain60 & " in excess of the " & iBlock & " min Planned Program Time"
            Else
                lblRemain60.Caption = sRemain60 & " runover into the next hour"
            End If

        ElseIf cTotalRemain <= -60 Then 'runs into the following hour
            lblRemain60.ForeColor = &HFF& 'Red
            lblRemain60.Caption = sRemain60 & " overrun beyond the next hour"
        End If
    End If

 '------Display Total Time (along with hour for program ending time)
    Dim Today As String
    Dim Hour As Integer
    Dim hour12 As String
    Dim hour24 As String
    Dim sHour As Integer

    If chkStopWatch.Value = 0 Then

        Today = Now

        Hour = (Format(Today, "h"))
       
        If iHourNow <= 12 Then
            iHourNow = iHourNow
        ElseIf iHourNow > 12 Then
            iHourNow = iHourNow - 12
        End If
        
'-----------TEST 5-24-2015

        Dim iMinute1 As Integer
        Dim XX As Integer
        XX = iHourNow
        
        If cmdAdjTime.Visible = False Then

            iMinute1 = Minute(Time) + Val(txtMinAdj)

            If iMinute1 >= 60 Then
                XX = XX + 1
            Else
                XX = XX
            End If
        End If
               
        hour12 = XX
'---------end TEST 5-24-2015

    ElseIf chkStopWatch.Value = 1 And (IsNull(mRunTime) = False) Then
        hour12 = Format(mRunTime, " h")
    End If

'----------- miTotal readout
'    Dim cMinAx As Currency
    Dim cSecAx As Currency
    Dim lLabel As String

    If cMinA < 0 Then
        cMinAx = (cMinA * -1) - 1
        cSecAx = 60 - cSecA
        lLabel = Format$(cMinAx, " 0") & " min " & Format$(cSecAx, "#0") & " sec before the end of the previous hour. "


    ElseIf cMinA > 59 And cMinA <= 63 Then 'sets format for over/under 3 minutes overrun
        cMinAx = cMinA - 60
        cSecAx = cSecA
        lLabel = Format$(cMinAx, " 0") & " min " & Format$(cSecAx, "#0") & " sec past the hour. "
    Else
        cMinAx = cMinA
        cSecAx = cSecA
        lLabel = Format$(cMinAx, " 0") & " min " & Format$(cSecAx, "#0") & " sec past the hour. "

    End If
'-------
    If iiMinute1 < 0 Then
       hour12 = Val(hour12) - 1
    End If

    If cMinA >= 0 Then 'prevents miTotal showing system clock adjustments as programmed time

        If txtMinute1 = "" And txtSecond1 = "" Or cMinA > 59.9 Then 'formats time-added total display as min & sec or as time
    miTotal = Format$(cMinAx, " 0") & " min " & Format$(cSecAx, "#0") & " sec "

        Else
            If chkStopWatch.Value = 0 Then 'to show hour along with music ending time

                If mCurrentTime = 0 Then

                    If txtMinute2 <> "" Or txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
                        Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "" Or txtSecond2 <> "" _
                        Or txtSecond4 <> "" Or txtSecond5 <> "" Or txtSecond6 <> "" Or txtSecond7 <> "" Or txtSecond8 <> "" _
                        Or txtSecond9 <> "" Or txtSecond10 <> "" Or txtSecond11 <> "" Then

                        miTotal = hour12 & ":" & Format$(cMinAx, "00") & ":" & Format$(cSecAx, "00") & " "
                    Else

                        miTotal = hour12 & ":" & Format$(cMinAx, "00") & ":" & Format$(cSecAx, "00") & " "
                    End If
                Else
                    miTotal = lLabel
                End If

            ElseIf chkStopWatch.Value = 1 Then
                If mCurrentTime = 0 Then
                    If hour12 = "" Or hour12 = " 0" Then
                        miTotal = " run time " & Format$(cMinAx, "0") & " min " & Format$(cSecAx, "00") & " sec "
                    ElseIf hour12 <> "" Or hour12 <> " 0" Then
                        miTotal = " run time" & hour12 & " hr " & Format$(cMinAx, "0") & " min " & Format$(cSecAx, "00") & " sec "
                    End If
                Else
                    miTotal = lLabel
                End If
            End If
        End If

    End If
   '--------

    If miTotal = lLabel Then 'to prevent miTotal from running under fraFrame3
        fraFrame3.Visible = False
    Else
        fraFrame3.Visible = True
    End If
   '--------
    'shows total time of music lineup, only if there are entries
    If gcMinNA > "0" And (txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
    Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "") Or _
    (txtMinute1 <> "" And txtMinute2 <> "") Then

        If mCurrentTime = 1 Then 'program begin in previous hour selected
            lblTotalS.Visible = False
        Else
            lblTotalS.Visible = True
        End If

    Else
        lblTotalS.Visible = False
    End If

    lblTotalS.Caption = Format$(gcMinNA, "#0") & ":" & Format$(cSecNA, "00")

    Dim cSecAdd1 As Integer
    Dim cMinAdd1 As Integer
    Dim cMCal11 As Currency
    Dim cMCal12 As Currency
    Dim cMCal13 As Currency
    Dim cMCal14 As Currency
    Dim cHCal11 As Currency
    Dim cSCal12 As Currency
    Dim cHCal13 As Currency
    Dim cHCal14 As Currency

    Dim test44 As Currency
    Dim test45 As Currency
    Dim test46 As Currency
    Dim testRemain4 As Currency

 '----
    'calculation for minutes/seconds display
    cSecAdd1 = aiSecond1 + aiSecond2
    cMinAdd1 = aiMinute1 + aiMinute2
    cMCal11 = cMinAdd1 + cSecAdd1 / 60 'adds min & sec/60

    cMCal12 = cMCal11 - Int(cMCal11) ' leaves fraction
    cMCal13 = cMCal11 - cMCal12 'removes fraction, leaving whole MINUTES
    cMCal14 = cMCal12 * 60 'multiplies fraction by 60 producing SECONDS
    'display cMinA for minutes; display cSecA for seconds

    cMusicRemain = 60 - cMCal11 '60 minus min & decimal min.

    Dim cMins As Currency
    Dim cSecs As Currency

    If cMCal13 < 0 Then '---computation when program begins in previous hour

        If cMCal14 <> 0 Then 'if seconds are not zero
            cMins = (cMCal13 * -1) - 1
            cSecs = 60 - cMCal14
         ElseIf cMCal14 = 0 Then 'if seconds are zero, (example) produces display such as 6:00 rather than 5:60
            cMins = (cMCal13 * -1)
            cSecs = 0
         End If

        Label13.Caption = "prior to the end of the previous hour"
    Else                        '---computation if program begins in the current hour
        cMins = cMCal13
        cSecs = cMCal14
        Label13.Caption = "past the hour"
    End If

    If lblEndTime.Visible = True Then
        If cMusicRemain > 0 Then
            lblEndTime.ForeColor = &H80&      'rust'&H80000008 ' black
            Label12.ForeColor = &H80000008   'black '&HC00000   'blue
           ' imgOnAirSign.Left = 2205
            Label12.Caption = "Currently playing CD ends at"
            lblEndTime.Caption = Format$(cMins, "#0") & " min " & Format$(cSecs, "00") & " sec " 'lblEndTime display

        Else
            testRemain4 = cMusicRemain * (-1)
            test44 = testRemain4 - Int(testRemain4)
            test45 = testRemain4 - test44
            test46 = test44 * 60
            lblEndTime.ForeColor = vbRed
            Label13.Caption = ""
            imgOnAirSign.Visible = False
            Label12.ForeColor = vbRed
            Label12.Caption = "• Currently playing CD carries over into the next hour by "
            lblEndTime.Caption = Format$(test45, "0") & " min " & Format$(test46, "00") & " sec"
        End If
    End If

    If frmPlanner!mnuLinkTimeRemain.Checked = True Then

        frmPlanner!staStatus.Panels(4) = gsRemain1

        If giTimesDiffer = 0 Then
            frmPlanner!lblProgramTime.Visible = True
        Else
            frmPlanner!lblProgramTime.Caption = " Planning and program lineup times differ "
        End If

        If frmPlanner!staStatus.Panels(2) = "" Then
            frmPlanner!staStatus.Panels(3) = "Program " & miTotal

            If cTotalRemain < 0 And cTotalRemain > -60 Then
                frmPlanner!staStatus.Panels(3) = "Program " & sRemain60 & " overrun"
            Else
                frmPlanner!staStatus.Panels(3) = ""
            End If
        Else
            If cTotalMinA < 60 Then
                frmPlanner!staStatus.Panels(3) = "Program ending" & miTotal

            ElseIf cTotalMinA >= 60 And cTotalMinA < 63 Then
                frmPlanner!staStatus.Panels(3) = "Program ends" & miTotal & "into the next hour"

            Else
                frmPlanner!staStatus.Panels(3) = "*Program ending overrun of 3 min or more"
            End If
        End If
    End If

    '--------------Program Information (lblProgramInfo) text
'1
    If cMinA = 60 And cSecA = 0 And chkStopWatch.Value = 0 Then
        lblProgramInfo.Caption = "  Program ends on the hour: " & miTotal & " "
        imgClock.Visible = True
        If cTotalRemain > 3.16 Then
            imgMusic.Visible = True
        End If
 '2
    ElseIf cMinA >= 60 And cMinA < 64 Then
        lblProgramInfo.Caption = "  You have programmed into the next hour by " & miTotal & " "
        imgClock.Visible = False
        imgMusic.Visible = False
        lblProgramInfo.ForeColor = &H80&  'rust

    ElseIf cTotalMinA > 60 And cTotalMinA < 120 Then
        If cTotalMinA < 65 Then
            lblProgramInfo.ForeColor = &H80& 'rust
        Else
            lblProgramInfo.ForeColor = &H80000008  'windows text
        End If
'3
        lblProgramInfo.Caption = "  You have programmed " & miTotal & " "
        imgClock.Visible = False
        imgMusic.Visible = False
'4
    ElseIf cTotalMinA >= 120 Then
        lblProgramInfo.ForeColor = vbRed
        lblProgramInfo.Caption = "  This program total extends beyond the next hour: " & miTotal & " "
        imgClock.Visible = False
        imgMusic.Visible = False
    Else
        lblProgramInfo.ForeColor = &H80& 'rust

        If txtMinute1 <> "" Or txtSecond1 <> "" Then

            If txtMinAdj <> "" Or txtSecAdj <> "" And chkStopWatch.Value = 0 Then
                lblProgramInfo.Caption = "  * Program ends at " & miTotal & " "
               
                imgClock.Visible = True
                If cTotalRemain > 3.16 Then
                    imgMusic.Visible = True
                End If
            Else
'5
                If txtMinute2 <> "" Or txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
                Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "" Or txtSecond2 <> "" _
                Or txtSecond4 <> "" Or txtSecond5 <> "" Or txtSecond6 <> "" Or txtSecond7 <> "" Or txtSecond8 <> "" _
                Or txtSecond9 <> "" Or txtSecond10 <> "" Or txtSecond11 <> "" Then
                   
                If cMinAx >= 50 Then
                    lblProgramInfo.Caption = " Including annc time what you have programmed will end at " & miTotal & ""
                Else
                    lblProgramInfo.Caption = " What you have programmed will end at " & miTotal & " "
                End If
    
                Else

                    If iCurrentTime = 1 Then '6
                      
                        lblProgramInfo.Caption = "  Current time past the hour is " & miTotal & " "
                       
                    Else
                        lblProgramInfo.Caption = "  You have entered the current time past the hour as " & miTotal & " "
                    End If
                End If

                If chkStopWatch.Value = 0 And iCurrentTime = 1 Then
                    imgClock.Visible = True
                    If cTotalRemain > 3.16 Then
                        imgMusic.Visible = True
                    End If
                End If
            End If
'7
        ElseIf txtMinute1 = "" And txtSecond1 = "" Then
            lblProgramInfo.Caption = "  You have programmed " & miTotal & " "
            imgClock.Visible = False
            imgMusic.Visible = False
        End If
    End If

    If cTotalRemain < 3.16 Then
        imgMusic.Visible = False
    End If

    If Val(txtMinute1) >= 57 Then
        imgClock.Visible = False
    End If

    If (txtMinute1 <> "" Or txtSecond1 <> "") And (txtMinute2 <> "" Or txtSecond2 <> "") And Label12.ForeColor <> vbRed Then
        imgOnAirSign.Visible = True

    ElseIf txtMinute2 = "" And txtSecond2 = "" Then
        imgOnAirSign.Visible = False
    End If

    If txtMinute1 <> "" And (txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" Or txtMinute8 <> "" Or txtMinute9 <> "" _
     Or txtMinute10 <> "" Or txtMinute11 <> "") Then
        Label10.ForeColor = &HC00000   'blue
        imgCheck.Visible = True
        imgCheck2.Visible = True
    Else
        Label10.ForeColor = &HC0C0C0   '&H808080 'gray &HC0C0C0   'lite gray '&H808080 'gray
        imgCheck.Visible = False
        imgCheck2.Visible = False
    End If

    If frmPlanner!lblTotal1 <> "" Then 'shows program time only if there is music time shown
        frmPlanner!lblProgramTime = lblTotalS & " " & sPlannerProgramTime & "  •  " & sRemain
    Else
        frmPlanner!lblProgramTime = ""
    End If

End Sub

Private Sub Check1_Click(Index As Integer)
    '1
    If Check1(0).Value = 0 And txtMinute4 <> "" And txtAnnc4 = "" Then
    
        If txtBackAnnc = "" Or txtBackAnnc = "0" Then
            txtAnnc4 = Val(txtIntro)
        Else
            txtAnnc4 = Val(txtIntro) - Val(txtBackAnnc)
        End If
        
    ElseIf Check1(0).Value = 1 And txtMinute4 <> "" Then
        txtAnnc4 = ""
    End If
    
    '2
    If Check1(1).Value = 0 And txtMinute5 <> "" And txtAnnc5 = "" Then
        txtAnnc5 = txtIntro
    ElseIf Check1(1).Value = 1 And txtMinute5 <> "" Then
        txtAnnc5 = ""
    End If
    '3
    If Check1(2).Value = 0 And txtMinute6 <> "" And txtAnnc6 = "" Then
        txtAnnc6 = txtIntro
    ElseIf Check1(2).Value = 1 And txtMinute6 <> "" Then
        txtAnnc6 = ""
    End If
    '4
    If Check1(3).Value = 0 And txtMinute7 <> "" And txtAnnc7 = "" Then
        txtAnnc7 = txtIntro
    ElseIf Check1(3).Value = 1 And txtMinute7 <> "" Then
        txtAnnc7 = ""
    End If
    '5
    If Check1(4).Value = 0 And txtMinute8 <> "" And txtAnnc8 = "" Then
        txtAnnc8 = txtIntro
    ElseIf Check1(4).Value = 1 And txtMinute8 <> "" Then
        txtAnnc8 = ""
    End If
    '6
    If Check1(5).Value = 0 And txtMinute9 <> "" And txtAnnc9 = "" Then
        txtAnnc9 = txtIntro
    ElseIf Check1(5).Value = 1 And txtMinute9 <> "" Then
        txtAnnc9 = ""
    End If
    '7
    If Check1(6).Value = 0 And txtMinute10 <> "" And txtAnnc10 = "" Then
        txtAnnc10 = txtIntro
    ElseIf Check1(6).Value = 1 And txtMinute10 <> "" Then
        txtAnnc10 = ""
    End If
    '8
    If Check1(7).Value = 0 And txtMinute11 <> "" And txtAnnc11 = "" Then
        txtAnnc11 = txtIntro
    ElseIf Check1(7).Value = 1 And txtMinute11 <> "" Then
        txtAnnc11 = ""
    End If
        
End Sub

Private Sub chkBackAnnc_Click()

    Check1(0).Value = False

    If chkBackAnnc.Value = 1 Then
 
        If Val(txtIntro) > 29 Then
            txtBackAnnc = "0"
            chkBackAnnc.Caption = "Uncheck to set back-announce time to " & Format((Val((txtIntro) / 2) - 10), "##") & " sec "
            Label21.ToolTipText = ""
            txtBackAnnc.BackColor = &H80000016  ' lite gray
            txtBackAnnc.Enabled = False
            
        ElseIf txtIntro <= 29 Then
         
            chkBackAnnc.Caption = "Back announce time is 0"
            txtBackAnnc = "0"
        End If
        
        pSetFocus
'-----------
        
    ElseIf chkBackAnnc.Value = 0 Then
        
        txtBackAnnc.Enabled = True
        If Val(txtIntro) > 29 Then
            txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
        Else
            txtBackAnnc = "0"
        End If
        txtBackAnnc.BackColor = &HFFFFFF 'white
        chkBackAnnc.Caption = "Check if CD will not be back-announced"
 
        pSetFocus
    End If
    
    txtBackAnnc.Enabled = True

End Sub

Private Sub ck10Played_Click()

    If txtMinute10 = "" Then
        ck10Played.Value = 0
    End If

    If ck10Played.Value = 0 Then
        txtMinute10.Enabled = True
        txtMinute10.BackColor = &H80000005  'white
        txtSecond10.Enabled = True
        txtSecond10.BackColor = &H80000005  'white
        txtAnnc10.Enabled = True
        txtComposer10.Enabled = True
    End If
    
    Check1(6).Value = 0
    
    'adjusts announce time when check-CD-Played is clicked
    If ck10Played.Value = 1 And txtComposer10 <> "" And (txtMinute10 <> "" Or txtSecond10 <> "") Then
        iiAnnc10 = "0"
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = False And ck10Played.Value = 1 And (txtMinute10 <> "" Or txtSecond10 <> "") Then
        Label32(6).Caption = txtMinute10 & ":" & txtSecond10
    ElseIf ck10Played.Value = 0 Then
        Label32(6).Caption = ""
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = True And ck10Played.Value = 1 And (frmPlanner!txtMinute7 <> "" Or frmPlanner!txtSecond7 <> "") Then
        Label32(6).Caption = frmPlanner!txtMinute7 & ":" & frmPlanner!txtSecond7
    ElseIf ck10Played.Value = 0 Then
        Label32(6).Caption = ""
    End If

    If ck10Played.Value = 1 Then
    
        iTMin10 = Val(txtMinute10)
        txtMinute10 = ""
        iTSec10 = Val(txtSecond10)
        txtSecond10 = ""
        txtAnnc10.Enabled = False
        txtComposer10.Enabled = False
        txtMinute10.Enabled = False
        txtMinute10.BackColor = &H80000018  'lite yellow
        txtSecond10.Enabled = False
        txtSecond10.BackColor = &H80000018  'lite yellow
        cmdClearPads.Visible = True
        
    ElseIf ck10Played.Value = 0 Then
        txtMinute10 = iTMin10
        txtSecond10 = iTSec10
        If txtMinute10 <> "" Then
             If Annc10 <> "" Then
                txtAnnc10 = Annc10
            Else
                txtAnnc10 = txtIntro
            End If
        End If
        
        If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
        And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then
            cmdClearPads.Visible = False
        End If
    End If
    
    pSetFocus
    Exit Sub

End Sub

Private Sub ck11Played_Click()

    If txtMinute11 = "" Then
        ck11Played.Value = 0
    End If

    If ck11Played.Value = 0 Then
        txtMinute11.Enabled = True
        txtMinute11.BackColor = &H80000005  'white
        txtSecond11.Enabled = True
        txtSecond11.BackColor = &H80000005  'white
        txtAnnc11.Enabled = True
        txtComposer11.Enabled = True
    End If
    
    Check1(7).Value = 0
    
    'adjusts announce time when check-CD-Played is clicked
    If ck11Played.Value = 1 And txtComposer11 <> "" And (txtMinute11 <> "" Or txtSecond11 <> "") Then
        iiAnnc11 = "0"
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = False And ck11Played.Value = 1 And (txtMinute11 <> "" Or txtSecond11 <> "") Then
        Label32(7).Caption = txtMinute11 & ":" & txtSecond11
    ElseIf ck11Played.Value = 0 Then
        Label32(7).Caption = ""
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = True And ck11Played.Value = 1 And (frmPlanner!txtMinute8 <> "" Or frmPlanner!txtSecond8 <> "") Then
        Label32(7).Caption = frmPlanner!txtMinute8 & ":" & frmPlanner!txtSecond8
    ElseIf ck11Played.Value = 0 Then
        Label32(7).Caption = ""
    End If

    If ck11Played.Value = 1 Then 'play info hidden
    
        iTMin11 = Val(txtMinute11)
        txtMinute11 = ""
        iTSec11 = Val(txtSecond11)
        txtSecond11 = ""
        txtAnnc11.Enabled = False
        txtComposer11.Enabled = False
        txtMinute11.Enabled = False
        txtMinute11.BackColor = &H80000018  'lite yellow
        txtSecond11.Enabled = False
        txtSecond11.BackColor = &H80000018  'lite yellow
        cmdClearPads.Visible = True
        
    ElseIf ck11Played.Value = 0 Then 'play info visible
        txtMinute11 = iTMin11
        txtSecond11 = iTSec11
        
        If txtMinute11 <> "" Then
             If Annc11 <> "" Then
                txtAnnc11 = Annc11
            Else
                txtAnnc11 = txtIntro
            End If
        End If
        If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
        And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then
            cmdClearPads.Visible = False
        End If
    End If
    
    txtSecond11.Text = Format$(txtSecond11, "00")
    pSetFocus
    Exit Sub

End Sub

Private Sub ck4Played_Click()

    If txtMinute4 = "" Then
        ck4Played.Value = 0
    End If
    
    If ck4Played.Value = 0 Then
        txtMinute4.Enabled = True
        txtMinute4.BackColor = &H80000005  'white
        txtSecond4.Enabled = True
        txtSecond4.BackColor = &H80000005 'white
        txtAnnc4.Enabled = True
        txtComposer4.Enabled = True
    End If
    
    Check1(0).Value = 0

    If ck4Played.Value = 1 And txtComposer4 <> "" And (txtMinute4 <> "" Or txtSecond4 <> "") Then
        iiAnnc4 = "0" 'iiAnnc4 & iAnnc4 dim as string
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = False And ck4Played.Value = 1 And (txtMinute4 <> "" Or txtSecond4 <> "") Then
        Label32(0).Caption = txtMinute4 & ":" & txtSecond4
    ElseIf ck4Played.Value = 0 Then
        Label32(0).Caption = ""
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = True And ck4Played.Value = 1 And (frmPlanner!txtMinute1 <> "" Or frmPlanner!txtSecond1 <> "") Then
        Label32(0).Caption = frmPlanner!txtMinute1 & ":" & frmPlanner!txtSecond1
    ElseIf ck4Played.Value = 0 Then
        Label32(0).Caption = ""
    End If
 '-------------------------------------
    If ck4Played.Value = 1 Then

        iTMin4 = Val(txtMinute4)
        txtMinute4 = ""
        iTSec4 = Val(txtSecond4)
        txtSecond4 = ""
        txtAnnc4.Enabled = False
        txtComposer4.Enabled = False
        txtMinute4.Enabled = False
        txtMinute4.BackColor = &H80000018  'lite yellow
        txtSecond4.Enabled = False
        txtSecond4.BackColor = &H80000018  'lite yellow
        cmdClearPads.Visible = True
        
    ElseIf ck4Played.Value = 0 Then
        txtMinute4 = iTMin4
        txtSecond4 = iTSec4
        
'-------------due occasional mis-match error, modified to integter for minus computation 8-22-17
        Dim vAnnc4 As Integer
        If txtMinute4 <> "" Or txtSecond4 <> "" Then
            If txtBackAnnc <> "" Then
                vAnnc4 = Val(txtIntro) - Val(txtBackAnnc)
            Else
                vAnnc4 = txtIntro
            End If
        End If
        txtAnnc4 = vAnnc4
'------------
        If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
        And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then
            cmdClearPads.Visible = False
        End If
        
    End If

    pSetFocus
    Exit Sub

End Sub
Private Sub ck5Played_Click()

    If txtMinute5 = "" Then
        ck5Played.Value = 0
    End If

    If ck5Played.Value = 0 Then 'enabled
        txtMinute5.Enabled = True
        txtMinute5.BackColor = &H80000005  'white
        txtSecond5.Enabled = True
        txtSecond5.BackColor = &H80000005  'white
        txtAnnc5.Enabled = True
        txtComposer5.Enabled = True
    End If
    
    Check1(1).Value = 0
    '--------------
    If frmPlanner!mnuLinkTimeRemain.Checked = False And ck5Played.Value = 1 And (txtMinute5 <> "" Or txtSecond5 <> "") Then
        Label32(1).Caption = txtMinute5 & ":" & txtSecond5
    ElseIf ck5Played.Value = 0 Then
        Label32(1).Caption = ""
    End If
    '----------
    If ck5Played.Value = 1 And txtComposer5 <> "" And (txtMinute5 <> "" Or txtSecond5 <> "") Then
       iiAnnc5 = "0"
    End If
    '----------------
    If frmPlanner!mnuLinkTimeRemain.Checked = True And ck5Played.Value = 1 And (frmPlanner!txtMinute2 <> "" Or frmPlanner!txtSecond2 <> "") Then
        Label32(1).Caption = frmPlanner!txtMinute2 & ":" & frmPlanner!txtSecond2
    ElseIf ck5Played.Value = 0 Then
        Label32(1).Caption = ""
    End If
'------------------------------------------------
On Error GoTo HandleErrors
    If ck5Played.Value = 1 Then
    
        iTMin5 = Val(txtMinute5)
        txtMinute5 = ""
        iTSec5 = Val(txtSecond5)
        txtSecond5 = ""
        txtAnnc5.Enabled = False
        txtComposer5.Enabled = False
        txtMinute5.Enabled = False
        txtMinute5.BackColor = &H80000018  'lite yellow
        txtSecond5.Enabled = False
        txtSecond5.BackColor = &H80000018  'lite yellow
        cmdClearPads.Visible = True
'------------------
    ElseIf ck5Played.Value = 0 Then
        txtMinute5 = iTMin5
        txtSecond5 = iTSec5
        If txtMinute5 <> "" Then
             If Annc5 <> "" Then
                txtAnnc5 = Annc5
            Else
                txtAnnc5 = txtIntro
            End If
        End If
        
        If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
        And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then
            cmdClearPads.Visible = False
        End If
    End If

    pSetFocus
    Exit Sub
HandleErrors:
    Close #501
End Sub

Private Sub ck6Played_Click()

    If txtMinute6 = "" Then
        ck6Played.Value = 0
    End If

    If ck6Played.Value = 0 Then
        txtMinute6.Enabled = True
        txtMinute6.BackColor = &H80000005  'white
        txtSecond6.Enabled = True
        txtSecond6.BackColor = &H80000005  'white
        txtAnnc6.Enabled = True
        txtComposer6.Enabled = True
    End If
    
    Check1(2).Value = 0
    
    'adjusts announce time when check-CD-Played is clicked
    If ck6Played.Value = 1 And txtComposer6 <> "" And (txtMinute6 <> "" Or txtSecond6 <> "") Then
        iiAnnc6 = "0"
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = False And ck6Played.Value = 1 And (txtMinute6 <> "" Or txtSecond6 <> "") Then
        Label32(2).Caption = txtMinute6 & ":" & txtSecond6
    ElseIf ck6Played.Value = 0 Then
        Label32(2).Caption = ""
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = True And ck6Played.Value = 1 And (frmPlanner!txtMinute3 <> "" Or frmPlanner!txtSecond3 <> "") Then
        Label32(2).Caption = frmPlanner!txtMinute3 & ":" & frmPlanner!txtSecond3
    ElseIf ck6Played.Value = 0 Then
        Label32(2).Caption = ""
    End If
 
    If ck6Played.Value = 1 Then
    
        iTMin6 = Val(txtMinute6)
        txtMinute6 = ""
        iTSec6 = Val(txtSecond6)
        txtSecond6 = ""
        txtAnnc6.Enabled = False
        txtComposer6.Enabled = False
        txtMinute6.Enabled = False
        txtMinute6.BackColor = &H80000018  'lite yellow
        txtSecond6.Enabled = False
        txtSecond6.BackColor = &H80000018  'lite yellow
        cmdClearPads.Visible = True
        
    ElseIf ck6Played.Value = 0 Then
        txtMinute6 = iTMin6
        txtSecond6 = iTSec6
        If txtMinute6 <> "" Then
             If Annc6 <> "" Then
                txtAnnc6 = Annc6
            Else
                txtAnnc6 = txtIntro
            End If
        End If
        
        If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
        And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then
            cmdClearPads.Visible = False
        End If
    End If

    pSetFocus
    Exit Sub
    
End Sub

Private Sub ck7Played_Click()

    If txtMinute7 = "" Then
        ck7Played.Value = 0
    End If

    If ck7Played.Value = 0 Then
        txtMinute7.Enabled = True
        txtMinute7.BackColor = &H80000005  'white
        txtSecond7.Enabled = True
        txtSecond7.BackColor = &H80000005  'white
        txtAnnc7.Enabled = True
        txtComposer7.Enabled = True
    End If
    
    Check1(3).Value = 0
    
    'adjusts announce time when check-CD-Played is clicked
    If ck7Played.Value = 1 And txtComposer7 <> "" And (txtMinute7 <> "" Or txtSecond7 <> "") Then
        iiAnnc7 = "0"
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = False And ck7Played.Value = 1 And (txtMinute7 <> "" Or txtSecond7 <> "") Then
        Label32(3).Caption = txtMinute7 & ":" & txtSecond7
    ElseIf ck7Played.Value = 0 Then
        Label32(3).Caption = ""
    End If
 
    If frmPlanner!mnuLinkTimeRemain.Checked = True And ck7Played.Value = 1 And (frmPlanner!txtMinute4 <> "" Or frmPlanner!txtSecond4 <> "") Then
        Label32(3).Caption = frmPlanner!txtMinute4 & ":" & frmPlanner!txtSecond4
    ElseIf ck7Played.Value = 0 Then
        Label32(3).Caption = ""
    End If

    If ck7Played.Value = 1 Then
    
        iTMin7 = Val(txtMinute7)
        txtMinute7 = ""
        iTSec7 = Val(txtSecond7)
        txtSecond7 = ""
        txtAnnc7.Enabled = False
        txtComposer7.Enabled = False
        txtMinute7.Enabled = False
        txtMinute7.BackColor = &H80000018  'lite yellow
        txtSecond7.Enabled = False
        txtSecond7.BackColor = &H80000018  'lite yellow
        cmdClearPads.Visible = True
        
    ElseIf ck7Played.Value = 0 Then
        txtMinute7 = iTMin7
        txtSecond7 = iTSec7
        If txtMinute7 <> "" Then
             If Annc7 <> "" Then
                txtAnnc7 = Annc7
            Else
                txtAnnc7 = txtIntro
            End If
        End If
        
        If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
        And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then
            cmdClearPads.Visible = False
        End If
    End If

    pSetFocus
    Exit Sub

End Sub

Private Sub ck8Played_Click()

    If txtMinute8 = "" Then
        ck8Played.Value = 0
    End If

    If ck8Played.Value = 0 Then
        txtMinute8.Enabled = True
        txtMinute8.BackColor = &H80000005  'white
        txtSecond8.Enabled = True
        txtSecond8.BackColor = &H80000005  'white
        txtAnnc8.Enabled = True
        txtComposer8.Enabled = True
    End If
    
    Check1(4).Value = 0
    
    'adjusts announce time when check-CD-Played is clicked
    If ck8Played.Value = 1 And txtComposer8 <> "" And (txtMinute8 <> "" Or txtSecond8 <> "") Then
        iiAnnc8 = "0"
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = False And ck8Played.Value = 1 And (txtMinute8 <> "" Or txtSecond8 <> "") Then
        Label32(4).Caption = txtMinute8 & ":" & txtSecond8
    ElseIf ck8Played.Value = 0 Then
        Label32(4).Caption = ""
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = True And ck8Played.Value = 1 And (frmPlanner!txtMinute5 <> "" Or frmPlanner!txtSecond5 <> "") Then
        Label32(4).Caption = frmPlanner!txtMinute5 & ":" & frmPlanner!txtSecond5
    ElseIf ck8Played.Value = 0 Then
        Label32(4).Caption = ""
    End If

    If ck8Played.Value = 1 Then
    
        iTMin8 = Val(txtMinute8)
        txtMinute8 = ""
        iTSec8 = Val(txtSecond8)
        txtSecond8 = ""
        txtAnnc8.Enabled = False
        txtComposer8.Enabled = False
        txtMinute8.Enabled = False
        txtMinute8.BackColor = &H80000018  'lite yellow
        txtSecond8.Enabled = False
        txtSecond8.BackColor = &H80000018  'lite yellow
        cmdClearPads.Visible = True
        
    ElseIf ck8Played.Value = 0 Then
        txtMinute8 = iTMin8
        txtSecond8 = iTSec8
        If txtMinute8 <> "" Then
             If Annc8 <> "" Then
                txtAnnc8 = Annc8
            Else
                txtAnnc8 = txtIntro
            End If
        End If
        
        If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
        And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then
            cmdClearPads.Visible = False
        End If
    End If

    pSetFocus
    Exit Sub
 
End Sub

Private Sub ck9Played_CliCk()

    If txtMinute9 = "" Then
        ck9Played.Value = 0
    End If

    If ck9Played.Value = 0 Then
        txtMinute9.Enabled = True
        txtMinute9.BackColor = &H80000005  'white
        txtSecond9.Enabled = True
        txtSecond9.BackColor = &H80000005  'white
        txtAnnc9.Enabled = True
        txtComposer9.Enabled = True
    End If
    
    Check1(5).Value = 0
    
    'adjusts announce time when check-CD-Played is clicked
    If ck9Played.Value = 1 And txtComposer9 <> "" And (txtMinute9 <> "" Or txtSecond9 <> "") Then
        iiAnnc9 = "0"
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = False And ck9Played.Value = 1 And (txtMinute9 <> "" Or txtSecond9 <> "") Then
        Label32(5).Caption = txtMinute9 & ":" & txtSecond9
    ElseIf ck9Played.Value = 0 Then
        Label32(5).Caption = ""
    End If
    
  If frmPlanner!mnuLinkTimeRemain.Checked = True And ck9Played.Value = 1 And (frmPlanner!txtMinute6 <> "" Or frmPlanner!txtSecond6 <> "") Then
        Label32(5).Caption = frmPlanner!txtMinute6 & ":" & frmPlanner!txtSecond6
    ElseIf ck9Played.Value = 0 Then
        Label32(5).Caption = ""
    End If

    If ck9Played.Value = 1 Then
    
        iTMin9 = Val(txtMinute9)
        txtMinute9 = ""
        iTSec9 = Val(txtSecond9)
        txtSecond9 = ""
        txtAnnc9.Enabled = False
        txtComposer9.Enabled = False
        txtMinute9.Enabled = False
        txtMinute9.BackColor = &H80000018  'lite yellow
        txtSecond9.Enabled = False
        txtSecond9.BackColor = &H80000018  'lite yellow
        cmdClearPads.Visible = True
        
    ElseIf ck9Played.Value = 0 Then
        txtMinute9 = iTMin9
        txtSecond9 = iTSec9
        If txtMinute9 <> "" Then
             If Annc9 <> "" Then
                txtAnnc9 = Annc9
            Else
                txtAnnc9 = txtIntro
            End If
        End If
        
        If ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
        And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0 Then
            cmdClearPads.Visible = False
        End If
    End If

    pSetFocus
    Exit Sub

End Sub

Private Sub chkAnnounce_Click()

    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
    Check1(6).Value = 0
    Check1(7).Value = 0

    If chkAnnounce.Value = 1 Then 'music selections do NOT include announce time
        txtMinute3 = ""
        txtSecond3 = ""
        
        txtMinute3.Visible = False
        txtSecond3.Visible = False
        lblS.Visible = False
        lblMinSec3Div.Visible = False
        shpTime3.Visible = False
        lblAnncMin.Visible = False
        lblAnncSec.Visible = False
        Label9.Visible = False
        
        txtSpotsS = ""
        txtSpotsS.Visible = False 'test spots
        lblSpots.Visible = False
        txtSpotLength.Visible = False
        Label3.Visible = False
        
        txtCloseOut.Visible = False
        Label24.Visible = False
        Label26.Visible = False
        
        'Label9.Alignment = 2
        Label9.ForeColor = &HC00000   'Blue
        Label9.Caption = "You can replace the program's estimated announce time with your estimate of the announce time you will need:"
    
        frmPlanner!mnuLinkTimeRemain.Checked = False

        frmPlanner!staStatus.Panels(4) = "F4 to Link Lineup to Time Remain Page"
        frmPlanner!txtSpots.Visible = False
        frmPlanner!Shape1.Visible = False
                
        lblAverageTime.Visible = False
        txtIntro.Visible = False
        Frame12.Height = 750
        Label17.Visible = False
        Frame8.Enabled = False
        'Frame8.Caption = ""

        chkAnnounce.ForeColor = &HC00000   'Blue '&H404040   '&H80000012 'black
        fraAnnounce.ForeColor = &HC00000   'Blue

        Label21.Visible = False
        chkBackAnnc.Visible = False
        chkBackAnnc.Value = 0
        txtBackAnnc.Visible = False
        txtBackAnnc = ""
        cmdClearAnncTimes.Enabled = False

        Label20.Visible = False
         
        txtAnnc4.Visible = False
        txtAnnc5.Visible = False
        txtAnnc6.Visible = False
        txtAnnc7.Visible = False
        txtAnnc8.Visible = False
        txtAnnc9.Visible = False
        txtAnnc10.Visible = False
        txtAnnc11.Visible = False
        Label22.Visible = False
        lblAnncTime.Visible = False
        Frame2.Caption = "Time Remaining"
        Frame4.Caption = ""
         
        If lblLinked.Visible = True Then
            lblLinked_DblClick
        End If
        
        Check1(0).Enabled = False
        Check1(1).Enabled = False
        Check1(2).Enabled = False
        Check1(3).Enabled = False
        Check1(4).Enabled = False
        Check1(5).Enabled = False
        Check1(6).Enabled = False
        Check1(7).Enabled = False
            
    ElseIf chkAnnounce.Value = 0 Then 'includes announce times
    
        txtMinute3 = ""
        txtSecond3 = ""
        
        txtSpotsS.Visible = True
        lblSpots.Visible = True
      
        lblProgramRemain.Visible = True
        txtAnnc4.Visible = True
        txtAnnc5.Visible = True
        txtAnnc6.Visible = True
        txtAnnc7.Visible = True
        txtAnnc8.Visible = True
        txtAnnc9.Visible = True
        txtAnnc10.Visible = True
        txtAnnc11.Visible = True
        Label22.Visible = True
        lblAnncTime.Visible = True
        txtSpotLength.Visible = True
        Label3.Visible = True
        Label17.Visible = True
        Frame8.Enabled = True
'        'Frame8.Caption = miCloseOut & " Sec Allocated for Closeout && ID"
        'Frame8.Caption = " Seconds Allocated for Closeout && ID"
        
        fraAnnounce.Enabled = True
        fraAnnounce.Caption = "Estimated Announce Time"
        
        chkAnnounce.ForeColor = &H80& 'rust
        fraAnnounce.ForeColor = &H80& 'rust
        fraAnnounce.Caption = "Estimated Announce Time"
 
        If lblEndTime.Visible = True Then
            Label21.Visible = True
            chkBackAnnc.Visible = True
            chkBackAnnc.Value = 0
            txtBackAnnc.Visible = True
            
            If Val(txtIntro) > 29 And chkBackAnnc.Visible = True Then
                txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
            Else
                txtBackAnnc = "0"
            End If
           
        End If
        
        txtCloseOut.Visible = True
        Label24.Visible = True
        Label26.Visible = True
        cmdClearAnncTimes.Enabled = True

        Label20.Visible = True
        
        'Label9.Alignment = 0
        Label9.ForeColor = &H80&       'rust
        Label9.Caption = "You can replace the program's estimated announce time with your estimate of the announce time you will need:"

        lblAverageTime.Visible = True
        txtIntro.Visible = True
        Frame12.Height = 1200
       
        Check1(0).Enabled = True
        Check1(1).Enabled = True
        Check1(2).Enabled = True
        Check1(3).Enabled = True
        Check1(4).Enabled = True
        Check1(5).Enabled = True
        Check1(6).Enabled = True
        Check1(7).Enabled = True
    End If
    pAnnounce
    pSetFocus
End Sub

Private Sub chkStopWatch_Click()
    pSetFocus
    Frame7.Caption = "Set Current Time"
    pAdd
    If chkStopWatch.Value = 1 Then 'stopwatch in use
        frmStopWatch.Show
        cmdStopwatch.Enabled = False
        
        fraAdjustTime.Visible = False
        fraAdjustedTime.Visible = False
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Stopwatch) F5"
        lblHour.Visible = False
        Shape1.Visible = True
        imgClock.Visible = False
        imgMusic.Visible = False
       
        txtMinAdj.Text = ""
        txtSecAdj.Text = ""
        txtMinAdj.BackColor = &HE0E0E0
        txtSecAdj.BackColor = &HE0E0E0
        txtMinAdj.Enabled = False
        txtSecAdj.Enabled = False
       
    ElseIf chkStopWatch.Value = 0 Then 'stop watch not used
    
        If txtMinAdj = "" And txtSecAdj = "" Then
            cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
        Else
             cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
        End If
        
        txtMinute1.Text = ""
        txtSecond1.Text = ""
        txtMinute2.Text = ""
        txtSecond2.Text = ""
        cmdSystemTime.Visible = True
        txtMinAdj.BackColor = &H80000005
        txtSecAdj.BackColor = &H80000005
        txtMinAdj.Enabled = True
        txtSecAdj.Enabled = True
        
    End If
End Sub

Private Sub cmdAddTime_Click()
      
    If frmAddTime.Visible = True Then
        
        If Val(frmAddTime!lblTotal1) <= 0 Then
            cmdAddTime.ToolTipText = " Calculator for adding up times entered as minutes & seconds "
            cmdAddTime.Caption = "AddTi&me F9"
        End If
        
        frmAddTime.Hide
        
       If txtComposer4 <> "" And txtMinute4 = "" Then
            txtMinute4.SetFocus
        ElseIf txtComposer5 <> "" And txtMinute5 = "" Then
            txtMinute5.SetFocus
        ElseIf txtComposer6 <> "" And txtMinute6 = "" Then
            txtMinute6.SetFocus
        ElseIf txtComposer7 <> "" And txtMinute7 = "" Then
            txtMinute7.SetFocus
        ElseIf txtComposer8 <> "" And txtMinute8 = "" Then
            txtMinute8.SetFocus
        ElseIf txtComposer9 <> "" And txtMinute9 = "" Then
            txtMinute9.SetFocus
        ElseIf txtComposer10 <> "" And txtMinute10 = "" Then
            txtMinute10.SetFocus
        ElseIf txtComposer11 <> "" And txtMinute11 = "" Then
            txtMinute11.SetFocus
        
        Else
            pSetFocus
        End If

        Exit Sub
    End If
    
    cmdAddTime.ToolTipText = " Close AddTime Calculator "
 
    frmAddTime.Show
    Unload frmAddHelp
 
End Sub

Private Sub cmdAdjTime_Click()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If

    Dim iResponse As Integer

    msgSelect = msgSelect + 1

   If msgSelect = 1 Then
    
        iResponse = MsgBox("The timing accuracy of this program depends upon the computer clock being set to the correct time." _
        & vbCrLf & vbCrLf & "If the computer clock time is incorrect and you are unable to adjust it, here is a way to adjust program time to compensate." & vbCrLf & vbCrLf & _
            "•  If the COMPUTER CLOCK is SLOW, enter the time difference between the computer clock time and the correct time." & vbCrLf & vbCrLf & _
        "•  If the COMPUTER CLOCK is FAST, enter the time difference between the computer clock time and the correct time as a  NEGATIVE number (precede the number with the - minus sign). " _
        & vbCrLf & vbCrLf & "In summary: If computer clock is slow, enter the time difference to catch up. If computer clock is fast, enter the time difference as a negative to drop back.", _
         vbOKCancel, "Adjusting Program Time to Compensate for Computer Clock Time Error")
    
    ElseIf msgSelect > 1 Then
        iResponse = MsgBox("Computer clock SLOW?  Enter time adjustment as a positive number to catch up with the correct time" & vbCrLf & vbCrLf & _
        " Computer clock FAST?  Enter time adjustment as a negative number to fall back to the correct time", vbOKCancel, _
        "Adjusting Program Time to Compensate for Computer Clock Time Error")
    End If
    
    If iResponse = vbOK Then

        cmdAdjTime.Visible = False
        fraTimeAdjust2.Visible = False
        
        iHourNow = Hour(Now)
        
        Dim Y As Integer
            
        If iHourNow <= 12 Then 'set to 12 hour time
            Y = iHourNow
        ElseIf iHourNow > 12 Then
            Y = iHourNow - 12
        End If
        
        lblHourAdj = Y

       txtMinAdj.Visible = True
       txtSecAdj.Visible = True
       Label29.Visible = True
       lblMinutes.Visible = True
       lblSeconds.Visible = True
       fraAdjustTime.Visible = True
       fraAdjustedTime.Visible = True
       txtMinAdj.SetFocus
    End If
End Sub

Private Sub cmdClearAnncTimes_Click()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
    
    mnuToolsExportLineCopy.Checked = False

    If txtMinute4 = "" And txtMinute5 = "" And txtMinute6 = "" And txtMinute7 = "" And txtMinute8 = "" _
        And txtMinute9 = "" And txtMinute10 = "" And txtMinute11 = "" Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
        pSetFocus
        Exit Sub
    End If
    
'CLEAR------------------

    If cmdClearAnncTimes.Caption = "Clear Announce Times" Then
        
        If txtAnnc4 <> "" Or txtAnnc5 <> "" Or txtAnnc6 <> "" Or txtAnnc7 <> "" Or txtAnnc8 <> "" _
        Or txtAnnc9 <> "" Or txtAnnc10 <> "" Or txtAnnc11 <> "" Then
        
         If ck4Played.Value = 1 Then
             ck4Played.Value = 0
         End If
         
         If ck5Played.Value = 1 Then
             ck5Played.Value = 0
         End If
         
         If ck6Played.Value = 1 Then
             ck6Played.Value = 0
         End If
         
         If ck7Played.Value = 1 Then
             ck7Played.Value = 0
         End If
         
         If ck8Played.Value = 1 Then
             ck8Played.Value = 0
         End If
         
         If ck9Played.Value = 1 Then
             ck9Played.Value = 0
         End If
         
         If ck10Played.Value = 1 Then
             ck10Played.Value = 0
         End If
         
         If ck11Played.Value = 1 Then
             ck11Played.Value = 0
         End If
        
        
            Close #501
            Open "AnncTime.dat" For Output As #501
            Write #501, txtBackAnnc, txtAnnc4, txtAnnc5, txtAnnc6, txtAnnc7, txtAnnc8, txtAnnc9, txtAnnc10, txtAnnc11
            Close #501
             
            Check1(0).Value = 0
            Check1(1).Value = 0
            Check1(2).Value = 0
            Check1(3).Value = 0
            Check1(4).Value = 0
            Check1(5).Value = 0
            Check1(6).Value = 0
            Check1(7).Value = 0
            
            If txtBackAnnc.Enabled And txtBackAnnc <> "0" And txtBackAnnc <> "" Then
               txtBackAnnc = ""
            End If
            
            txtAnnc4.Text = ""
            txtAnnc5.Text = ""
            txtAnnc6.Text = ""
            txtAnnc7.Text = ""
            txtAnnc8.Text = ""
            txtAnnc9.Text = ""
            txtAnnc10.Text = ""
            txtAnnc11.Text = ""

         End If
         
        cmdClearAnncTimes.Caption = "Restore Annc Times"
        chkBackAnnc.Enabled = False
        
'RESTORE------------------------------

    ElseIf cmdClearAnncTimes.Caption = "Restore Annc Times" Then
    
        Close #501
        Open "AnncTime.dat" For Input As #501
            Input #501, BackAnnc, Annc4, Annc5, Annc6, Annc7, Annc8, Annc9, Annc10, Annc11
        Close #501
            
        If chkBackAnnc = 0 Then
            txtBackAnnc.Enabled = True
        End If

        txtBackAnnc = BackAnnc
        txtAnnc4 = Annc4
        txtAnnc5 = Annc5
        txtAnnc6 = Annc6
        txtAnnc7 = Annc7
        txtAnnc8 = Annc8
        txtAnnc9 = Annc9
        txtAnnc10 = Annc10
        txtAnnc11 = Annc11
    
        cmdClearAnncTimes.Caption = "Clear Announce Times"
        chkBackAnnc.Enabled = True
    End If
    pSetFocus
End Sub

Private Sub cmdClearBlock_CliCk()
    'clears time-block related text boxes
   Dim iBlock As Integer
    iBlock = iBlock + 1
    If iBlock = 1 Or iBlock = 3 Then
        txtBlock = ""
        lblRemain60 = ""
        txtBlock.SetFocus
    ElseIf iBlock = 2 Or iBlock = 4 Then
        pSetFocus
    End If
End Sub

Private Sub cmdClearLineup_Click()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
    
    On Error GoTo HandleErrors
    
     If ck4Played.Value = 1 Then
         ck4Played.Value = 0
     End If
     
     If ck5Played.Value = 1 Then
         ck5Played.Value = 0
     End If
     
     If ck6Played.Value = 1 Then
         ck6Played.Value = 0
     End If
     
     If ck7Played.Value = 1 Then
         ck7Played.Value = 0
     End If
     
     If ck8Played.Value = 1 Then
         ck8Played.Value = 0
     End If
     
     If ck9Played.Value = 1 Then
         ck9Played.Value = 0
     End If
     
     If ck10Played.Value = 1 Then
         ck10Played.Value = 0
     End If
     
     If ck11Played.Value = 1 Then
         ck11Played.Value = 0
     End If
    
          
    If txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
    Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "" Then
    
        
        Open "TimeRemain.dat" For Output As #500 'saves even if no data
        Write #500, txtComposer4, txtComposer5, txtComposer6, txtComposer7, txtComposer8, txtComposer9, txtComposer10, txtComposer11, _
            ; iMinute4, iMinute5, iMinute6, iMinute7, iMinute8, iMinute9, iMinute10, iMinute11, _
            iSecond4, iSecond5, iSecond6, iSecond7, iSecond8, iSecond9, iSecond10, iSecond11, _
            txtCD4, txtCD5, txtCD6, txtCD7, txtCD8, txtCD9, txtCD10, txtCD11, iAnncSum
        Close #500

        If txtAnnc4 <> "" Or txtAnnc5 <> "" Or txtAnnc6 <> "" Or txtAnnc7 <> "" Or txtAnnc8 <> "" _
        Or txtAnnc9 <> "" Or txtAnnc10 <> "" Or txtAnnc11 <> "" Then

            Close #501
            Open "AnncTime.dat" For Output As #501
            Write #501, txtBackAnnc, txtAnnc4, txtAnnc5, txtAnnc6, txtAnnc7, txtAnnc8, txtAnnc9, txtAnnc10, txtAnnc11
            Close #501
        End If

        ck4Played.Value = 0
        ck5Played.Value = 0
        ck6Played.Value = 0
        ck7Played.Value = 0
        ck8Played.Value = 0
        ck9Played.Value = 0
        ck10Played.Value = 0
        ck11Played.Value = 0
    
'    Check1(0).Value = 0
'    Check1(1).Value = 0
'    Check1(2).Value = 0
'    Check1(3).Value = 0
'    Check1(4).Value = 0
'    Check1(5).Value = 0
'    Check1(6).Value = 0
'    Check1(7).Value = 0
'
       mnuToolsExportLineCopy.Checked = False
       mnuToolsExportLineCopy.Enabled = True
       lblExport.Visible = False
    
       If txtComposer4 = "" Then
           'txtAnnc4 = ""
           txtMinute4 = ""
           txtSecond4 = ""
           txtCD4 = ""
       End If
    
        If txtComposer5 = "" Then
           ' txtAnnc5 = ""
            txtMinute5 = ""
            txtSecond5 = ""
            txtCD5 = ""
        End If
        
        If txtComposer6 = "" Then
           ' txtAnnc6 = ""
            txtMinute6 = ""
            txtSecond6 = ""
            txtCD6 = ""
        End If
        
        If txtComposer7 = "" Then
           ' txtAnnc7 = ""
            txtMinute7 = ""
            txtSecond7 = ""
            txtCD7 = ""
        End If
        
        If txtComposer8 = "" Then
           ' txtAnnc8 = ""
            txtMinute8 = ""
            txtSecond8 = ""
            txtCD8 = ""
        End If
        
        If txtComposer9 = "" Then
           ' txtAnnc9 = ""
            txtMinute9 = ""
            txtSecond9 = ""
            txtCD9 = ""
        End If
        
        If txtComposer10 = "" Then
           ' txtAnnc10 = ""
            txtMinute10 = ""
            txtSecond10 = ""
            txtCD10 = ""
        End If
        
        If txtComposer11 = "" Then
            'txtAnnc11 = ""
            txtMinute11 = ""
            txtSecond11 = ""
            txtCD11 = ""
        End If
    
        cmdRestoreEntries.BackColor = &H8000000F 'gray
        lblAnncTime.BorderStyle = 0
        
        cmdRestoreEntries.ToolTipText = ""
        
        If frmPlanner!mnuLinkTimeRemain.Checked = True Then
            lblLinked_DblClick
        End If
        
        If mnuToolsExportLineCopy.Checked = True Then
            mnuToolsExportLineCopy_Click
        End If
    
        If lblLinked.Visible = True Then
            lblLinked_DblClick
        End If
                
    End If
   
HandleErrors:

    Label32(0) = ""
    Label32(1) = ""
    Label32(2) = ""
    Label32(3) = ""
    Label32(4) = ""
    Label32(5) = ""
    Label32(6) = ""
    Label32(7) = ""
   
    txtMinute4 = ""
    txtSecond4 = ""
    txtMinute5 = ""
    txtSecond5 = ""
    txtMinute6 = ""
    txtSecond6 = ""
    txtMinute7 = ""
    txtSecond7 = ""
    txtMinute8 = ""
    txtSecond8 = ""
    txtMinute9 = ""
    txtSecond9 = ""
    txtMinute10 = ""
    txtSecond10 = ""
    txtMinute11 = ""
    txtSecond11 = ""
    
    txtComposer4 = ""
    txtComposer5 = ""
    txtComposer6 = ""
    txtComposer7 = ""
    txtComposer8 = ""
    txtComposer9 = ""
    txtComposer10 = ""
    txtComposer11 = ""
    
    txtCD4 = ""
    txtCD5 = ""
    txtCD6 = ""
    txtCD7 = ""
    txtCD8 = ""
    txtCD9 = ""
    txtCD10 = ""
    txtCD11 = ""
    
    iAnnc4 = ""
    iAnnc5 = ""
    iAnnc6 = ""
    iAnnc7 = ""
    iAnnc8 = ""
    iAnnc9 = ""
    iAnnc10 = ""
    iAnnc11 = ""
    
    txtAnnc4 = ""
    txtAnnc5 = ""
    txtAnnc6 = ""
    txtAnnc7 = ""
    txtAnnc8 = ""
    txtAnnc9 = ""
    txtAnnc10 = ""
    txtAnnc11 = ""
    
    AnncControl4 = 0
    AnncControl5 = 0
    AnncControl6 = 0
    AnncControl7 = 0
    AnncControl8 = 0
    AnncControl9 = 0
    AnncControl10 = 0
    AnncControl11 = 0
   
    txtComposer4.BackColor = &H80000005 ' white
    txtComposer5.BackColor = &H80000005 ' white
    txtComposer6.BackColor = &H80000005 ' white
    txtComposer7.BackColor = &H80000005 ' white
    txtComposer8.BackColor = &H80000005 ' white
    txtComposer9.BackColor = &H80000005 ' white
    txtComposer10.BackColor = &H80000005 ' white
    txtComposer11.BackColor = &H80000005 ' white
    mnuExport.Enabled = False

   ' txtSpotsS = ""
    lblAnncTime = " Annc Time"
    lblAnncTime.BorderStyle = 0
        
    miAnncTime = 0
    pAnnounce
    
    If chkAnnounce.Value = 0 And txtMinute2 = "" And txtSecond2 = "" Then
        txtMinute3.Visible = False
        txtSecond3.Visible = False
        lblS.Visible = False
        lblMinSec3Div.Visible = False
        shpTime3.Visible = False
        lblAnncMin.Visible = False
        lblAnncSec.Visible = False
        Label9.Visible = False
    End If
   
    imgMic.Visible = False
    cmdRestoreEntries.Enabled = True
    
    If mnuToolsExportLineCopy.Checked = True Then
        mnuToolsExportLineCopy_Click
    End If

    txtComposer4.Enabled = True
    txtComposer4.SetFocus
   
    If cMinAx >= 50 Then
        lblProgramInfo.Caption = " Including annc time what you have programmed will end at " & miTotal & ""
    Else
        lblProgramInfo.Caption = " What you have programmed will end at " & miTotal & " "
    End If
    
    If chkStopWatch.Value = 0 Then
       
        imgClock.Visible = True
        If cTotalRemain >= 3.16 Then
            imgMusic.Visible = True
        End If
    End If
    cmdRestoreEntries.Enabled = True
    cmdRestoreEntries.Caption = "&Restore Lineup"
    cmdRestoreEntries.ToolTipText = ""
    cmdClearAnncTimes.Caption = "Clear Announce Times"
   
End Sub

Private Sub cmdClearPads_Click()
    ck4Played.Value = 0
    ck5Played.Value = 0
    ck6Played.Value = 0
    ck7Played.Value = 0
    ck8Played.Value = 0
    ck9Played.Value = 0
    ck10Played.Value = 0
    ck11Played.Value = 0
    pSetFocus
    
End Sub

Private Sub cmdCloseTimeSet_Click()
    mnuOptionsTime.Checked = False
    fraIntro.Visible = False
    chkAnnounce.Enabled = True
    cmdDefaults.Enabled = True
    Label24.ForeColor = &H80& 'rust color
    Frame8.ForeColor = &H80& 'rust color
    Label26.ForeColor = &H80& 'rust color
    
    Frame8.BackColor = &H8000000F 'button face
    Label24.BackColor = &H8000000F
    Label26.BackColor = &H8000000F
End Sub

Private Sub cmdDefaults_Click()
On Error GoTo HandleErrors
    Dim iFileNumber As Integer
    iFileNumber = FreeFile 'assigns file # as last free number
 
    Dim sIntroOut, sClose, sSpot, sSignature As String

    Open "DefaultTimes.dat" For Input As #22
        Input #22, sIntroOut, sClose, sSpot
    Close #22
   
     If sIntroOut = "" Or sClose = "" Or sSpot = "" Then
        MsgBox "Default program times are not set", vbOKOnly, "Default Data Missing"
        giClockShow = 5
        cmdDefaults.Enabled = False
        giClockShow = 0
        Exit Sub
    End If
        
    txtBlock = "60.0"
    txtIntroSetting = sIntroOut
    txtCloseOut = sClose
    txtSpotLength = sSpot
   
    '---------selecting default is saved
    Open "Times.dat" For Output As #23
    Write #23, frmDefaults!txtPlanTime, txtIntroSetting, txtCloseOut, txtSpotLengthSetting
    Close #23
        
    cmdCloseTimeSet.SetFocus
    
    Exit Sub
HandleErrors:
    Close #22
    Close #23
End Sub

Private Sub cmdExit_Click()
    Unload frmMemos
    giClockShow = 5
    frmPlanner.Show
    frmTimeRemain.Hide
End Sub

Private Sub cmdMinMinus_Click()

    If txtMinAdj = "" Then
        txtMinAdj = -1
    ElseIf txtMinAdj <> "" And txtMinAdj > -58 Then
        mtMin = Val(txtMinAdj)
        mtMin = mtMin - 1
        txtMinAdj = mtMin
    End If
    
    If txtSecAdj <> "" Then
        If Val(txtMinAdj) > 0 And Val(txtSecAdj) < 0 Then
            txtSecAdj = 0
        ElseIf Val(txtMinAdj) < 0 And Val(txtSecAdj > 0) Then
             txtSecAdj = 0
        End If
    End If

End Sub

Private Sub cmdMinPlus_Click()
      If txtMinAdj = "" Then
        txtMinAdj = 1
    ElseIf txtMinAdj <> "-" And txtMinAdj < 59 Then
        mtMin = Val(txtMinAdj)
        mtMin = mtMin + 1
        txtMinAdj = mtMin
    End If
    
    If txtSecAdj <> "" Then
        If (Val(txtMinAdj) > 0 And Val(txtSecAdj) < 0) Then
            txtSecAdj = 0
        ElseIf Val(txtMinAdj) < 0 And Val(txtSecAdj > 0) Then
             txtSecAdj = 0
        End If
    End If
End Sub

Private Sub cmdPower_Click()
    giClockShow = 5
    frmTransmitter!cmdPrevious.Caption = "&Return to Previous Page F6"
    frmTransmitter.Show
    frmTimeRemain.Hide
End Sub

Private Sub cmdPrint_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    Dim dDate As Date
    Dim sDate As String
    dDate = Now
    sDate = Format(dDate, "Long Date")
   
    If txtComposer4 = "" And txtComposer5 = "" And txtComposer6 = "" And txtComposer7 = "" _
    And txtComposer8 = "" And txtComposer9 = "" And txtComposer10 = "" And txtComposer11 = "" Then
        MsgBox "There is no Music Lineup to Print", vbOKOnly, "No Data"
        pSetFocus
        Exit Sub
    End If
   
    iResponse = MsgBox("Print a copy of the Music Lineup?", vbYesNo, "Print Lineup")
    If iResponse = vbNo Then
        pSetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then

        Printer.FontName = "Arial"
        Printer.FontSize = 12
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.Print Tab(7); "Music Planning Draft"
        Printer.FontBold = False
        Printer.FontSize = 9
        Printer.Print Tab(11); sDate
        Printer.Print
        Printer.Print
        Printer.FontSize = 11
        
        '---------------
        Dim sCD4 As String
        If txtCD4 <> "" Then
            sCD4 = " CD# " & txtCD4
        Else
            sCD4 = ""
        End If
 
        If txtComposer4 <> "" And ck4Played.Value = 0 Then
            Printer.Print Tab(8); txtComposer4 & "   •  " & sCD4 & "    (" & txtMinute4 & ":" & txtSecond4 & ")"
            Printer.Print
            Printer.Print
        End If
        
        '-----------
        Dim sCD5 As String
        If txtCD5 <> "" Then
            sCD5 = " CD# " & txtCD5
        Else
            sCD5 = ""
        End If
        
        If txtComposer5 <> "" And ck5Played.Value = 0 Then
            Printer.Print Tab(8); txtComposer5 & "   •  " & sCD5 & "    (" & txtMinute5 & ":" & txtSecond5 & ")"
            Printer.Print
            Printer.Print
        End If
        
        '-----------
        Dim sCD6 As String
        If txtCD6 <> "" Then
            sCD6 = " CD# " & txtCD6
        Else
            sCD6 = ""
        End If

        If txtComposer6 <> "" And ck6Played.Value = 0 Then
            Printer.Print Tab(8); txtComposer6 & "   •  " & sCD6 & "    (" & txtMinute6 & ":" & txtSecond6 & ")"
            Printer.Print
            Printer.Print
        End If
        
        '-----------
        Dim sCD7 As String
        If txtCD7 <> "" Then
            sCD7 = " CD# " & txtCD7
        Else
            sCD7 = ""
        End If
        
        If txtComposer7 <> "" And ck7Played.Value = 0 Then
            Printer.Print Tab(8); txtComposer7 & "   •  " & sCD7 & "    (" & txtMinute7 & ":" & txtSecond7 & ")"
            Printer.Print
            Printer.Print
        End If
        
        '-----------
        Dim sCD8 As String
        If txtCD8 <> "" Then
            sCD8 = " CD# " & txtCD8
        Else
            sCD8 = ""
        End If
        
        If txtComposer8 <> "" And ck8Played.Value = 0 Then
            Printer.Print Tab(8); txtComposer8 & "   •  " & sCD8 & "    (" & txtMinute8 & ":" & txtSecond8 & ")"
            Printer.Print
            Printer.Print
        End If
        
        '-----------
        Dim sCD9 As String
        If txtCD9 <> "" Then
            sCD9 = " CD# " & txtCD9
        Else
            sCD9 = ""
        End If
        
        If txtComposer9 <> "" And ck9Played.Value = 0 Then
            Printer.Print Tab(8); txtComposer9 & "   •  " & sCD9 & "    (" & txtMinute9 & ":" & txtSecond9 & ")"
            Printer.Print
            Printer.Print
        End If
        
        '-----------
        Dim sCD10 As String
        If txtCD10 <> "" Then
            sCD10 = " CD# " & txtCD10
        Else
            sCD10 = ""
        End If
        
        If txtComposer10 <> "" And ck10Played.Value = 0 Then
            Printer.Print Tab(8); txtComposer10 & "   •  " & sCD10 & "    (" & txtMinute10 & ":" & txtSecond10 & ")"
            Printer.Print
            Printer.Print
        End If
        
        '-----------
        Dim sCD11 As String
        If txtCD11 <> "" Then
            sCD11 = " CD# " & txtCD11
        Else
            sCD11 = ""
        End If
        
        If txtComposer11 <> "" And ck11Played.Value = 0 Then
            Printer.Print Tab(8); txtComposer11 & "   •  " & sCD11 & "    (" & txtMinute11 & ":" & txtSecond11 & ")"
            Printer.Print
            Printer.Print
        End If
        Printer.FontSize = 10
        Printer.Print Tab(45); "———— Lineup:  " & lblTotalS & " ————"
        Printer.EndDoc
        
    End If
    pSetFocus
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"

End Sub

Private Sub cmdReadMe_Click()
    frmNote.Show vbModal
End Sub

Private Sub cmdRestoreTimes_Click()

    Dim iMinute1 As String
    Dim iMinute2 As String
    Dim iSecond1 As String
    Dim iSecond2 As String
    Dim iiHour As String
    
On Error GoTo HandleErrors
    
    Open "RunTime.dat" For Input As #502
    Input #502, iMinute1, iMinute2, iSecond1, iSecond2, iiHour
    Close #502
    lblHour.Visible = True
    txtMinute1.Text = iMinute1
    txtSecond1.Text = iSecond1
    txtMinute2.Text = iMinute2
    txtSecond2.Text = iSecond2
    lblHour = iiHour
 
    imgHand.Visible = False
    Label6.ForeColor = &H80000008  'black
    Label12.ForeColor = &H80000008 'black
    
    imgClock.Visible = True
    
    If txtMinute1 < "55" Then
        lblCurrentTime.ForeColor = vbBlack
    Else
        lblCurrentTime.ForeColor = &HFF0000    'blue
        lblCurrentTime.Alignment = 1
        lblCurrentTime.Caption = "Approaching the end of the hour"
    End If
    
    If txtMinute1 >= "57" And txtMinute1 <= "59" Then
        Dim iResponse As Integer
        iResponse = MsgBox("The restored time is " & txtMinute1 & " minutes and " & txtSecond1 & _
        " seconds past the hour. If you are beginning the next hour of programming early, from this time, click 'Yes'." _
        & vbCrLf & vbCrLf & "Click 'No' to continue within the current hour." _
        , vbYesNo, "Begin new hour?")
        
        If iResponse = vbYes Then
            lblCurrentTime_DblClick
        End If
   End If
    
    Frame7.Caption = "Set Current Time"
    If chkStopWatch.Value = 0 Then
    
        If txtMinAdj = "" And txtSecAdj = "" Then
            cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
        Else
             cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
        End If
        
    Else
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Stopwatch) F5"
    End If
    cmdSetTime.Visible = False
    imgMusic.Visible = True
'---------
    If chkBackAnnc.Value = 1 Then
        chkBackAnnc.Value = 0
    End If
    
    If txtSpotsS <> "" And txtSpotsS <> "0" And txtSecond3.Visible = True Then
        lblS.Visible = True
        lblSspot.Visible = True
        Label31.Visible = True
        If txtSpotsS = "1" Then
            lblS.Caption = txtSpotsS & " spot"
            Label31.Caption = txtSpotsS & " spot"
        Else
            lblS.Caption = txtSpotsS & " spots"
            Label31.Caption = txtSpotsS & " spots"
        End If
       
    Else
        lblS.Visible = False
        lblSspot.Visible = False
        Label31.Visible = False
        lblS.Caption = ""
        Label31.Caption = ""

    End If
    
    Shape1.Visible = True
    
    If lblHour <> "" Then
        Shape1.Width = 1140 'expanded
        Shape1.Left = 5440
    Else
        Shape1.Width = 840 'normal
        Shape1.Left = 5715
    End If
    
    If chkAnnounce.Value = 0 And txtMinute4 <> "" And txtMinute3 = "" And txtSecond3 = "" Then
        txtBackAnnc.Visible = True
        chkBackAnnc.Visible = True
        Label21.Visible = True
        
        If Val(txtIntro) > 29 Then
            txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
        Else
            txtBackAnnc = "0"
        End If
    End If

    pAnnounce
    pSetFocus
    Exit Sub
    
HandleErrors:
    Close #502
  
End Sub

Private Sub cmdSave_Click()
On Error GoTo HandleErrors

    Open "Times.dat" For Output As #23
    Write #23, frmDefaults!txtPlanTime, txtIntroSetting, txtCloseOut, txtSpotLengthSetting
    Close #23
    
    If IsNumeric(txtIntroSetting) Then
        frmDefaults!txtIntroOut = txtIntroSetting
    End If
    
     If IsNumeric(txtCloseOut) Then
        frmDefaults!txtClose = txtCloseOut
     End If
     
    If IsNumeric(txtSpotLengthSetting) Then
       frmDefaults!txtSpot = txtSpotLengthSetting
    End If

HandleErrors:
    cmdCloseTimeSet.SetFocus
End Sub

Private Sub cmdSecMinus_Click()
  
    If txtSecAdj = "" Then
        txtSecAdj = -5
    ElseIf txtSecAdj <> "-" And txtSecAdj > -54 Then
        mtSec = Val(txtSecAdj)
        mtSec = mtSec - 5
        txtSecAdj = mtSec
    End If
    
    If txtMinAdj <> "" Then
        If Val(txtSecAdj) > 0 And Val(txtMinAdj) < 0 Then
           txtMinAdj = 0
        ElseIf Val(txtSecAdj) < 0 And Val(txtMinAdj) > 0 Then
            txtMinAdj = 0
        End If
    End If
 
End Sub

Private Sub cmdSecPlus_Click()

    If txtSecAdj = "" Then
        txtSecAdj = 5
    ElseIf txtSecAdj <> "-" And txtSecAdj > -60 Then
        mtSec = Val(txtSecAdj)
        mtSec = mtSec + 5
        txtSecAdj = mtSec
    End If

    If txtMinAdj <> "" Then
        If Val(txtSecAdj) > 0 And Val(txtMinAdj) < 0 Then
            txtMinAdj = 0
        ElseIf Val(txtSecAdj) < 0 And Val(txtMinAdj) > 0 Then
             txtMinAdj = 0
        End If
    End If
End Sub

Private Sub cmdSetTime_Click()

    txtMinute2.SetFocus
    If chkStopWatch.Value = 1 Then
        txtMinute1 = Minute(mRunTime)
        txtSecond1 = Format(mRunTime, "ss")

        txtMinute2 = ""
        txtSecond2 = ""
        cmdSetTime.Visible = False
        txtMinute2.SetFocus
        Label6.BackColor = &H80000018 'yellow
        Label6.BorderStyle = 1
        Exit Sub
    Else
    End If
 '-----------------
    Dim tMinBefore As Integer
    Dim tSecBefore As Integer
    Dim tHour As Integer
    Dim iMinute1 As Integer
    Dim iSecond1 As Integer
    
    iMinute1 = Minute(Time) + Val(txtMinAdj)
  
'----test 5-24-2015
    If iMinute1 >= 60 Then
        iMinute1 = iMinute1 - 60
    End If
'----------

    iiMinute1 = Minute(Time) + Val(txtMinAdj)
    iSecond1 = Second(Time) + Val(txtSecAdj)

'test, also go to keyword TimeCorrectionTest for remainder of test comment out

'TEST 1=================4/19/2015

   iHourNow = Hour(Now)
      
    If cmdAdjTime.Visible = True Then
        
        If Hour(Time) <= 12 Then
            tHour = Hour(Time)
    
        ElseIf Hour(Time) > 12 Then
            tHour = (Hour(Time) - 12)
        End If
        
    Else
    
        If iHourNow <= 12 Then
            iHourNow = iHourNow
        ElseIf iHourNow > 12 Then
            iHourNow = iHourNow - 12
        End If
        
        tHour = iHourNow
    
    End If
'end test======4/19/2015
   
    '-----------second1
    
    If iSecond1 > 59 Then
        iSecond1 = (iSecond1 - 60)
        iMinute1 = (iMinute1 + 1)
    ElseIf iSecond1 < 0 Then
        iSecond1 = (iSecond1 + 60)
        iMinute1 = (iMinute1 - 1)
    End If
    
'TEST 2======+++++++++==========4/19/2015

    '---------minute1
    
    If cmdAdjTime.Visible = True Then
     
        If iMinute1 > 59 Then
            iMinute1 = 0
        ElseIf iMinute1 < 0 Then
            iMinute1 = iMinute1 + 60
        End If
     '-----------
    ElseIf cmdAdjTime.Visible = False Then
    
        If iMinute1 > 59 Then
            iMinute1 = 0
        ElseIf iMinute1 < 0 Then
            iMinute1 = iMinute1 + 60
        End If
        
     End If
     
 '---------
    txtSecond1.Text = Format$(iSecond1, "00")
    txtMinute1.Text = Format$(iMinute1, "00")
    
'end test==========4/19/2015

    tSecBefore = (60 - iSecond1)
    tMinBefore = (59 - iMinute1)
    
    lblHour.Visible = True 'these must be after txtMinute1 & txtSecond1 code has run
    Shape1.Visible = False 'border around txtMinute1 & txtSecond1
'----------
    
    Frame7.Caption = "Set Current Time"
    
    If txtMinAdj = "" And txtSecAdj = "" Then
            cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
        Else
             cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
        End If
    
    cmdSetTime.Visible = False
    
    txtMinute2 = ""
    txtSecond2 = ""
    txtBackAnnc = ""
    Dim iResponse As Integer
     
    If txtMinute1 >= "57" And txtMinute1 < "58" Then
     
        iResponse = MsgBox("The current time is " & txtMinute1 & " minutes and " & txtSecond1 & " seconds past the hour." _
        & vbCrLf & vbCrLf & "If you are beginning of the next hour of programming early at " _
        & txtMinute1 & ":" & txtSecond1 & " click yes." & vbCrLf & vbCrLf & _
        "Otherwise to continue the current program within the current hour click 'No'.", vbYesNo + vbQuestion, _
        "Begin the next hour of programming early or continue within the current hour?")
    
         If iResponse = vbYes Then
            lblCurrentTime_DblClick
        Else
            lblCurrentTime.ForeColor = vbBlack
            lblCurrentTime.Caption = " 1. Click the 'Set Current Time' button to enter the current time:"
        End If
        miCurrentHour = 1
        
    ElseIf txtMinute1 >= "58" And txtMinute1 <= "59" And miCurrentHour = 0 Then

        iResponse = MsgBox("The time is " & tHour & ":" & txtMinute1 & ":" & txtSecond1 & vbCrLf & vbCrLf & _
        "Are you beginning the next hour of programming early before the end of the current hour?", _
        vbYesNo + vbDefaultButton1, "Begin the next hour of programming early?")

        If iResponse = vbYes Then
            lblCurrentTime_DblClick
        Else
            lblCurrentTime.ForeColor = vbBlack
            lblCurrentTime.Caption = " 1. Click the 'Set Current Time' button to enter the current time:"
       End If
    Else
        miCurrentHour = 0
    End If
  
    Label6.BackColor = &H80000018 'yellow
    Label6.BorderStyle = 1
    imgDisc.Visible = True
    
    If lblLinked.Visible = True Then
        lblLinked_DblClick
    End If
    
On Error GoTo HandleErrors
    
    If cmdAdjTime.Visible = False Then 'command button
        Dim ZZ As Integer
        ZZ = Val(lblHourAdj)
        lblHour.Caption = ZZ & ":"
    ElseIf cmdAdjTime.Visible = True Then
        lblHour.Caption = iHourNow & ":"
    End If
    
    If txtMinute1 <> "" Or txtSecond1 <> "" Then
        Open "RunTime.dat" For Output As #502
        Write #502, txtMinute1, txtMinute2, txtSecond1, txtSecond2, lblHour
        Close #502
    End If
    
     Shape1.Visible = True
    
    If lblHour <> "" Then
        Shape1.Width = 1140 'expanded
        Shape1.Left = 5440
    Else
        Shape1.Width = 840 'normal
        Shape1.Left = 5715
    End If


HandleErrors:
    Close #502

End Sub

Private Sub cmdSetTime_GotFocus()
    cmdSetTime.BackColor = &H80000018 'yellow
    cmdSetTime.Caption = " Note the time remaining on current CD then click space bar or Enter key"
End Sub

Private Sub cmdSetTime_LostFocus()
    cmdSetTime.BackColor = &HFFFFFF 'white
    cmdSetTime.Caption = " Note the time remaining on current CD then click Enter key or CANCEL"
End Sub

Private Sub cmdClearTimes_Click()

On Error GoTo HandleErrors

    If txtMinute2 <> "" Or txtSecond2 <> "" Then
    
    
        If txtMinute1 <> "" Or txtMinute2 <> "" Or txtSecond1 <> "" Or txtSecond2 <> "" Then
            Open "RunTime.dat" For Output As #502
            Write #502, txtMinute1, txtMinute2, txtSecond1, txtSecond2, lblHour
            Close #502
        End If
            txtMinute2 = ""
            txtSecond2 = ""
            
            If txtAnnc4 <> "" Then
                txtAnnc4 = txtIntro
            End If
          
            txtMinute2.SetFocus
    Else
        txtMinute1 = ""
        txtSecond1 = ""
        lblHour = ""
        lblHour.Visible = False

        Shape1.Width = 840 'normal
        Shape1.Left = 5715

        iCurrentTime = 0
        lblCurrentTime.Alignment = 0 'left
        lblCurrentTime.ForeColor = vbBlack
        lblCurrentTime.Caption = " 1. Click the 'Set Current Time' button to enter the current time:"
        txtBackAnnc = ""
        If lblRemain30.Visible = True Then
            lblSpots.Caption = "Enter the number of (" & Val(txtSpotLength) & _
            "-second average time) spot, promo, PSA, weather, etc. inserts scheduled in the hour (or half-hour) time period"
            
        Else
            lblSpots.Caption = "Enter the number of (" & Val(txtSpotLength) & _
            "-second average time) spot, promo, PSA, weather, etc. inserts scheduled in the hour time period"
        End If
        
        mCurrentTime = 0
        lblCurrentTime.ForeColor = vbBlack
        frmPlanner!staStatus.Panels(2) = ""
        pSetFocus
    End If
    
    If chkBackAnnc.Value = 1 Then
        chkBackAnnc.Value = 0
    End If
    Label12.Visible = False
    imgOnAirSign.Visible = False
    imgDisc.Visible = False
    lblEndTime.Visible = False
    Label13.Visible = False
    imgMic.Visible = False
    imgClock.Visible = False
   '------
    If frmPlanner!mnuLinkTimeRemain.Checked = True Then
        shpLink.Visible = True
    End If
   '---------
   
    If txtSpotsS <> "" And txtSpotsS <> "0" And txtSecond3.Visible = True Then
        lblS.Visible = True
        lblSspot.Visible = True
        Label31.Visible = True
        If txtSpotsS = "1" Then
            lblS.Caption = txtSpotsS & " spot"
            Label31.Caption = txtSpotsS & " spot"
        Else
            lblS.Caption = txtSpotsS & " spots"
            Label31.Caption = txtSpotsS & " spots"
        End If

    Else
       lblS.Visible = False
       lblSspot.Visible = False
       Label31.Visible = False
       lblS.Caption = ""
       Label31.Caption = ""
    End If
    
    If txtMinute1 = "" And txtSecond1 = "" Then
        imgHand.Visible = True
    End If
    
    Shape1.Visible = True
    
    pAnnounce
    
    Exit Sub
HandleErrors:
    Close #502
End Sub

Private Sub cmdRestoreEntries_Click()

pSetFocus

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
    
    mnuToolsExportLineCopy.Checked = False
    
    If cmdRestoreEntries.Caption = "&Restore Lineup" Then
        
        Dim CountEntries As Integer
        Dim anncSum As String
        
            ck4Played.Value = 0
            ck5Played.Value = 0
            ck6Played.Value = 0
            ck7Played.Value = 0
            ck8Played.Value = 0
            ck9Played.Value = 0
            ck10Played.Value = 0
            ck11Played.Value = 0
            
            txtComposer4.BackColor = &H80000005 ' white
            txtComposer5.BackColor = &H80000005 ' white
            txtComposer6.BackColor = &H80000005 ' white
            txtComposer7.BackColor = &H80000005 ' white
            txtComposer8.BackColor = &H80000005 ' white
            txtComposer9.BackColor = &H80000005 ' white
            txtComposer10.BackColor = &H80000005 ' white
            txtComposer11.BackColor = &H80000005 ' white
                        
        cmdRestoreEntries.BackColor = &H8000000F 'gray
        
        If frmPlanner!mnuLinkTimeRemain.Checked = False Then 'NOT linked
            cmdRestoreEntries.Enabled = False
    
            RecallControl = 1
On Error GoTo HandleErrors
    
            Close #500
            Open "TimeRemain.dat" For Input As #500
                Input #500, Composer4, Composer5, Composer6, Composer7, Composer8, Composer9, Composer10, Composer11, _
                Minute4, Minute5, Minute6, Minute7, Minute8, Minute9, Minute10, Minute11, _
                Second4, Second5, Second6, Second7, Second8, Second9, Second10, Second11, _
              CD4, CD5, CD6, CD7, CD8, CD9, CD10, CD11, anncSum
            Close #500
                    
            Dim iAncTime As Integer
            iAncTime = Val(Annc4) + Val(Annc5) + Val(Annc6) + Val(Annc7) + Val(Annc8) + Val(Annc9) + Val(Annc10) + Val(Annc11)
            mRAnncSum = iAncTime
    
            If lblLinked.Visible = True Then
                lblLinked_DblClick
            End If
    
            txtMinute4 = Minute4
            txtSecond4 = Second4
            txtSecond4.Text = Format$(txtSecond4, "00")
            txtCD4 = CD4
    
            txtMinute5 = Minute5
            txtSecond5 = Second5
            txtSecond5 = Format$(txtSecond5, "00")
            txtCD5 = CD5
    
            txtMinute6 = Minute6
            txtSecond6 = Second6
            txtSecond6 = Format$(txtSecond6, "00")
            txtCD6 = CD6
    
            txtMinute7 = Minute7
            txtSecond7 = Second7
            txtSecond7 = Format$(txtSecond7, "00")
            txtCD7 = CD7
    
            txtMinute8 = Minute8
            txtSecond8 = Second8
            txtSecond8 = Format$(txtSecond8, "00")
            txtCD8 = CD8
    
            txtMinute9 = Minute9
            txtSecond9 = Second9
            txtSecond9 = Format$(txtSecond9, "00")
            txtCD9 = CD9
    
            txtMinute10 = Minute10
            txtSecond10 = Second10
            txtSecond10 = Format$(txtSecond10, "00")
            txtCD10 = CD10
    
            txtMinute11 = Minute11
            txtSecond11 = Second11
            txtSecond11 = Format$(txtSecond11, "00")
            txtCD11 = CD11
            
            If Len(Composer4) < 3 And Composer4 <> "" Then
                txtComposer4 = Composer4 & "--"
            Else
                txtComposer4 = Composer4
            End If
            
            If Len(Composer5) < 3 And Composer5 <> "" Then
                txtComposer5 = Composer5 & "--"
            Else
                txtComposer5 = Composer5
            End If
            
            If Len(Composer6) < 3 And Composer6 <> "" Then
                txtComposer6 = Composer6 & "--"
            Else
                txtComposer6 = Composer6
            End If
            
            If Len(Composer7) < 3 And Composer7 <> "" Then
                txtComposer7 = Composer7 & "--"
            Else
                txtComposer7 = Composer7
            End If
            
            If Len(Composer8) < 3 And Composer8 <> "" Then
                txtComposer8 = Composer8 & "--"
            Else
                txtComposer8 = Composer8
            End If
            
            If Len(Composer9) < 3 And Composer9 <> "" Then
                txtComposer9 = Composer9 & "--"
            Else
                txtComposer9 = Composer9
            End If
            
            If Len(Composer10) < 3 And Composer10 <> "" Then
                txtComposer10 = Composer10 & "--"
            Else
                txtComposer10 = Composer10
            End If
            
            If Len(Composer11) < 3 And Composer11 <> "" Then
                txtComposer11 = Composer11 & "--"
            Else
                txtComposer11 = Composer11
            End If
    
            If txtBlock <> "" Then
               pAnnounce
            End If
    
            If mnuToolsExportLineCopy.Checked = True Then
                mnuToolsExportLineCopy_Click
            End If

            If txtComposer4 = "" Then
                txtComposer4.SetFocus
            ElseIf txtComposer5 = "" Then
                txtComposer5.SetFocus
            ElseIf txtComposer6 = "" Then
                txtComposer6.SetFocus
            ElseIf txtComposer7 = "" Then
                txtComposer7.SetFocus
            ElseIf txtComposer8 = "" Then
                txtComposer8.SetFocus
            ElseIf txtComposer9 = "" Then
                txtComposer9.SetFocus
            ElseIf txtComposer10 = "" Then
                txtComposer10.SetFocus
            ElseIf txtComposer11 = "" Then
                txtComposer11.SetFocus
            End If
    
        ElseIf frmPlanner!mnuLinkTimeRemain.Checked = True Then 'Linked
    
         If frmPlanner!txtComposer1 <> "" Then
                txtComposer4 = frmPlanner!txtComposer1
                txtMinute4 = frmPlanner!txtMinute1
                txtSecond4 = frmPlanner!txtSecond1
                frmPlanner!txtAnnc1 = txtAnnc4
            Else
                txtComposer4 = ""
                txtMinute4 = ""
                txtSecond4 = ""
                frmPlanner!txtAnnc1 = ""
            End If
    
            If frmPlanner!txtComposer2 <> "" Then
                txtComposer5 = frmPlanner!txtComposer2
                txtMinute5 = frmPlanner!txtMinute2
                txtSecond5 = frmPlanner!txtSecond2
                frmPlanner!txtAnnc2 = txtAnnc5
            Else
                txtComposer5 = ""
                txtMinute5 = ""
                txtSecond5 = ""
                frmPlanner!txtAnnc2 = ""
            End If
    
            If frmPlanner!txtComposer3 <> "" Then
                txtComposer6 = frmPlanner!txtComposer3
                txtMinute6 = frmPlanner!txtMinute3
                txtSecond6 = frmPlanner!txtSecond3
                frmPlanner!txtAnnc3 = txtAnnc6
            Else
                txtComposer6 = ""
                txtMinute6 = ""
                txtSecond6 = ""
                frmPlanner!txtAnnc3 = ""
            End If
    
            If frmPlanner!txtComposer4 <> "" Then
                txtComposer7 = frmPlanner!txtComposer4
                txtMinute7 = frmPlanner!txtMinute4
                txtSecond7 = frmPlanner!txtSecond4
                frmPlanner!txtAnnc4 = txtAnnc7
            Else
                txtComposer7 = ""
                txtMinute7 = ""
                txtSecond7 = ""
                frmPlanner!txtAnnc4 = ""
            End If
    
            If frmPlanner!txtComposer5 <> "" Then
                txtComposer8 = frmPlanner!txtComposer5
                txtMinute8 = frmPlanner!txtMinute5
                txtSecond8 = frmPlanner!txtSecond5
                frmPlanner!txtAnnc5 = txtAnnc8
            Else
                txtComposer8 = ""
                txtMinute8 = ""
                txtSecond8 = ""
                frmPlanner!txtAnnc5 = ""
            End If
    
            If frmPlanner!txtComposer6 <> "" Then
                txtComposer9 = frmPlanner!txtComposer6
                txtMinute9 = frmPlanner!txtMinute6
                txtSecond9 = frmPlanner!txtSecond6
                frmPlanner!txtAnnc6 = txtAnnc9
            Else
                txtComposer9 = ""
                txtMinute9 = ""
                txtSecond9 = ""
                frmPlanner!txtAnnc6 = ""
            End If
    
            If frmPlanner!txtComposer7 <> "" Then
                txtComposer10 = frmPlanner!txtComposer7
                txtMinute10 = frmPlanner!txtMinute7
                txtSecond10 = frmPlanner!txtSecond7
                frmPlanner!txtAnnc7 = txtAnnc11
            Else
                txtComposer10 = ""
                txtMinute10 = ""
                txtSecond10 = ""
                frmPlanner!txtAnnc7 = ""
            End If
    
            If frmPlanner!txtComposer8 <> "" Or frmPlanner!txtComposition8 <> "" Then
    
                If frmPlanner!chkNonCD8.Value = 0 Then
                     txtComposer11 = frmPlanner!txtComposer8
                Else
                    txtComposer11 = "Non-Music Entry"
                End If
    
                txtMinute11 = frmPlanner!txtMinute8
                txtSecond11 = frmPlanner!txtSecond8
                frmPlanner!txtAnnc8 = txtAnnc11
            Else
                txtComposer11 = ""
                txtMinute11 = ""
                txtSecond11 = ""
                frmPlanner!txtAnnc8 = ""
            End If
    
            '----------
            frmPlanner!lblProgramTime.Visible = True
            frmPlanner!lblProgramTime = lblTotalS & " " & sPlannerProgramTime & "  •  " & sRemain
        Else
            '------------
    
            frmTimeRemain!cmdRestoreEntries.Enabled = False
            '-----
            If txtMinute2 = "" And txtSecond2 = "" Then
                Frame1.Caption = "Music Lineup"
                shpLink.Visible = True
            ElseIf txtMinute2 <> "" Then
                Frame1.Caption = "Additional Music Lineup"
            End If
            '------
    
            shpLink.BackColor = &HC000&
            shpLink.BorderColor = &HC000&
            giTimesDiffer = 0
            pSetFocus
        End If
        
        Close #501
        Open "AnncTime.dat" For Input As #501
        Input #501, BackAnnc, Annc4, Annc5, Annc6, Annc7, Annc8, Annc9, Annc10, Annc11
        Close #501
            
        txtBackAnnc = BackAnnc
        txtAnnc4 = Annc4
        txtAnnc5 = Annc5
        txtAnnc6 = Annc6
        txtAnnc7 = Annc7
        txtAnnc8 = Annc8
        txtAnnc9 = Annc9
        txtAnnc10 = Annc10
        txtAnnc11 = Annc11
        
HandleErrors:
'----------------------------------------------------------------------
    ElseIf cmdRestoreEntries.Caption = "&Remove Trial Entries" Then
    
         If Len(txtComposer4) < 3 Then
            txtComposer4.BackColor = &H80000005 ' white
            txtMinute4.BackColor = &H80000005  'white
            txtSecond4.BackColor = &H80000005  'white
            ck4Played.Value = 0
            txtComposer4 = ""
            txtMinute4 = ""
            txtSecond4 = ""
            txtAnnc4 = ""
        End If
        
        If Len(txtComposer5) < 3 Then
            txtComposer5.BackColor = &H80000005 ' white
            txtMinute5.BackColor = &H80000005  'white
            txtSecond5.BackColor = &H80000005  'white
            ck5Played.Value = 0
            txtComposer5 = ""
            txtMinute5 = ""
            txtSecond5 = ""
            txtAnnc5 = ""
        End If
        
        If Len(txtComposer6) < 3 Then
            txtComposer6.BackColor = &H80000005 ' white
            txtMinute6.BackColor = &H80000005  'white
            txtSecond6.BackColor = &H80000005  'white
            ck6Played.Value = 0
            txtComposer6 = ""
            txtMinute6 = ""
            txtSecond6 = ""
            txtAnnc6 = ""
        End If
        
        If Len(txtComposer7) < 3 Then
            txtComposer7.BackColor = &H80000005 ' white
            txtMinute7.BackColor = &H80000005  'white
            txtSecond7.BackColor = &H80000005  'white
            ck7Played.Value = 0
            txtComposer7 = ""
            txtMinute7 = ""
            txtSecond7 = ""
            txtAnnc7 = ""
        End If
        
        If Len(txtComposer8) < 3 Then
            txtComposer8.BackColor = &H80000005 ' white
            txtMinute8.BackColor = &H80000005  'white
            txtSecond8.BackColor = &H80000005  'white
            ck8Played.Value = 0
            txtComposer8 = ""
            txtMinute8 = ""
            txtSecond8 = ""
            txtAnnc8 = ""
        End If
        
        If Len(txtComposer9) < 3 Then
            txtComposer9.BackColor = &H80000005 ' white
            txtMinute9.BackColor = &H80000005  'white
            txtSecond9.BackColor = &H80000005  'white
            ck9Played.Value = 0
            txtComposer9 = ""
            txtMinute9 = ""
            txtSecond9 = ""
            txtAnnc9 = ""
        End If
                
        If Len(txtComposer10) < 3 Then
            txtComposer10.BackColor = &H80000005 ' white
            txtMinute10.BackColor = &H80000005  'white
            txtSecond10.BackColor = &H80000005  'white
            ck10Played.Value = 0
            txtComposer10 = ""
            txtMinute10 = ""
            txtSecond10 = ""
            txtAnnc10 = ""
        End If
        
        If Len(txtComposer11) < 3 Then
            txtComposer11.BackColor = &H80000005 ' white
            txtMinute11.BackColor = &H80000005  'white
            txtSecond11.BackColor = &H80000005  'white
            ck11Played.Value = 0
            txtComposer11 = ""
            txtMinute11 = ""
            txtSecond11 = ""
            txtAnnc11 = ""
        End If
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdClearAnncTimes.Caption = "Clear Announce Times"
        cmdRestoreEntries.ToolTipText = ""
        
        If txtSpotsS <> "" And txtSpotsS <> "0" And txtSecond3.Visible = True Then
            lblS.Visible = True
            lblSspot.Visible = True
            Label31.Visible = True
            If txtSpotsS = "1" Then
                lblS.Caption = txtSpotsS & " spot"
                Label31.Caption = txtSpotsS & " spot"
            Else
                lblS.Caption = txtSpotsS & " spots"
                Label31.Caption = txtSpotsS & " spots"
            End If

        Else
            lblS.Visible = False
            lblSspot.Visible = False
            Label31.Visible = False
            lblS.Caption = ""
            Label31.Caption = ""
        End If
       
    End If
  
    cmdRestoreEntries.Enabled = False
End Sub

Private Sub cmdPlanner_Click()
    giClockShow = 5
    frmPlanner.Show
    frmTimeRemain.Hide
End Sub

Private Sub cmdStopwatch_Click()
    If frmStopWatch.Visible = False Then
        frmStopWatch.Show
        cmdStopwatch.Enabled = False
    End If
End Sub

Private Sub cmdSystemTime_Click()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
     
    If chkStopWatch.Value = 1 And ((IsNull(mRunTime) = True Or mRunTime = 0 Or mRunTime = "")) Then
        MsgBox "To be the source for the 'Current Time' setting, the StopWatch timer must me running." & vbCrLf & "Click F12 to access the Stopwatch and start the timer.", vbOKOnly + vbExclamation, "StopWatch Timer is Not Running"
        cmdSetTime.Visible = False
        Exit Sub
    End If
    
    txtMinute1 = ""
    txtSecond1 = ""
    txtMinute2 = ""
    txtSecond2 = ""
    
    Shape1.Width = 840 'normal
    Shape1.Left = 5715
    

    iCurrentTime = 1
    
    If chkBackAnnc.Value = 1 Then
        chkBackAnnc.Value = 0
    End If
    
    imgHand.Visible = False
    imgDisc.Visible = False
    lblHour.Visible = False

    If cmdSetTime.Visible = False Then
        
        cmdSetTime.Visible = True
        cmdSystemTime.Caption = "CANCEL Setting Current Time Past the &Hour  F5"
        Frame7.Caption = "CANCEL"
        cmdSetTime.SetFocus
    
    ElseIf cmdSetTime.Visible = True Then
        cmdSetTime.Visible = False
        Frame7.Caption = "Set Current Time"
        If chkStopWatch.Value = 0 Then

            If txtMinAdj = "" And txtSecAdj = "" Then
                cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
            Else
                cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
            End If
            
        Else
            cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Stopwatch) F5"
        End If
                
    End If
  
End Sub

Private Sub cmdConvert_Click()

    If txtConvert = "" Then
        If Label8 = "enter min ---- then click" Then
            Label8 = "enter sec ---- then click"
            Frame6.Caption = "Convert sec to min"
            lblMinSec = "sec"
        ElseIf Label8 = "enter sec ---- then click" Then
            Label8 = "enter min ---- then click"
            Frame6.Caption = "Convert min to sec"
            lblMinSec = "min"
        End If
    End If

    If Not IsNumeric(txtConvert) And txtConvert <> "" Then
         MsgBox "You have entered the non-numeric character " & txtConvert, vbOKOnly, "Entry Error"
         txtConvert = ""
         txtConvert.SetFocus
         Exit Sub
    End If

    If IsNumeric(txtConvert) And txtConvert <> "" And txtConvert <> "0" Then
      
        Dim convert As String
        
        If Label8 = "enter min ---- then click" Then
            convert = Val(txtConvert * 60)
            txtConvert = convert
            Label8 = "enter sec ---- then click"
            Frame6.Caption = "Convert sec to min"
            lblMinSec = "sec"
            txtConvert.MaxLength = 4
        
        ElseIf Label8 = "enter sec ---- then click" Then
            convert = Val(txtConvert / 60)
            txtConvert = convert
            Label8 = "enter min ---- then click"
            Frame6.Caption = "Convert min to sec"
            lblMinSec = "min"
            txtConvert.MaxLength = 3
        Else
            txtConvert = ""
        End If
    End If

End Sub

Private Sub Form_Click()
    If fraIntro.Visible = True Then
        cmdCloseTimeSet_Click
    End If
    pSetFocus
End Sub

Private Sub Form_DblClick()
    pSetFocus
End Sub

Private Sub Form_Deactivate()
    
    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
    
     If frmPlanner!mnuLinkTimeRemain.Checked = True Then 'clearilng non-linked time entries
    
        If txtMinute4 <> frmPlanner!txtMinute1 And frmPlanner!txtMinute1 = "" Then
        txtComposer4 = ""
        txtMinute4 = ""
        End If

        If txtMinute5 <> frmPlanner!txtMinute2 And frmPlanner!txtMinute2 = "" Then
        txtComposer5 = ""
        txtMinute5 = ""
        End If

        If txtMinute6 <> frmPlanner!txtMinute3 And frmPlanner!txtMinute3 = "" Then
        txtComposer6 = ""
        txtMinute6 = ""
        End If

        If txtMinute7 <> frmPlanner!txtMinute4 And frmPlanner!txtMinute4 = "" Then
        txtComposer7 = ""
        txtMinute7 = ""
        End If

        If txtMinute8 <> frmPlanner!txtMinute5 And frmPlanner!txtMinute5 = "" Then
        txtComposer8 = ""
        txtMinute8 = ""
        End If

        If txtMinute9 <> frmPlanner!txtMinute6 And frmPlanner!txtMinute6 = "" Then
        txtComposer9 = ""
        txtMinute9 = ""
        End If

        If txtMinute10 <> frmPlanner!txtMinute7 And frmPlanner!txtMinute7 = "" Then
        txtComposer10 = ""
        txtMinute10 = ""
        End If

        If txtMinute11 <> frmPlanner!txtMinute8 And frmPlanner!txtMinute8 = "" Then
        txtComposer11 = ""
        txtMinute11 = ""
        End If
        '----------
        If txtSecond4 <> frmPlanner!txtSecond1 And frmPlanner!txtSecond1 = "" Then
        txtSecond4 = ""
        End If

        If txtSecond5 <> frmPlanner!txtSecond2 And frmPlanner!txtSecond2 = "" Then
        txtSecond5 = ""
        End If

        If txtSecond6 <> frmPlanner!txtSecond3 And frmPlanner!txtSecond3 = "" Then
        txtSecond6 = ""
        End If

        If txtSecond7 <> frmPlanner!txtSecond4 And frmPlanner!txtSecond4 = "" Then
        txtSecond7 = ""
        End If

        If txtSecond8 <> frmPlanner!txtSecond5 And frmPlanner!txtSecond5 = "" Then
        txtSecond8 = ""
        End If

        If txtSecond9 <> frmPlanner!txtSecond6 And frmPlanner!txtSecond6 = "" Then
        txtSecond9 = ""
        End If

        If txtSecond10 <> frmPlanner!txtSecond7 And frmPlanner!txtSecond7 = "" Then
        txtSecond10 = ""
        End If

        If txtSecond11 <> frmPlanner!txtSecond8 And frmPlanner!txtSecond8 = "" Then
        txtSecond11 = ""
        End If
    End If

 '------------
    If txtMinute4 = frmPlanner!txtMinute1 And txtMinute5 = frmPlanner!txtMinute2 And txtMinute6 = frmPlanner!txtMinute3 _
    And txtMinute7 = frmPlanner!txtMinute4 And txtMinute8 = frmPlanner!txtMinute5 And txtMinute9 = frmPlanner!txtMinute6 _
    And txtMinute10 = frmPlanner!txtMinute7 And txtMinute11 = frmPlanner!txtMinute8 Then
            
        shpLink.BackColor = &HC000&
        shpLink.BorderColor = &HC000&
        giTimesDiffer = 0
        
        If frmPlanner!lblProgramTime.Caption = " Planning and program lineup times differ " Then
            frmPlanner!lblProgramTime = lblTotalS & " " & sPlannerProgramTime & sRemain
        End If
        
    End If
    
    If gcTotalMinNT = gcMCal1 Then
        cmdRestoreEntries.BackColor = &H8000000F 'gray
        
        cmdRestoreEntries.ToolTipText = ""
    End If
End Sub

Private Sub Form_Load()

    txtBlock = "60.0"
    
On Error GoTo HandleErrors

    Dim PlanTime, IntroOut, sClose, Spot, Signature As String
    
    Open "Times.dat" For Input As #23
    Input #23, PlanTime, IntroOut, sClose, Spot
    Close #23
          
     If IntroOut = "" Or sClose = "" Or Spot = "" Or PlanTime = "" Then
        MsgBox "Default program times are not set", vbOKOnly, "Default Data Missing"
        giClockShow = 5
        cmdDefaults.Enabled = False
        giClockShow = 0
        Exit Sub
    End If
      
    txtBlock = "60.0"

    miCloseOut = sClose

    If IntroOut <> "" Then
        txtIntroSetting = IntroOut
    Else
        txtIntroSetting = "50"
    End If

    If sClose = "" Then
        txtCloseOut = "30"
    Else
        txtCloseOut = sClose
    End If

    If Spot = "" Then
        txtSpotLengthSetting = "30"
    Else
        txtSpotLengthSetting = Spot
    End If

    lblSpots.Caption = "Enter the number of (" & Spot & _
    "-second average time) spot, promo, PSA, weather, etc. inserts scheduled in the hour (or half-hour) time period"

    Dim iCloseout As String
    
    If Val(miCloseOut) < 60 Then
      ' 'Frame8.Caption = Format(miCloseOut, " ##") & " sec allocated for Closeout && ID"
       'Frame8.Caption = "Seconds Allocated for Closeout && ID"
       
    Else
       iCloseout = Val(miCloseOut) / 60
'       'Frame8.Caption = Format(iCloseout, " 0.#") & " min allocated for Closeout && ID"
       'Frame8.Caption = "Seconds Allocated for Closeout && ID"
       
    End If

    Label20.Caption = " • " & miCloseOut & " secs will remain for show closeout && station ID"

    Label15.Caption = "Time allocated for CLOSEOUT and station ID is " & txtCloseOut & " sec. To change, overwrite the entry in the 'time allocated for closeout && ID' box below. To save the change, click the 'Save Your Changes' button."
     
    chkAnnounce.Caption = "Check if you do Not want to include estimated announce times of " & txtIntro & _
    " sec for each selection and " & txtCloseOut & " sec for program closeout"
    
    Randomize 'using clock, set seed or starting point for random events
    iRandomNumber = Int(Rnd * 10) 'random msgbox reminders
    
    Exit Sub
    
HandleErrors:
    txtIntroSetting = "51"
    txtCloseOut = "31"
    txtSpotLengthSetting = "31"
End Sub
Private Sub Form_Activate()

    If frmPlanner!txtComposer1 <> "" Or frmPlanner!txtComposer2 <> "" Or frmPlanner!txtComposer3 <> "" Or frmPlanner!txtComposer4 <> "" Or _
    frmPlanner!txtComposer5 <> "" Or frmPlanner!txtComposer6 <> "" Or frmPlanner!txtComposer7 <> "" Or frmPlanner!txtComposer8 <> "" Then
        mnuImport.Enabled = True
    Else
        mnuImport.Enabled = False
    End If
    
    If txtComposer4 <> "" Or txtComposer5 <> "" Or txtComposer6 <> "" Or txtComposer7 <> "" Or _
    txtComposer8 <> "" Or txtComposer9 <> "" Or txtComposer10 <> "" Or txtComposer11 <> "" Then
        mnuExport.Enabled = True
    Else
        mnuExport.Enabled = False
    End If
   
    iHourNow = Hour(Now)
   
    cmdSetTime.Visible = False
    Label6.ForeColor = &HC00000    'BLUE
    cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
    cmdSystemTime.SetFocus
   
End Sub

Private Sub pSetFocus()
    'set focus to first blank minute entry box
On Error GoTo HandleErrors
 
   If txtComposer4 = "" Then
        txtComposer4.SetFocus
    ElseIf txtMinute4 = "" And ck4Played.Value = 0 Then
        txtMinute4.SetFocus
    ElseIf txtComposer5 = "" Then
        txtComposer5.SetFocus
    ElseIf txtMinute5 = "" And ck5Played.Value = 0 Then
        txtMinute5.SetFocus
    ElseIf txtComposer6 = "" Then
        txtComposer6.SetFocus
    ElseIf txtMinute6 = "" And ck6Played.Value = 0 Then
        txtMinute6.SetFocus
    ElseIf txtComposer7 = "" Then
        txtComposer7.SetFocus
    ElseIf txtMinute7 = "" And ck7Played.Value = 0 Then
        txtMinute7.SetFocus
    
    ElseIf txtComposer8 = "" Then
        txtComposer8.SetFocus
    ElseIf txtMinute8 = "" And ck8Played.Value = 0 Then
        txtMinute8.SetFocus
    
    ElseIf txtComposer9 = "" Then
        txtComposer9.SetFocus
    ElseIf txtMinute9 = "" And ck9Played.Value = 0 Then
        txtMinute9.SetFocus
        
    ElseIf txtComposer10 = "" Then
        txtComposer10.SetFocus
    ElseIf txtMinute10 = "" And ck10Played.Value = 0 Then
        txtMinute10.SetFocus
    
    ElseIf txtComposer11 = "" Then
        txtComposer11.SetFocus
    ElseIf txtMinute11 = "" And ck11Played.Value = 0 Then
        txtMinute11.SetFocus
    
    Else
      cmdClearLineup.SetFocus
    End If
    
HandleErrors:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
    
    Unload frmMemos
    frmPlanner.Show
End Sub

Private Sub Frame2_DblClick()
    pSetFocus
End Sub

Private Sub Frame7_Click()
    If (iRandomNumber = 1 Or iRandomNumber = 5 Or iRandomNumber = 9) Then
        frmNote.Show vbModal
    End If
End Sub

Private Sub imgClockSetTime_Click()
    If chkStopWatch.Value = 1 And ((IsNull(mRunTime) = True Or mRunTime = 0 Or mRunTime = "")) Then
        MsgBox "To be the source for the 'Current Time' setting, the StopWatch timer must me running." & vbCrLf & "Click F12 to access the Stopwatch and start the timer.", vbOKOnly + vbExclamation, "StopWatch Timer is Not Running"
        cmdSetTime.Visible = False
        Exit Sub
    End If
    
    iCurrentTime = 1
    
    If chkBackAnnc.Value = 1 Then
        chkBackAnnc.Value = 0
    End If
    
    imgHand.Visible = False
    imgDisc.Visible = False

    If cmdSetTime.Visible = False Then
        cmdSetTime.Visible = True
        Frame7.Caption = "CANCEL"
        cmdSystemTime.Caption = "CANCEL Setting Current Time Past the &Hour  F5"
        cmdSetTime.SetFocus
    
    ElseIf cmdSetTime.Visible = True Then
        cmdSetTime.Visible = False
        Frame7.Caption = "Set Current Time"
        If chkStopWatch.Value = 0 Then
            
        If txtMinAdj = "" And txtSecAdj = "" Then
            cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
        Else
             cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
        End If
            
        Else
            cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Stopwatch) F5"
        End If
        
        pSetFocus
    End If
    
End Sub

Private Sub Label17_DblClick()
    txtAnnc4 = ""
    txtAnnc5 = ""
    txtAnnc6 = ""
    txtAnnc7 = ""
    txtAnnc8 = ""
    txtAnnc9 = ""
    txtAnnc10 = ""
    txtAnnc11 = ""
End Sub

Private Sub Label31_DblClick()
    If txtSpotsS > "1" Then
        txtSpotsS = Format((Val(txtSpotsS) - 1), "##")
    Else
        txtSpotsS = ""
    End If
End Sub

Private Sub Label33_Click()

    If txtSecAdj <> "" Then
        txtSecAdj = ""
        'Exit Sub
    End If
    
    If txtMinAdj <> "" Then
        txtMinAdj = ""
    End If
    
    If txtMinAdj = "" And txtSecAdj = "" Then
        txtMinAdj.SetFocus
    End If
    
    If txtMinAdj = "" And txtSecAdj = "" Then
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
    End If
    
End Sub

Private Sub lblAdjustTimeExit_Click()

        txtMinAdj.Visible = False
        txtMinAdj.Text = ""
        txtSecAdj.Visible = False
        txtSecAdj.Text = ""
        Label29.Visible = False
        lblMinutes.Visible = False
        lblSeconds.Visible = False
        fraAdjustTime.Visible = False
        fraAdjustedTime.Visible = False
        cmdAdjTime.Visible = True
        fraTimeAdjust2.Visible = True
        
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
End Sub

Private Sub Label21_DblClick()

    If txtBackAnnc = "" Then 'prevents runtime error (mismatch) if txtBackAnnc = ""
        txtBackAnnc = "0"
    End If
    
     If chkBackAnnc.Value = 0 And Val(txtBackAnnc.Text) <> (txtIntro / 2) Then
     
        If Val(txtIntro) > 29 Then
            txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
        Else
            txtBackAnnc = "0"
        End If
       
        Label21.ToolTipText = ""
    End If
    
End Sub

Private Sub Label24_DblClick()
    If Val(txtCloseOut) < 15 Then
        txtCloseOut = frmDefaults!txtClose
    Else
        txtCloseOut = "0"
    End If
End Sub

Private Sub Label29_dblClick()
    MsgBox "If the computer clock is slow, enter in minutes and seconds the difference between the correct time" _
        & vbCrLf & "and the computer clock time. If the computer clock is fast, enter the difference as a negative number." _
        & vbCrLf & vbCrLf & "• Computer clock slow, enter difference." & vbCrLf & vbCrLf & "• Computer clock fast, enter difference as a negative.", _
         vbOKOnly, "Clock Error Adjustment"
End Sub

Private Sub Label3_dblClick()

On Error GoTo HandleErrors

    If txtSpotLength = "" Then
        Dim PlanTime, IntroOut, sClose, Spot, Signature As String
    
        Open "Times.dat" For Input As #23
        Input #23, PlanTime, IntroOut, sClose, Spot
        Close #23
        
        txtSpotLength = Spot
        txtSpotLength.SelStart = 0 'begin selection at start
        txtSpotLength.SelLength = Len(txtSpotLength)
        
    Else
        txtSpotLength = ""
        txtSpotLength.SetFocus
    End If
    
HandleErrors:
End Sub

Private Sub lbl1_dblClick()

    If mnuToolsExportLineCopy.Checked = True And txtComposer4 <> "" And txtMinute4 <> "" Then
    
        Dim iResponse As Integer

        iResponse = MsgBox("Line 1: Composer (" & txtComposer4 & "), CD number (if existing), and playing time will be copied to Music Planning page lineup.", _
        vbOKCancel + vbInformation, "Export Line 1")
        
        If iResponse = vbOK Then
            
                If frmPlanner!txtComposer1 = "" Then
                    frmPlanner!txtComposer1 = txtComposer4
                    frmPlanner!txtCD1 = txtCD4
                    frmPlanner!txtAnnc1 = txtAnnc4
                    frmPlanner!txtMinute1 = txtMinute4
                    frmPlanner!txtSecond1 = txtSecond4
                
                ElseIf frmPlanner!txtComposer2 = "" Then
                    frmPlanner!txtComposer2 = txtComposer4
                    frmPlanner!txtCD2 = txtCD4
                    frmPlanner!txtAnnc2 = txtAnnc4
                    frmPlanner!txtMinute2 = txtMinute4
                    frmPlanner!txtSecond2 = txtSecond4
                    
                ElseIf frmPlanner!txtComposer3 = "" Then
                    frmPlanner!txtComposer3 = txtComposer4
                    frmPlanner!txtCD3 = txtCD4
                    frmPlanner!txtAnnc3 = txtAnnc4
                    frmPlanner!txtMinute3 = txtMinute4
                    frmPlanner!txtSecond3 = txtSecond4
                    
                ElseIf frmPlanner!txtComposer4 = "" Then
                    frmPlanner!txtComposer4 = txtComposer4
                    frmPlanner!txtCD4 = txtCD4
                    frmPlanner!txtAnnc4 = txtAnnc4
                    frmPlanner!txtMinute4 = txtMinute4
                    frmPlanner!txtSecond4 = txtSecond4
                    
                ElseIf frmPlanner!txtComposer5 = "" Then
                    frmPlanner!txtComposer5 = txtComposer4
                    frmPlanner!txtCD5 = txtCD4
                    frmPlanner!txtAnnc5 = txtAnnc4
                    frmPlanner!txtMinute5 = txtMinute4
                    frmPlanner!txtSecond5 = txtSecond4
                    
                ElseIf frmPlanner!txtComposer6 = "" Then
                    frmPlanner!txtComposer6 = txtComposer4
                    frmPlanner!txtCD6 = txtCD4
                    frmPlanner!txtAnnc6 = txtAnnc4
                    frmPlanner!txtMinute6 = txtMinute4
                    frmPlanner!txtSecond6 = txtSecond4
                    
                ElseIf frmPlanner!txtComposer7 = "" Then
                    frmPlanner!txtComposer7 = txtComposer4
                    frmPlanner!txtCD7 = txtCD4
                    frmPlanner!txtAnnc7 = txtAnnc4
                    frmPlanner!txtMinute7 = txtMinute4
                    frmPlanner!txtSecond7 = txtSecond4
                    
                ElseIf frmPlanner!txtComposer8 = "" Then
                    frmPlanner!txtComposer8 = txtComposer4
                    frmPlanner!txtCD8 = txtCD4
                    frmPlanner!txtAnnc8 = txtAnnc4
                    frmPlanner!txtMinute8 = txtMinute4
                    frmPlanner!txtSecond8 = txtSecond4
                    
                Else
                    frmPlanner!txtComposer1 = txtComposer4
                
                End If
      
            lbl1.BorderStyle = 1
            txtMinute4.ForeColor = &H80000008  'black
            txtSecond4.ForeColor = &H80000008  'black
            
            frmPlanner!txtComposition1 = ""
            frmPlanner!txtTrack1 = ""
            frmPlanner!txtDisc1 = ""
            
             pAdd 'prevents "Time remain and planning times differ" message
        End If
    End If
End Sub

Private Sub lbl2_dblClick()

    If mnuToolsExportLineCopy.Checked = True And txtComposer5 <> "" And txtMinute5 <> "" Then
    
        Dim iResponse As Integer

        iResponse = MsgBox("Line 2: Composer (" & txtComposer5 & "), CD number (if existing), and playing time will be copied to Music Planning page lineup.", _
        vbOKCancel + vbInformation, "Export Line 2")
        
        If iResponse = vbOK Then
            
            If frmPlanner!txtComposer1 = "" Then
                frmPlanner!txtComposer1 = txtComposer5
                frmPlanner!txtCD1 = txtCD5
                frmPlanner!txtAnnc1 = txtAnnc5
                frmPlanner!txtMinute1 = txtMinute5
                frmPlanner!txtSecond1 = txtSecond5
            
            ElseIf frmPlanner!txtComposer2 = "" Then
                frmPlanner!txtComposer2 = txtComposer5
                frmPlanner!txtCD2 = txtCD5
                frmPlanner!txtAnnc2 = txtAnnc5
                frmPlanner!txtMinute2 = txtMinute5
                frmPlanner!txtSecond2 = txtSecond5
                
            ElseIf frmPlanner!txtComposer3 = "" Then
                frmPlanner!txtComposer3 = txtComposer5
                frmPlanner!txtCD3 = txtCD5
                frmPlanner!txtAnnc3 = txtAnnc5
                frmPlanner!txtMinute3 = txtMinute5
                frmPlanner!txtSecond3 = txtSecond5
                
            ElseIf frmPlanner!txtComposer4 = "" Then
                frmPlanner!txtComposer4 = txtComposer5
                frmPlanner!txtCD4 = txtCD5
                frmPlanner!txtAnnc4 = txtAnnc5
                frmPlanner!txtMinute4 = txtMinute5
                frmPlanner!txtSecond4 = txtSecond5
                
            ElseIf frmPlanner!txtComposer5 = "" Then
                frmPlanner!txtComposer5 = txtComposer5
                frmPlanner!txtCD5 = txtCD5
                frmPlanner!txtAnnc5 = txtAnnc5
                frmPlanner!txtMinute5 = txtMinute5
                frmPlanner!txtSecond5 = txtSecond5
                
            ElseIf frmPlanner!txtComposer6 = "" Then
                frmPlanner!txtComposer6 = txtComposer5
                frmPlanner!txtCD6 = txtCD5
                frmPlanner!txtAnnc6 = txtAnnc5
                frmPlanner!txtMinute6 = txtMinute5
                frmPlanner!txtSecond6 = txtSecond5
                
            ElseIf frmPlanner!txtComposer7 = "" Then
                frmPlanner!txtComposer7 = txtComposer5
                frmPlanner!txtCD7 = txtCD5
                frmPlanner!txtAnnc7 = txtAnnc5
                frmPlanner!txtMinute7 = txtMinute5
                frmPlanner!txtSecond7 = txtSecond5
                
            ElseIf frmPlanner!txtComposer8 = "" Then
                frmPlanner!txtComposer8 = txtComposer5
                frmPlanner!txtCD8 = txtCD5
                frmPlanner!txtAnnc8 = txtAnnc5
                frmPlanner!txtMinute8 = txtMinute5
                frmPlanner!txtSecond8 = txtSecond5
                
            Else
                frmPlanner!txtComposer2 = txtComposer5
            
            End If
                        
            lbl2.BorderStyle = 1
            txtMinute5.ForeColor = &H80000008  'black
            txtSecond5.ForeColor = &H80000008  'black
            
            frmPlanner!txtComposition2 = ""
            frmPlanner!txtTrack2 = ""
            frmPlanner!txtDisc2 = ""
          
            pAdd 'prevents "Time remain and planning times differ" message
        End If
    End If
End Sub

Private Sub lbl3_dblClick()
    
    If mnuToolsExportLineCopy.Checked = True And txtComposer6 <> "" And txtMinute6 <> "" Then
    
    Dim iResponse As Integer

    iResponse = MsgBox("Line 3: Composer (" & txtComposer6 & "), CD number (if existing), and playing time will be copied to Music Planning page lineup.", _
    vbOKCancel + vbInformation, "Export Line 3")
        
        If iResponse = vbOK Then
        
            If frmPlanner!txtComposer1 = "" Then
                 frmPlanner!txtComposer1 = txtComposer6
                 frmPlanner!txtCD1 = txtCD6
                 frmPlanner!txtAnnc1 = txtAnnc6
                 frmPlanner!txtMinute1 = txtMinute6
                 frmPlanner!txtSecond1 = txtSecond6
             
             ElseIf frmPlanner!txtComposer2 = "" Then
                 frmPlanner!txtComposer2 = txtComposer6
                 frmPlanner!txtCD2 = txtCD6
                 frmPlanner!txtAnnc2 = txtAnnc6
                 frmPlanner!txtMinute2 = txtMinute6
                 frmPlanner!txtSecond2 = txtSecond6
                 
             ElseIf frmPlanner!txtComposer3 = "" Then
                 frmPlanner!txtComposer3 = txtComposer6
                 frmPlanner!txtCD3 = txtCD6
                 frmPlanner!txtAnnc3 = txtAnnc6
                 frmPlanner!txtMinute3 = txtMinute6
                 frmPlanner!txtSecond3 = txtSecond6
                 
             ElseIf frmPlanner!txtComposer4 = "" Then
                 frmPlanner!txtComposer4 = txtComposer6
                 frmPlanner!txtCD4 = txtCD6
                 frmPlanner!txtAnnc4 = txtAnnc6
                 frmPlanner!txtMinute4 = txtMinute6
                 frmPlanner!txtSecond4 = txtSecond6
                 
             ElseIf frmPlanner!txtComposer5 = "" Then
                 frmPlanner!txtComposer5 = txtComposer6
                 frmPlanner!txtCD5 = txtCD6
                 frmPlanner!txtAnnc5 = txtAnnc6
                 frmPlanner!txtMinute5 = txtMinute6
                 frmPlanner!txtSecond5 = txtSecond6
                 
             ElseIf frmPlanner!txtComposer6 = "" Then
                 frmPlanner!txtComposer6 = txtComposer6
                 frmPlanner!txtCD6 = txtCD6
                 frmPlanner!txtAnnc6 = txtAnnc6
                 frmPlanner!txtMinute6 = txtMinute6
                 frmPlanner!txtSecond6 = txtSecond6
                 
             ElseIf frmPlanner!txtComposer7 = "" Then
                 frmPlanner!txtComposer7 = txtComposer6
                 frmPlanner!txtCD7 = txtCD6
                 frmPlanner!txtAnnc7 = txtAnnc6
                 frmPlanner!txtMinute7 = txtMinute6
                 frmPlanner!txtSecond7 = txtSecond6
                 
             ElseIf frmPlanner!txtComposer8 = "" Then
                 frmPlanner!txtComposer8 = txtComposer6
                 frmPlanner!txtCD8 = txtCD6
                 frmPlanner!txtAnnc8 = txtAnnc6
                 frmPlanner!txtMinute8 = txtMinute6
                 frmPlanner!txtSecond8 = txtSecond6
                 
             Else
                 frmPlanner!txtComposer3 = txtComposer6
             
             End If
                        
            lbl3.BorderStyle = 1
            txtMinute6.ForeColor = &H80000008  'black
            txtSecond6.ForeColor = &H80000008  'black
            
            frmPlanner!txtComposition3 = ""
            frmPlanner!txtTrack3 = ""
            frmPlanner!txtDisc3 = ""
            Beep
            
            frmPlanner!txtComposition3 = ""
            frmPlanner!txtTrack3 = ""
            frmPlanner!txtDisc3 = ""
            
            pAdd 'prevents "Time remain and planning times differ" message
        End If
    End If
End Sub

Private Sub lbl4_dblClick()

    If mnuToolsExportLineCopy.Checked = True And txtComposer7 <> "" And txtMinute7 <> "" Then
    
    Dim iResponse As Integer

        iResponse = MsgBox("Line 4: Composer (" & txtComposer7 & "), CD number (if existing), and playing time will be copied to Music Planning page lineup.", _
        vbOKCancel + vbInformation, "Export Line 4")
        
        If iResponse = vbOK Then
           
                If frmPlanner!txtComposer1 = "" Then
                    frmPlanner!txtComposer1 = txtComposer7
                    frmPlanner!txtCD1 = txtCD7
                    frmPlanner!txtAnnc1 = txtAnnc7
                    frmPlanner!txtMinute1 = txtMinute7
                    frmPlanner!txtSecond1 = txtSecond7
                
                ElseIf frmPlanner!txtComposer2 = "" Then
                    frmPlanner!txtComposer2 = txtComposer7
                    frmPlanner!txtCD2 = txtCD7
                    frmPlanner!txtAnnc2 = txtAnnc7
                    frmPlanner!txtMinute2 = txtMinute7
                    frmPlanner!txtSecond2 = txtSecond7
                    
                ElseIf frmPlanner!txtComposer3 = "" Then
                    frmPlanner!txtComposer3 = txtComposer7
                    frmPlanner!txtCD3 = txtCD7
                    frmPlanner!txtAnnc3 = txtAnnc7
                    frmPlanner!txtMinute3 = txtMinute7
                    frmPlanner!txtSecond3 = txtSecond7
                    
                ElseIf frmPlanner!txtComposer4 = "" Then
                    frmPlanner!txtComposer4 = txtComposer7
                    frmPlanner!txtCD4 = txtCD7
                    frmPlanner!txtAnnc4 = txtAnnc7
                    frmPlanner!txtMinute4 = txtMinute7
                    frmPlanner!txtSecond4 = txtSecond7
                    
                ElseIf frmPlanner!txtComposer5 = "" Then
                    frmPlanner!txtComposer5 = txtComposer7
                    frmPlanner!txtCD5 = txtCD7
                    frmPlanner!txtAnnc5 = txtAnnc7
                    frmPlanner!txtMinute5 = txtMinute7
                    frmPlanner!txtSecond5 = txtSecond7
                    
                ElseIf frmPlanner!txtComposer6 = "" Then
                    frmPlanner!txtComposer6 = txtComposer7
                    frmPlanner!txtCD6 = txtCD7
                    frmPlanner!txtAnnc6 = txtAnnc7
                    frmPlanner!txtMinute6 = txtMinute7
                    frmPlanner!txtSecond6 = txtSecond7
                    
                ElseIf frmPlanner!txtComposer7 = "" Then
                    frmPlanner!txtComposer7 = txtComposer7
                    frmPlanner!txtCD7 = txtCD7
                    frmPlanner!txtAnnc7 = txtAnnc7
                    frmPlanner!txtMinute7 = txtMinute7
                    frmPlanner!txtSecond7 = txtSecond7
                    
                ElseIf frmPlanner!txtComposer8 = "" Then
                    frmPlanner!txtComposer8 = txtComposer7
                    frmPlanner!txtCD8 = txtCD7
                    frmPlanner!txtAnnc8 = txtAnnc7
                    frmPlanner!txtMinute8 = txtMinute7
                    frmPlanner!txtSecond8 = txtSecond7
                    
                Else
                    frmPlanner!txtComposer4 = txtComposer7
                
                End If
                        
            lbl4.BorderStyle = 1
            txtMinute7.ForeColor = &H80000008  'black
            txtSecond7.ForeColor = &H80000008  'black
            
            frmPlanner!txtComposition4 = ""
            frmPlanner!txtTrack4 = ""
            frmPlanner!txtDisc4 = ""
            
            pAdd 'prevents "Time remain and planning times differ" message
        End If
    End If
End Sub

Private Sub lbl5_dblClick()

    If mnuToolsExportLineCopy.Checked = True And txtComposer8 <> "" And txtMinute8 <> "" Then
    
    Dim iResponse As Integer

        iResponse = MsgBox("Line 5: Composer (" & txtComposer8 & "), CD number (if existing), and playing time will be copied to Music Planning page lineup.", _
        vbOKCancel + vbInformation, "Export Line 5")
        
        If iResponse = vbOK Then
        
                If frmPlanner!txtComposer1 = "" Then
                    frmPlanner!txtComposer1 = txtComposer8
                    frmPlanner!txtCD1 = txtCD8
                    frmPlanner!txtAnnc1 = txtAnnc8
                    frmPlanner!txtMinute1 = txtMinute8
                    frmPlanner!txtSecond1 = txtSecond8
                
                ElseIf frmPlanner!txtComposer2 = "" Then
                    frmPlanner!txtComposer2 = txtComposer8
                    frmPlanner!txtCD2 = txtCD8
                    frmPlanner!txtAnnc2 = txtAnnc8
                    frmPlanner!txtMinute2 = txtMinute8
                    frmPlanner!txtSecond2 = txtSecond8
                    
                ElseIf frmPlanner!txtComposer3 = "" Then
                    frmPlanner!txtComposer3 = txtComposer8
                    frmPlanner!txtCD3 = txtCD8
                    frmPlanner!txtAnnc3 = txtAnnc8
                    frmPlanner!txtMinute3 = txtMinute8
                    frmPlanner!txtSecond3 = txtSecond8
                    
                ElseIf frmPlanner!txtComposer4 = "" Then
                    frmPlanner!txtComposer4 = txtComposer8
                    frmPlanner!txtCD4 = txtCD8
                    frmPlanner!txtAnnc4 = txtAnnc8
                    frmPlanner!txtMinute4 = txtMinute8
                    frmPlanner!txtSecond4 = txtSecond8
                    
                ElseIf frmPlanner!txtComposer5 = "" Then
                    frmPlanner!txtComposer5 = txtComposer8
                    frmPlanner!txtCD5 = txtCD8
                    frmPlanner!txtAnnc5 = txtAnnc8
                    frmPlanner!txtMinute5 = txtMinute8
                    frmPlanner!txtSecond5 = txtSecond8
                    
                ElseIf frmPlanner!txtComposer6 = "" Then
                    frmPlanner!txtComposer6 = txtComposer8
                    frmPlanner!txtCD6 = txtCD8
                    frmPlanner!txtAnnc6 = txtAnnc8
                    frmPlanner!txtMinute6 = txtMinute8
                    frmPlanner!txtSecond6 = txtSecond8
                    
                ElseIf frmPlanner!txtComposer7 = "" Then
                    frmPlanner!txtComposer7 = txtComposer8
                    frmPlanner!txtCD7 = txtCD8
                    frmPlanner!txtAnnc7 = txtAnnc8
                    frmPlanner!txtMinute7 = txtMinute8
                    frmPlanner!txtSecond7 = txtSecond8
                    
                ElseIf frmPlanner!txtComposer8 = "" Then
                    frmPlanner!txtComposer8 = txtComposer8
                    frmPlanner!txtCD8 = txtCD8
                    frmPlanner!txtAnnc8 = txtAnnc8
                    frmPlanner!txtMinute8 = txtMinute8
                    frmPlanner!txtSecond8 = txtSecond8
                    
                Else
                    frmPlanner!txtComposer5 = txtComposer8
                
                End If
                        
            lbl5.BorderStyle = 1
            txtMinute8.ForeColor = &H80000008  'black
            txtSecond8.ForeColor = &H80000008  'black
                    
            frmPlanner!txtComposition5 = ""
            frmPlanner!txtTrack5 = ""
            frmPlanner!txtDisc5 = ""
           
            pAdd 'prevents "Time remain and planning times differ" message
        End If
    End If
End Sub

Private Sub lbl6_dblClick()

    If mnuToolsExportLineCopy.Checked = True And txtComposer9 <> "" And txtMinute9 <> "" Then
    
    Dim iResponse As Integer

        iResponse = MsgBox("Line 6: Composer (" & txtComposer9 & "), CD number (if existing), and playing time will be copied to Music Planning page lineup.", _
        vbOKCancel + vbInformation, "Export Line 6")
        
        If iResponse = vbOK Then

                If frmPlanner!txtComposer1 = "" Then
                    frmPlanner!txtComposer1 = txtComposer9
                    frmPlanner!txtCD1 = txtCD9
                    frmPlanner!txtAnnc1 = txtAnnc9
                    frmPlanner!txtMinute1 = txtMinute9
                    frmPlanner!txtSecond1 = txtSecond9
                
                ElseIf frmPlanner!txtComposer2 = "" Then
                    frmPlanner!txtComposer2 = txtComposer9
                    frmPlanner!txtCD2 = txtCD9
                    frmPlanner!txtAnnc2 = txtAnnc9
                    frmPlanner!txtMinute2 = txtMinute9
                    frmPlanner!txtSecond2 = txtSecond9
                    
                ElseIf frmPlanner!txtComposer3 = "" Then
                    frmPlanner!txtComposer3 = txtComposer9
                    frmPlanner!txtCD3 = txtCD9
                    frmPlanner!txtAnnc3 = txtAnnc9
                    frmPlanner!txtMinute3 = txtMinute9
                    frmPlanner!txtSecond3 = txtSecond9
                    
                ElseIf frmPlanner!txtComposer4 = "" Then
                    frmPlanner!txtComposer4 = txtComposer9
                    frmPlanner!txtCD4 = txtCD9
                    frmPlanner!txtAnnc4 = txtAnnc9
                    frmPlanner!txtMinute4 = txtMinute9
                    frmPlanner!txtSecond4 = txtSecond9
                    
                ElseIf frmPlanner!txtComposer5 = "" Then
                    frmPlanner!txtComposer5 = txtComposer9
                    frmPlanner!txtCD5 = txtCD9
                    frmPlanner!txtAnnc5 = txtAnnc9
                    frmPlanner!txtMinute5 = txtMinute9
                    frmPlanner!txtSecond5 = txtSecond9
                    
                ElseIf frmPlanner!txtComposer6 = "" Then
                    frmPlanner!txtComposer6 = txtComposer9
                    frmPlanner!txtCD6 = txtCD9
                    frmPlanner!txtAnnc6 = txtAnnc9
                    frmPlanner!txtMinute6 = txtMinute9
                    frmPlanner!txtSecond6 = txtSecond9
                    
                ElseIf frmPlanner!txtComposer7 = "" Then
                    frmPlanner!txtComposer7 = txtComposer9
                    frmPlanner!txtCD7 = txtCD9
                    frmPlanner!txtAnnc7 = txtAnnc9
                    frmPlanner!txtMinute7 = txtMinute9
                    frmPlanner!txtSecond7 = txtSecond9
                    
                ElseIf frmPlanner!txtComposer8 = "" Then
                    frmPlanner!txtComposer8 = txtComposer9
                    frmPlanner!txtCD8 = txtCD9
                    frmPlanner!txtAnnc8 = txtAnnc9
                    frmPlanner!txtMinute8 = txtMinute9
                    frmPlanner!txtSecond8 = txtSecond9
                    
                Else
                    frmPlanner!txtComposer6 = txtComposer9
                
                End If
                        
            lbl6.BorderStyle = 1
            txtMinute9.ForeColor = &H80000008  'black
            txtSecond9.ForeColor = &H80000008  'black
            
            frmPlanner!txtComposition6 = ""
            frmPlanner!txtTrack6 = ""
            frmPlanner!txtDisc6 = ""
            
            pAdd 'prevents "Time remain and planning times differ" message
        End If
    End If
End Sub

Private Sub lbl7_dblClick()

    If mnuToolsExportLineCopy.Checked = True And txtComposer10 <> "" And txtMinute10 <> "" Then
    
    Dim iResponse As Integer

        iResponse = MsgBox("Line 7: Composer (" & txtComposer10 & "), CD number (if existing), and playing time will be copied to Music Planning page lineup.", _
        vbOKCancel + vbInformation, "Export Line 7")
        
        If iResponse = vbOK Then
        
                If frmPlanner!txtComposer1 = "" Then
                    frmPlanner!txtComposer1 = txtComposer10
                    frmPlanner!txtCD1 = txtCD10
                    frmPlanner!txtAnnc1 = txtAnnc10
                    frmPlanner!txtMinute1 = txtMinute10
                    frmPlanner!txtSecond1 = txtSecond10
                
                ElseIf frmPlanner!txtComposer2 = "" Then
                    frmPlanner!txtComposer2 = txtComposer10
                    frmPlanner!txtCD2 = txtCD10
                    frmPlanner!txtAnnc2 = txtAnnc10
                    frmPlanner!txtMinute2 = txtMinute10
                    frmPlanner!txtSecond2 = txtSecond10
                    
                ElseIf frmPlanner!txtComposer3 = "" Then
                    frmPlanner!txtComposer3 = txtComposer10
                    frmPlanner!txtCD3 = txtCD10
                    frmPlanner!txtAnnc3 = txtAnnc10
                    frmPlanner!txtMinute3 = txtMinute10
                    frmPlanner!txtSecond3 = txtSecond10
                    
                ElseIf frmPlanner!txtComposer4 = "" Then
                    frmPlanner!txtComposer4 = txtComposer10
                    frmPlanner!txtCD4 = txtCD10
                    frmPlanner!txtAnnc4 = txtAnnc10
                    frmPlanner!txtMinute4 = txtMinute10
                    frmPlanner!txtSecond4 = txtSecond10
                    
                ElseIf frmPlanner!txtComposer5 = "" Then
                    frmPlanner!txtComposer5 = txtComposer10
                    frmPlanner!txtCD5 = txtCD10
                    frmPlanner!txtAnnc5 = txtAnnc10
                    frmPlanner!txtMinute5 = txtMinute10
                    frmPlanner!txtSecond5 = txtSecond10
                    
                ElseIf frmPlanner!txtComposer6 = "" Then
                    frmPlanner!txtComposer6 = txtComposer10
                    frmPlanner!txtCD6 = txtCD10
                    frmPlanner!txtAnnc6 = txtAnnc10
                    frmPlanner!txtMinute6 = txtMinute10
                    frmPlanner!txtSecond6 = txtSecond10
                    
                ElseIf frmPlanner!txtComposer7 = "" Then
                    frmPlanner!txtComposer7 = txtComposer10
                    frmPlanner!txtCD7 = txtCD10
                    frmPlanner!txtAnnc7 = txtAnnc10
                    frmPlanner!txtMinute7 = txtMinute10
                    frmPlanner!txtSecond7 = txtSecond10
                    
                ElseIf frmPlanner!txtComposer8 = "" Then
                    frmPlanner!txtComposer8 = txtComposer10
                    frmPlanner!txtCD8 = txtCD10
                    frmPlanner!txtAnnc8 = txtAnnc10
                    frmPlanner!txtMinute8 = txtMinute10
                    frmPlanner!txtSecond8 = txtSecond10
                    
                Else
                    frmPlanner!txtComposer7 = txtComposer10
                
                End If
                        
            lbl7.BorderStyle = 1
            txtMinute10.ForeColor = &H80000008  'black
            txtSecond10.ForeColor = &H80000008  'black
                    
            frmPlanner!txtComposition7 = ""
            frmPlanner!txtTrack7 = ""
            frmPlanner!txtDisc7 = ""
           
            pAdd 'prevents "Time remain and planning times differ" message
        End If
    End If
End Sub

Private Sub lbl8_dblClick()

    If mnuToolsExportLineCopy.Checked = True And txtComposer11 <> "" And txtMinute11 <> "" Then
    
    Dim iResponse As Integer

        iResponse = MsgBox("Line 8: Composer (" & txtComposer11 & "), CD number (if existing), and playing time will be copied to Music Planning page lineup.", _
        vbOKCancel + vbInformation, "Export Line 8")
        
        If iResponse = vbOK Then
        
                If frmPlanner!txtComposer1 = "" Then
                    frmPlanner!txtComposer1 = txtComposer11
                    frmPlanner!txtCD1 = txtCD11
                    frmPlanner!txtAnnc1 = txtAnnc11
                    frmPlanner!txtMinute1 = txtMinute11
                    frmPlanner!txtSecond1 = txtSecond11
                
                ElseIf frmPlanner!txtComposer2 = "" Then
                    frmPlanner!txtComposer2 = txtComposer11
                    frmPlanner!txtCD2 = txtCD11
                    frmPlanner!txtAnnc2 = txtAnnc11
                    frmPlanner!txtMinute2 = txtMinute11
                    frmPlanner!txtSecond2 = txtSecond11
                    
                ElseIf frmPlanner!txtComposer3 = "" Then
                    frmPlanner!txtComposer3 = txtComposer11
                    frmPlanner!txtCD3 = txtCD11
                    frmPlanner!txtAnnc3 = txtAnnc11
                    frmPlanner!txtMinute3 = txtMinute11
                    frmPlanner!txtSecond3 = txtSecond11
                    
                ElseIf frmPlanner!txtComposer4 = "" Then
                    frmPlanner!txtComposer4 = txtComposer11
                    frmPlanner!txtCD4 = txtCD11
                    frmPlanner!txtAnnc4 = txtAnnc11
                    frmPlanner!txtMinute4 = txtMinute11
                    frmPlanner!txtSecond4 = txtSecond11
                    
                ElseIf frmPlanner!txtComposer5 = "" Then
                    frmPlanner!txtComposer5 = txtComposer11
                    frmPlanner!txtCD5 = txtCD11
                    frmPlanner!txtAnnc5 = txtAnnc11
                    frmPlanner!txtMinute5 = txtMinute11
                    frmPlanner!txtSecond5 = txtSecond11
                    
                ElseIf frmPlanner!txtComposer6 = "" Then
                    frmPlanner!txtComposer6 = txtComposer11
                    frmPlanner!txtCD6 = txtCD11
                    frmPlanner!txtAnnc6 = txtAnnc11
                    frmPlanner!txtMinute6 = txtMinute11
                    frmPlanner!txtSecond6 = txtSecond11
                    
                ElseIf frmPlanner!txtComposer7 = "" Then
                    frmPlanner!txtComposer7 = txtComposer11
                    frmPlanner!txtCD7 = txtCD11
                    frmPlanner!txtAnnc7 = txtAnnc11
                    frmPlanner!txtMinute7 = txtMinute11
                    frmPlanner!txtSecond7 = txtSecond11
                    
                ElseIf frmPlanner!txtComposer8 = "" Then
                    frmPlanner!txtComposer8 = txtComposer11
                    frmPlanner!txtCD8 = txtCD11
                    frmPlanner!txtAnnc8 = txtAnnc11
                    frmPlanner!txtMinute8 = txtMinute11
                    frmPlanner!txtSecond8 = txtSecond11
                    
                Else
                    frmPlanner!txtComposer8 = txtComposer11
                
                End If
                        
            lbl8.BorderStyle = 1
            txtMinute11.ForeColor = &H80000008  'black
            txtSecond11.ForeColor = &H80000008  'black
            frmPlanner!txtComposition8 = ""
            frmPlanner!txtTrack8 = ""
            frmPlanner!txtDisc8 = ""
            
            pAdd 'prevents "Time remain and planning times differ" message
        End If
    End If
End Sub

Private Sub lblAnncTime_Change()

    If lblAnncTime.Caption = " Annc Time" Then
        lblAnncTime.MousePointer = 1
        lblAnncTime.BackColor = &H8000000F 'gray
        lblAnncTime.ForeColor = &H80&  'maroon
        If frmPlanner!mnuLinkTimeRemain.Checked = True Then
            frmPlanner!fraAnnc.ForeColor = &H80000008 'black
            frmPlanner!fraAnnc.ToolTipText = ""
        End If
    Else
        lblAnncTime.MousePointer = 99
        lblAnncTime.BackColor = &H80000018 ' yellow'&HFFFFFF    'white
        
        lblAnncTime.ForeColor = &H404040   '&H808080    'dark gray &H00404040&
        
        If frmPlanner!mnuLinkTimeRemain.Checked = True Then
            frmPlanner!fraAnnc.ForeColor = vbBlue
            frmPlanner!fraAnnc.ToolTipText = " Double-Click to reset Annc Times to " & txtIntro & " sec"
        End If
    End If
        
End Sub

Private Sub lblAverageTime_DblClick()

On Error GoTo HandleErrors

    If txtIntro = "" Then
        Dim PlanTime, IntroOut, sClose, Spot As Integer
        Open "Times.dat" For Input As #23
        Input #23, PlanTime, IntroOut, sClose, Spot
        Close #23
        txtIntro = IntroOut
        txtIntro.SelStart = 0 'begin selection at start
        txtIntro.SelLength = Len(txtIntro)
        
    Else
        txtIntro = ""
        txtIntro.SetFocus
    End If
HandleErrors:
End Sub

Private Sub Label4_dblClick()
    mnuToolsCD_Click
End Sub

Private Sub lblAnncTime_Click()
  
    If Val(txtIntro) > 29 Then
        txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
    Else
        txtBackAnnc = "0"
    End If
    
    If txtMinute4 <> "" And Check1(0).Value = 0 Then
        txtAnnc4 = Val(txtIntro) - Val(txtBackAnnc)
    End If
    
    If txtMinute5 <> "" And Check1(1).Value = 0 Then
        txtAnnc5 = txtIntro
    Else
        txtAnnc5 = ""
    End If
    
    If txtMinute6 <> "" And Check1(2).Value = 0 Then
        txtAnnc6 = txtIntro
    Else
        txtAnnc6 = ""
    End If
    
    If txtMinute7 <> "" And Check1(3).Value = 0 Then
        txtAnnc7 = txtIntro
    Else
        txtAnnc7 = ""
    End If
    
    If txtMinute8 <> "" And Check1(4).Value = 0 Then
        txtAnnc8 = txtIntro
    Else
        txtAnnc8 = ""
    End If
    
    If txtMinute9 <> "" And Check1(5).Value = 0 Then
        txtAnnc9 = txtIntro
    Else
        txtAnnc9 = ""
    End If
    
    If txtMinute10 <> "" And Check1(6).Value = 0 Then
        txtAnnc10 = txtIntro
    Else
        txtAnnc10 = ""
    End If
    
    If txtMinute11 <> "" And Check1(7).Value = 0 And frmPlanner!chkNonCD8.Value = 0 Then
        txtAnnc11 = txtIntro
    Else
        txtAnnc11 = ""
    End If
    
    lblAnncTime.BorderStyle = 0
    frmPlanner!fraAnnc.ForeColor = &H80000008 'black
    frmPlanner!fraAnnc.ToolTipText = ""
    cmdClearAnncTimes.Caption = "Clear Announce Times"
    
    lblAnncTime = " Annc Time"
    lblAnncTime.BorderStyle = 0
    Label27 = "Min"
    Label28 = "Sec"

End Sub

Private Sub lblCurrentTime_Change()

    Dim sSeconds As String 'to enter 0 seconds if txtSecond1 = ""
     
    If txtSecond1 <> "" Then
        sSeconds = txtSecond1
    Else
        sSeconds = "00"
    End If
    
    If lblCurrentTime.Caption = "Approaching the end of the hour" Then
        lblCurrentTime.ToolTipText = "Double-Click text line if timing for the following hour of programming begins in this hour hour at time " & txtMinute1 & " min " & sSeconds & " sec."
    Else
        lblCurrentTime.ToolTipText = ""
    End If
End Sub

Private Sub lblCurrentTime_DblClick()

    Dim sSeconds As String 'to enter 0 seconds if txtSecond1 = ""
    
    If txtSecond1 <> "" Then
        sSeconds = txtSecond1
    Else
        sSeconds = "00"
    End If
    
    If miTotalRemain <= -60 Then
        MsgBox "What you have programed runs beyond the next hour", vbOKOnly, "Program Runover"
        Exit Sub
    End If

    If mCurrentTime = 0 And (txtMinute1 <> "" Or txtSecond1 <> "") Then 'And lblTotalS.Visible = True Then
        lblCurrentTime.Alignment = 1 'left
        lblCurrentTime.Caption = "The current program timing began within the previous hour at:"
        lblCurrentTime.ForeColor = vbRed
       ' lblHour.Caption = Hour(Time) & ":"
        mCurrentTime = 1
               
        If frmPlanner!mnuLinkTimeRemain.Checked = True Then
            frmPlanner!staStatus.Panels(2) = "Timing began in previous hour at time " & txtMinute1 & " min " & sSeconds & " sec."
        End If
        
    ElseIf mCurrentTime = 1 Then
    
        If txtMinute1 < "55" Then
            lblCurrentTime.Alignment = 0 'left
            lblCurrentTime.ForeColor = vbBlack
            lblCurrentTime.Caption = " 1. Click the 'Set Current Time' button to enter the current time:"
            
            If frmPlanner!mnuLinkTimeRemain.Checked = True Then
                frmPlanner!staStatus.Panels(2) = "* Run time began at " & txtMinute1 & " min " & sSeconds & " sec past the hour."
            End If
        Else
            lblCurrentTime.Alignment = 1 'right
            lblCurrentTime.ForeColor = &HFF0000    'blue
            lblCurrentTime.Caption = "Approaching the end of the hour"
        End If
      '  lblHour.Caption = Hour(Time) & ":"
        mCurrentTime = 0
    End If
    
    pAdd
End Sub

Private Sub Label9_DblClick()
    txtMinute3 = ""
    txtSecond3 = ""
    
    If chkAnnounce.Value = 0 Then
        'Label9.Alignment = 0
        Label9.Caption = "You can replace the program's estimated announce time with your estimate of the announce time you will need:"
    ElseIf chkAnnounce.Value = 1 Then
        'Label9.Alignment = 2
        Label9.Caption = "If an estimated announce time is desired, enter the total estimate in minutes and seconds"
    End If
    
    pSetFocus
End Sub

Private Sub lblExport_dblClick()
    lbl1.BorderStyle = 0
    lbl2.BorderStyle = 0
    lbl3.BorderStyle = 0
    lbl4.BorderStyle = 0
    lbl5.BorderStyle = 0
    lbl6.BorderStyle = 0
    lbl7.BorderStyle = 0
    lbl8.BorderStyle = 0
    mnuToolsExportLineCopy_Click
End Sub

Private Sub lblHourAdj_DblClick()

    If iHourNow = Hour(Now) Then  'cycle hour between hour plus or minus 1
        iHourNow = Hour(Now) + 1
    ElseIf iHourNow = Hour(Now) + 1 Then
        iHourNow = Hour(Now) - 1
    ElseIf iHourNow = Hour(Now) - 1 Then
        iHourNow = Hour(Now)
    ElseIf iHourNow < Hour(Now) - 1 Or iHourNow > Hour(Now) + 1 Then
        iHourNow = Hour(Now)
    End If
    
    Dim Y As Integer
    
    If iHourNow <= 12 Then
        Y = iHourNow
    ElseIf iHourNow > 12 Then
        Y = iHourNow - 12
    End If
    
    lblHourAdj = Y
End Sub

Private Sub lblLinked_DblClick()

    'breaks link between Planner & Time Remain forms
    mnuFileSave_Click
    mnuToolsExportLineups.Enabled = True
    
    lblLinked.Visible = False
    shpLink.Visible = False
    F4Link = 0

    frmTimeRemain!cmdRestoreEntries.Enabled = True
    cmdRestoreEntries.BackColor = &H8000000F 'gray
    cmdRestoreEntries.ToolTipText = ""
    
    mnuToolsExportLineCopy.Checked = False
       
    frmPlanner!mnuLinkTimeRemain.Checked = False
    frmPlanner!mnuLinkTimeRemain.Caption = "&Link Time Remain Page Lineup to this Lineup"
    frmPlanner!staStatus.Panels(2) = ""
    frmPlanner!staStatus.Panels(3) = ""
    frmPlanner!staStatus.Panels(4) = "F4 to Link Lineup to Time Remain Page"
   
    frmPlanner!txtSpots.Visible = False
    frmPlanner!Shape1.Visible = False
    frmPlanner!lblSpotLength.Visible = False
    frmPlanner!lblDate2.Visible = True
    frmPlanner!lblAnnc.Visible = False
    frmPlanner!fraAnnc.Visible = False
    giTimesDiffer = 0

    frmPlanner!lblProgramTime.Visible = False
    frmPlanner!fraAnnc.ForeColor = &H80000008 'black
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, txtBackAnnc, txtAnnc4, txtAnnc5, txtAnnc6, txtAnnc7, txtAnnc8, txtAnnc9, txtAnnc10, txtAnnc11
    Close #501
    
    frmPlanner!txtAnnc1.Text = ""
    frmPlanner!txtAnnc2.Text = ""
    frmPlanner!txtAnnc3.Text = ""
    frmPlanner!txtAnnc4.Text = ""
    frmPlanner!txtAnnc5.Text = ""
    frmPlanner!txtAnnc6.Text = ""
    frmPlanner!txtAnnc7.Text = ""
    frmPlanner!txtAnnc8.Text = ""
    
    txtMinute4.ForeColor = &H80000012 'black
    txtMinute5.ForeColor = &H80000012
    txtMinute6.ForeColor = &H80000012
    txtMinute7.ForeColor = &H80000012
    txtMinute8.ForeColor = &H80000012
    txtMinute9.ForeColor = &H80000012
    txtMinute10.ForeColor = &H80000012
    txtMinute11.ForeColor = &H80000012
    txtSecond4.ForeColor = &H80000012
    txtSecond5.ForeColor = &H80000012
    txtSecond6.ForeColor = &H80000012
    txtSecond7.ForeColor = &H80000012
    txtSecond8.ForeColor = &H80000012
    txtSecond9.ForeColor = &H80000012
    txtSecond10.ForeColor = &H80000012
    txtSecond11.ForeColor = &H80000012
    
    frmPlanner!txtMinute1.ForeColor = &H80000012 'black
    frmPlanner!txtMinute2.ForeColor = &H80000012
    frmPlanner!txtMinute3.ForeColor = &H80000012
    frmPlanner!txtMinute4.ForeColor = &H80000012
    frmPlanner!txtMinute5.ForeColor = &H80000012
    frmPlanner!txtMinute6.ForeColor = &H80000012
    frmPlanner!txtMinute7.ForeColor = &H80000012
    frmPlanner!txtMinute8.ForeColor = &H80000012
    frmPlanner!txtSecond1.ForeColor = &H80000012
    frmPlanner!txtSecond2.ForeColor = &H80000012
    frmPlanner!txtSecond3.ForeColor = &H80000012
    frmPlanner!txtSecond4.ForeColor = &H80000012
    frmPlanner!txtSecond5.ForeColor = &H80000012
    frmPlanner!txtSecond6.ForeColor = &H80000012
    frmPlanner!txtSecond7.ForeColor = &H80000012
    frmPlanner!txtSecond8.ForeColor = &H80000012
    
On Error GoTo HandleErrors

    Close #501
    Open "AnncTime.dat" For Input As #501
    Input #501, BackAnnc, Annc4, Annc5, Annc6, Annc7, Annc8, Annc9, Annc10, Annc11
    Close #501
    
    txtBackAnnc = BackAnnc
    txtAnnc4 = Annc4
    txtAnnc5 = Annc5
    txtAnnc6 = Annc6
    txtAnnc7 = Annc7
    txtAnnc8 = Annc8
    txtAnnc9 = Annc9
    txtAnnc10 = Annc10
    txtAnnc11 = Annc11
    
    If mnuToolsExportLineCopy.Checked = True Then
        mnuToolsExportLineCopy_Click
    End If
HandleErrors:
    Close #501
End Sub

Private Sub lblMinutesLeft_Change()

'--------test 5-21-2015
    Dim rightnow
    rightnow = Now

    Dim W As Integer
    Dim V As Integer
    
    W = Minute(rightnow) + Val(txtMinAdj)
    
    If Val(txtMinAdj) >= 0 Then
    
        If W >= 60 Then
            V = iHourNow + 1
        Else
            V = iHourNow
        End If
    
    ElseIf Val(txtMinAdj) < 0 Then
    
        iiMinute1 = Minute(Time) + Val(txtMinAdj)
        
        If iiMinute1 < 0 Then
            V = iHourNow - 1
'        End If
         Else
            V = iHourNow
        End If
    
    End If
'---------

    If V <= 12 Then
        V = V
    ElseIf V > 12 Then
        V = V - 12
    End If
    
    lblHourAdj = V
    
End Sub

Private Sub lblRemain30_Change()
    If chkAnnounce.Value = 0 Then
    
        If txtMinute2 <> "" Or txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
        Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "" Then
    
            If lblRemain30 = "" Then
                lblRemain30.Visible = False
                lblSpots.Caption = "Enter the number of (" & Val(txtSpotLength) & _
                "-second average time) spot, promo, PSA, weather, etc. inserts REMAINING in the hour"
            Else
                lblRemain30.Visible = True
                lblSpots.Caption = "Enter the number of (" & Val(txtSpotLength) & _
                "-second average time) spot, promo, PSA, weather, etc. inserts REMAINING in the hour (or half-hour) time period"
            End If
            
        Else
            lblSpots.Caption = "Enter the number of (" & Val(txtSpotLength) & _
            "-second average time) spot, promo, PSA, weather, etc. inserts scheduled in the hour (or half-hour) time period"
            Exit Sub
            
        End If
            
    ElseIf chkAnnounce.Value = 1 Then
        lblRemain30.Visible = False
    End If
End Sub

Private Sub lblS_DblClick()
    If txtSpotsS > "1" Then
        txtSpotsS = Format((Val(txtSpotsS) - 1), "##")
    Else
        txtSpotsS = ""
    End If
End Sub

Private Sub lblSpots_dblClick()
    txtSpotsS = ""
End Sub

Private Sub lblSspot_DblClick()
    If txtSpotsS > "1" Then
        txtSpotsS = Format((Val(txtSpotsS) - 1), "##")
    Else
        txtSpotsS = ""
    End If
End Sub

Private Sub lblTimer_Change()
    If lblTimer.Caption = "" Then
        lblTimer.BackColor = &H8000000F
    Else
        lblTimer.BackColor = &HFFFFFF
    End If
End Sub

Private Sub lblTimer_DblClick()
    frmStopWatch.Show
    cmdStopwatch.Enabled = False
End Sub

Private Sub lblTotalS_Change()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
 
    If frmPlanner!mnuLinkTimeRemain.Checked = True And frmTimeRemain.Visible = True Then
 
        If txtMinute4 = frmPlanner!txtMinute1 And txtMinute5 = frmPlanner!txtMinute2 And txtMinute6 = frmPlanner!txtMinute3 _
        And txtMinute7 = frmPlanner!txtMinute4 And txtMinute8 = frmPlanner!txtMinute5 And txtMinute9 = frmPlanner!txtMinute6 _
        And txtMinute10 = frmPlanner!txtMinute7 And txtMinute11 = frmPlanner!txtMinute8 Then
        
            giTimesDiffer = 0
            cmdRestoreEntries.BackColor = &H8000000F 'gray
            cmdRestoreEntries.ToolTipText = ""
            frmTimeRemain!cmdRestoreEntries.Enabled = False
            shpLink.BackColor = &HC000&
            shpLink.BorderColor = &HC000&
            
        Else
            giTimesDiffer = 1
            frmTimeRemain!cmdRestoreEntries.Enabled = True
            shpLink.BackColor = vbRed
            shpLink.BorderColor = vbRed
        End If
    End If
    
    If lblTotalS.Visible = True Then
        Label6.ForeColor = &H80000008  'black
    End If
    
    pAnnounce
End Sub

Private Sub mnuCopyMusicLogList_Click()
    If frmPlanner!txtComposer1 <> "" Or frmPlanner!txtComposer2 <> "" Or frmPlanner!txtComposer3 <> "" Or frmPlanner!txtComposer4 <> "" Or _
    frmPlanner!txtComposer5 <> "" Or frmPlanner!txtComposer6 <> "" Or frmPlanner!txtComposer7 <> "" Or frmPlanner!txtComposer8 <> "" Then

        txtComposer4 = frmPlanner!txtComposer1
        txtMinute4 = frmPlanner!txtMinute1
        txtSecond4 = frmPlanner!txtSecond1
           
        txtComposer5 = frmPlanner!txtComposer2
        txtMinute5 = frmPlanner!txtMinute2
        txtSecond5 = frmPlanner!txtSecond2
        
        txtComposer6 = frmPlanner!txtComposer3
        txtMinute6 = frmPlanner!txtMinute3
        txtSecond6 = frmPlanner!txtSecond3
        
        txtComposer7 = frmPlanner!txtComposer4
        txtMinute7 = frmPlanner!txtMinute4
        txtSecond7 = frmPlanner!txtSecond4
        
        txtComposer8 = frmPlanner!txtComposer5
        txtMinute8 = frmPlanner!txtMinute5
        txtSecond8 = frmPlanner!txtSecond5
        
        txtComposer9 = frmPlanner!txtComposer6
        txtMinute9 = frmPlanner!txtMinute6
        txtSecond9 = frmPlanner!txtSecond6
        
        txtComposer10 = frmPlanner!txtComposer7
        txtMinute10 = frmPlanner!txtMinute7
        txtSecond10 = frmPlanner!txtSecond7
        
        txtComposer11 = frmPlanner!txtComposer8
        txtMinute11 = frmPlanner!txtMinute8
        txtSecond11 = frmPlanner!txtSecond8
        
    End If
     
End Sub

Private Sub mnuOptionsTime_Click()
    If mnuOptionsTime.Checked = True Then
        mnuOptionsTime.Checked = False
        fraIntro.Visible = False
    Else
        mnuOptionsTime.Checked = True
        fraIntro.Visible = True
        txtIntroSetting.SetFocus
        Label24.ForeColor = &HFF0000 'blue
        Frame8.ForeColor = &HFF0000
        Label26.ForeColor = &HFF0000
    End If
End Sub

Private Sub mnuPageAddTime_Click()
    frmAddTime.Show
End Sub

Private Sub mnuPagePlanner_Click()
    giClockShow = 5
    frmPlanner.Show
    frmTimeRemain.Hide
End Sub

Private Sub mnuPagePrevious_Click()

    If giClockShow <> 0 Then
        Select Case giClockShow
            Case 4
                frmPlanner.Show
            Case 3
                frmTransmitter.Show
                frmTransmitter!cmdPrevious.Caption = "&Return to Previous Page F6"
            Case Else
                frmPlanner.Show
        End Select
        frmTimeRemain.Hide
        giClockShow = 5
    Else
        frmPlanner.Show
        frmTimeRemain.Hide
        giClockShow = 5
    End If
    
End Sub

Private Sub mnuPageStopWatch_Click()
    frmStopWatch.Show
    cmdStopwatch.Enabled = False
End Sub

Private Sub mnuPageXmitter_Click()
    frmTransmitter!cmdPrevious.Caption = "&Return to Previous Page F6"
    giClockShow = 5
    frmTransmitter.Show
    frmTimeRemain.Hide
End Sub

Private Sub pAnnounce()

    Dim iAnnounceTime As Currency
    Dim iCal As Currency
    
    Dim iSpotRS As Integer
        
    If txtSpotsS.Visible = True And txtSpotsS <> "" Then
        iSpotRS = (Val(txtSpotsS) * Val(txtSpotLength))
    ElseIf txtSpotsS = "" Then
        iSpotRS = 0
    End If

    If chkAnnounce.Value = 0 Then
    
'----------Label to reset announce times to intro/out time -------------------------

    If (txtAnnc5 = txtIntro Or txtAnnc5 = "") And (txtAnnc6 = txtIntro Or txtAnnc6 = "") _
        And (txtAnnc7 = txtIntro Or txtAnnc7 = "") And (txtAnnc8 = txtIntro Or txtAnnc8 = "") And (txtAnnc9 = txtIntro Or txtAnnc9 = "") _
        And (txtAnnc10 = txtIntro Or txtAnnc10 = "") And (txtAnnc11 = txtIntro Or txtAnnc11 = "") Then

      lblAnncTime.BorderStyle = 0
      Label27 = "Min"
      Label28 = "Sec"
 
    ElseIf txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" Or txtMinute8 <> "" Or txtMinute9 <> "" _
        Or txtMinute10 <> "" Or txtMinute11 <> "" Or txtSecond5 <> "" Or txtSecond6 <> "" Or txtSecond7 <> "" _
        Or txtSecond8 <> "" Or txtSecond9 <> "" Or txtSecond10 <> "" Or txtSecond11 <> "" Then
    
        lblAnncTime = " Click to reset Annc Times to " & txtIntro & " sec "
        lblAnncTime.BorderStyle = 1 'black line border
        Label27 = ""
        Label28 = ""
            
    Else
        lblAnncTime = " Annc Time"
        lblAnncTime.BorderStyle = 0
        Label27 = "Min"
        Label28 = "Sec"
    End If
        
'--------------------Manual entry announce time De-activated---------------------------
   
    If txtMinute3 = "" And txtSecond3 = "" Then
        'controlling visibilities
        'visibilities with no music lineup entries
        If txtMinute2 = "" And txtMinute4 = "" And txtMinute5 = "" And txtMinute6 = "" And txtMinute7 = "" _
            And txtMinute8 = "" And txtMinute9 = "" And txtMinute10 = "" And txtMinute11 = "" Then
            miAnncTime = 0
            txtMinute3 = ""
            txtSecond3 = ""
            txtMinute3.Visible = False
            txtSecond3.Visible = False
            lblS.Visible = False
            lblMinSec3Div.Visible = False
            shpTime3.Visible = False
            lblAnncMin.Visible = False
            lblAnncSec.Visible = False
            Label9.Visible = False
            imgMic.Visible = False
        
        'visibilities with music lineup entries
        ElseIf txtMinute2 <> "" Or txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
            Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "" Then
            txtMinute3.Visible = True
            txtSecond3.Visible = True
            
    If txtSpotsS <> "" Then
        lblS.Visible = True
     End If
            lblMinSec3Div.Visible = True
            shpTime3.Visible = True
            lblAnncMin.Visible = True
            lblAnncSec.Visible = True
            Label9.Visible = True
            imgMic.Visible = True
        End If
 End If
'-------------Adds announce times to music times

        iAnncSum = Val(iiAnnc4) + Val(iiAnnc5) + Val(iiAnnc6) + Val(iiAnnc7) + Val(iiAnnc8) + Val(iiAnnc9) + Val(iiAnnc10) + Val(iiAnnc11) + (Val(txtMinute3) * 60) + Val(txtSecond3)
        '-------xxx
        If txtSpotsS <> "" Then 'if SpotsS has entry, computes from it
            mcAnncTimez = (Val(txtSpotsS) * Val(txtSpotLength)) + iAnncSum
            giSpots = Val(txtSpotsS)
        Else
            mcAnncTimez = (Val(txtSpotsS) * Val(txtSpotLength)) + iAnncSum
            giSpots = Val(txtSpotsS)
        End If
   
'-----------------Music lineup with the various time/CD combinations---------------------
    If txtMinute3 = "" And txtSecond3 = "" Then
  
        Dim iMinute2 As Integer
        iMinute2 = Val(txtMinute2)
        
        Dim iSecond2 As Integer
        iSecond2 = Val(txtSecond2)
        
        If (iMinute2 > 0 Or iSecond2 > 0) And (txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
        Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "") Then

            If RecallControl = 0 Then
                If miBackAnnc = 0 Then 'back announce current CD
                    miAnncTime = (Val((txtIntro) / 2) - 10) + (Val(txtSpotsS) * Val(txtSpotLength)) + iAnncSum
                ElseIf miBackAnnc = 1 Then 'will not back announce current CD
                    miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + iAnncSum + Val(txtBackAnnc)
                End If
                  
            ElseIf RecallControl = 1 Then
                If miBackAnnc = 0 Then
                    miAnncTime = (Val((txtIntro) / 2) - 10) + (Val(txtSpotsS) * Val(txtSpotLength)) + mRAnncSum
                ElseIf miBackAnnc = 1 Then
                    miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + mRAnncSum + Val(txtBackAnnc)
                End If
            End If
        
 '(2) music lineup entries, no time entry, and NO selection played checked
 
        ElseIf (txtMinute1 = "" And txtSecond1 = "") And (txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
            Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "") _
            And (ck4Played.Value = 0 And ck5Played.Value = 0 And ck6Played.Value = 0 And ck7Played.Value = 0 _
            And ck8Played.Value = 0 And ck9Played.Value = 0 And ck10Played.Value = 0 And ck11Played.Value = 0) Then

            If RecallControl = 0 Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + iAnncSum + Val(txtBackAnnc)
            ElseIf RecallControl = 1 Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + mRAnncSum
            End If
        
'(3) music lineup entries, no time entry, but a selection played IS checked

        ElseIf (txtMinute1 = "" And txtSecond1 = "") And (txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
            Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "") _
            And (ck4Played.Value <> 0 Or ck5Played.Value <> 0 Or ck6Played.Value <> 0 Or ck7Played.Value <> 0 _
            Or ck8Played.Value <> 0 Or ck9Played.Value <> 0 Or ck10Played.Value <> 0 Or ck11Played.Value <> 0) Then

            If RecallControl = 0 Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + iAnncSum
            ElseIf RecallControl = 1 Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + mRAnncSum
            End If
        
'(4) time entry + music lineup entries

        ElseIf (txtMinute1 <> "" Or txtSecond1 <> "") And (txtMinute2 = "" And txtSecond2 = "") And (txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" _
            Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "") Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + iAnncSum
                
            If RecallControl = 0 Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + iAnncSum
            ElseIf RecallControl = 1 Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + mRAnncSum
            End If
        
'(5) CD currently playing but no music lineup entries

        ElseIf (txtMinute1 <> "" Or txtSecond1 <> "") And (txtMinute2 <> "" Or txtSecond2 <> "") Then
            If miBackAnnc = 0 Then
                miAnncTime = (Val((txtIntro) / 2) - 10) + (Val(txtSpotsS) * Val(txtSpotLength))
            ElseIf miBackAnnc = 1 Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength)) + Val(txtBackAnnc)
            End If
            
'(6) Nothing listed or playing, no time, txtSpotsS active

        ElseIf txtMinute1 = "" And txtSecond1 = "" And txtMinute4 = "" And txtMinute5 = "" And txtMinute6 = "" _
            And txtMinute7 = "" And txtMinute8 = "" And txtMinute9 = "" And txtMinute10 = "" And txtMinute11 = "" Then
                miAnncTime = (Val(txtSpotsS) * Val(txtSpotLength))  '+ iAnncSum
            End If
            '--------------
            
            RecallControl = 0 'resets control
            
'---------------lblAnnounceTime, label that list the announce time included-----------

            'controling visibility of lblAnnounceTime
            If lblEndTime.Visible = True Or txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" _
                Or txtMinute7 <> "" Or txtMinute8 <> "" Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "" Then
                lblAnnounceTime.Visible = True
                Label20.Visible = True
            ElseIf lblEndTime.Visible = False And txtMinute4 = "" And txtMinute5 = "" And txtMinute6 = "" _
                And txtMinute7 = "" And txtMinute8 = "" And txtMinute9 = "" And txtMinute10 = "" And txtMinute11 = "" Then
                lblAnnounceTime.Visible = False
                Label20.Visible = False
            End If
       
    '-------lblAnnounceTime text

            iCal = miAnncTime / 60
            miCalSec = (iCal - Int(iCal)) * 60 'miCalSec, announce time seconds
            miCalMin = Int(iCal) 'miCalMin, announce time minutes
  
            If (txtSpotsS.Visible And txtSpotsS <> "" And txtSpotsS <> "0") Then
                lblAnnounceTime.Caption = "  • Including estimated spot && music announce times of " _
                & miCalMin & " min " & Format$(miCalSec, "#0") & " sec   "
            Else
                If Val(lblEndTime) = Val(lblTotalS) Then
                    lblAnnounceTime.Caption = "  • Including current CD back-announce time of " _
                    & miCalMin & " min " & Format$(miCalSec, "#0") & " sec   "
                Else
                    lblAnnounceTime.Caption = "  • Including current estimated announce time of " _
                    & miCalMin & " min " & Format$(miCalSec, "#0") & " sec   "
                End If
            End If

'------------------MANUAL ENTRY------------------------

        '----Manual Entry, chkAnnounce.Value = 0 (not actisvated)
        ElseIf txtMinute3 <> "" Or txtSecond3 <> "" Then
            
            Dim iManAnnc As Currency
            Dim iManCal As Currency
            Dim iMinute3 As Integer
            Dim iSecond3 As Integer
            Dim cSpots As Currency
            
            If txtMinute3 = "" Then
                iMinute3 = 0
            Else
                iMinute3 = Val(txtMinute3)
            End If
            
            If txtSecond3 = "" Then
                iSecond3 = 0
            Else
                iSecond3 = Val(txtSecond3)
            End If

            If txtSpotsS.Visible = True Then
                cSpots = Val(txtSpotsS)
            End If
            
            iManAnnc = ((cSpots * Val(txtSpotLength)) + iSecond3) / 60 + iMinute3 'converts to minutes
             
            iManSec = (iManAnnc - Int(iManAnnc)) * 60
            iManMin = Int(iManAnnc)
        
            miAnncTime = 0
            
            If txtMinute3 = "" And txtSpotsS = "" Then
                lblAnnounceTime.Caption = "  • Programmed: " & lblTotalS & " music and  " & iManSec & " sec announce time   "
            Else
                lblAnnounceTime.Caption = "  • Programmed: " & lblTotalS & " music and  " & iManMin & " min " & iManSec & " sec announce time   "
            End If
  '• Ending time
   
            lblAverageTime.Visible = False
            txtIntro.Visible = False
            Frame12.Height = 750
            Label20.Visible = False
            txtAnnc4.Visible = False
            txtAnnc5.Visible = False
            txtAnnc6.Visible = False
            txtAnnc7.Visible = False
            txtAnnc8.Visible = False
            txtAnnc9.Visible = False
            txtAnnc10.Visible = False
            txtAnnc11.Visible = False
            Label22.Visible = False
            lblAnncTime.Visible = False
         End If
    
     Else
     '----Manual Entry, chkAnnounce.Value = 1 (activated)
     
         txtMinute3.Visible = True
         txtSecond3.Visible = True
         
    If txtSpotsS <> "" Then
        lblS.Visible = True
    End If
    
         lblMinSec3Div.Visible = True
         shpTime3.Visible = True
         lblAnncMin.Visible = True
            lblAnncSec.Visible = True
         Label9.Visible = True
         miAnncTime = 0
         lblAnnounceTime.Visible = False

        If chkAnnounce.Value = 1 And (txtMinute3 <> "" Or txtSecond3 <> "") Then
            lblProgramRemain.Visible = True
            lblAnnounceTime.Visible = True
        ElseIf chkAnnounce.Value = 1 And (txtMinute3 = "" And txtSecond3 = "") Then
           lblProgramRemain.Visible = False
        End If

'--------
    Dim xSpots As Integer 'lblAnnounceTime.Caption when manual entry & chkAnnounce.Value = 1 which means no annc time

    If txtSpotsS <> "" And txtSpotsS.Visible = True Then
        xSpots = Val(txtSpotsS)
    Else
        xSpots = 0
    End If

    Dim Xsec As Integer
    Dim Xcal As String
    Dim Xcalmin As Integer
    Dim Xcalsec As Integer
    
    Xsec = (Val(txtMinute3) * 60) + Val(txtSecond3) + xSpots * Val(txtSpotLengthSetting)
    Xcal = Xsec / 60
    Xcalsec = (Xcal - Int(Xcal)) * 60 'announce time seconds
    Xcalmin = Int(Xcal)
 
   lblAnnounceTime.Caption = "  • Including selected announce time of " & Xcalmin & " min " & Format$(Xcalsec, "00") & " sec   "
   
    End If
   pAdd
End Sub

Private Sub mnuFileSave_Click()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If

On Error GoTo HandleErrors
    'no save...exits because nothing to save
    If txtComposer4 = "" And txtComposer5 = "" And txtComposer6 = "" And txtComposer7 = "" _
        And txtComposer8 = "" And txtComposer9 = "" And txtComposer10 = "" And txtComposer11 = "" Then
        MsgBox "There are no lineup entries to save", vbOKOnly, "No Data"
        Exit Sub
    End If
    
   Dim zComposer4, zComposer5, zComposer6, zComposer7, zComposer8, zComposer9, zComposer10, zComposer11 As String
   Dim ziMinute4, ziMinute5, ziMinute6, ziMinute7, ziMinute8, ziMinute9, ziMinute10, ziMinute11, zitestm As Integer
   Dim ziSecond4, ziSecond5, ziSecond6, ziSecond7, ziSecond8, ziSecond9, ziSecond10, ziSecond11, zitests As Integer
   
   'zitestm & zitests are throwaways preventing ziMinutes & ziSeconds from equalling 0 instead of empty. Why? I don't know.
   
   Dim zCD4, zCD5, zCD6, zCD7, zCD8, zCD9, zCD10, zCD11 As String
    
    'saves only if txtComposer contains entry of more than 3 letters. allows for trying different entries.
     
    If Len(txtComposer4) > 2 Then
        zComposer4 = txtComposer4
        ziMinute4 = iMinute4
        ziSecond4 = iSecond4
        zCD4 = txtCD4
    End If
    
    If Len(txtComposer5) > 2 Then
        zComposer5 = txtComposer5
        ziMinute5 = iMinute5
        ziSecond5 = iSecond5
        zCD5 = txtCD5
    End If
    
    If Len(txtComposer6) > 2 Or txtComposer6 = "N/A" Or txtComposer6 = "n/a" Then
        zComposer6 = txtComposer6
        ziMinute6 = iMinute6
        ziSecond6 = iSecond6
        zCD6 = txtCD6
    End If
    
    If Len(txtComposer7) > 2 Then
        zComposer7 = txtComposer7
        ziMinute7 = iMinute7
        ziSecond7 = iSecond7
        zCD7 = txtCD7
    End If
    
    If Len(txtComposer8) > 2 Then
        zComposer8 = txtComposer8
        ziMinute8 = iMinute8
        ziSecond8 = iSecond8
        zCD8 = txtCD8
    End If
    
    If Len(txtComposer9) > 2 Then
        zComposer9 = txtComposer9
        ziMinute9 = iMinute9
        ziSecond9 = iSecond9
        zCD9 = txtCD9
    End If
    
    If Len(txtComposer10) > 2 Then
        zComposer10 = txtComposer10
        ziMinute10 = iMinute10
        ziSecond10 = iSecond10
        zCD10 = txtCD10
    End If
    
    If Len(txtComposer11) > 2 Then
        zComposer11 = txtComposer11
        ziMinute11 = iMinute11
        ziSecond11 = iSecond11
        zCD11 = txtCD11
    End If

    Open "TimeRemain.dat" For Output As #500 'saves even if no data
    Write #500, zComposer4, zComposer5, zComposer6, zComposer7, zComposer8, zComposer9, zComposer10, zComposer11, _
        ; ziMinute4, ziMinute5, ziMinute6, ziMinute7, ziMinute8, ziMinute9, ziMinute10, ziMinute11, _
        ziSecond4, ziSecond5, ziSecond6, ziSecond7, ziSecond8, ziSecond9, ziSecond10, ziSecond11, _
        zCD4, zCD5, zCD6, zCD7, zCD8, zCD9, zCD10, zCD11, spotsS, iAnncSum
    Close #500

    pSetFocus
    Exit Sub
HandleErrors:
    Beep
    MsgBox "Lineup not saved.", vbOKOnly, "Error Saving"
    Close #500
    Close #501
End Sub

Private Sub mnuSetCurrentTime_Click()
    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
   cmdSystemTime_Click
End Sub

Private Sub mnuToolsExportLineCopy_Click()

    If txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" Or txtMinute8 <> "" _
    Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "" Then
    
        If mnuToolsCD.Checked = True Then
            mnuToolsCD_Click
        End If
        
        If mnuToolsExportLineCopy.Checked = True Then
            mnuToolsExportLineCopy.Checked = False
            
            lbl1.BorderStyle = 0
            lbl2.BorderStyle = 0
            lbl3.BorderStyle = 0
            lbl4.BorderStyle = 0
            lbl5.BorderStyle = 0
            lbl6.BorderStyle = 0
            lbl7.BorderStyle = 0
            lbl8.BorderStyle = 0
           
            lblExport.Visible = False
            
            lbl1.MousePointer = 1
            lbl2.MousePointer = 1
            lbl3.MousePointer = 1
            lbl4.MousePointer = 1
            lbl5.MousePointer = 1
            lbl6.MousePointer = 1
            lbl7.MousePointer = 1
            lbl8.MousePointer = 1
            
            lbl1.ForeColor = &H80000008  'black
            lbl2.ForeColor = &H80000008
            lbl3.ForeColor = &H80000008
            lbl4.ForeColor = &H80000008
            lbl5.ForeColor = &H80000008
            lbl6.ForeColor = &H80000008
            lbl7.ForeColor = &H80000008
            lbl8.ForeColor = &H80000008
        Else
            Dim iResponse As Integer
            iResponse = MsgBox("If there already is a Music Planning page lineup and it contains less than 8 items, the selected line will be added at end of the existing lineup." _
            & vbCrLf & vbCrLf & "If the existing Music Planning page lineup is full (8 items), the selected line will replace the contents of the Music Planning page line" _
            & vbCrLf & "which number corresponds to the selected line number." _
            & vbCrLf & vbCrLf & "To continue, double-click the LINE NUMBER (or numbers) of the line(s) to be copied to the Music Planning page lineup." _
            & vbCrLf & vbCrLf & "If this page is linked to the Planning page, this action will break the link.", vbOKCancel + vbInformation, "Reminder")
            If iResponse = vbOK Then
                
                giTxtCopy = 1
                If lblLinked.Visible = True Then
                    lblLinked_DblClick
                End If 'breaks frmPlanner & frmTimeRemain link
        
                mnuToolsExportLineCopy.Checked = True
                
                lblExport.Visible = True
                
                lbl1.MousePointer = 99
                lbl2.MousePointer = 99
                lbl3.MousePointer = 99
                lbl4.MousePointer = 99
                lbl5.MousePointer = 99
                lbl6.MousePointer = 99
                lbl7.MousePointer = 99
                lbl8.MousePointer = 99
                
                If txtComposer4 <> "" Then
                    lbl1.ForeColor = &HFF&       'red
                Else
                    lbl1.ForeColor = &H80000002 'blue
                End If
                
                If txtComposer5 <> "" Then
                    lbl2.ForeColor = &HFF& 'red
                Else
                    lbl2.ForeColor = &H80000002 'blue
                End If
                
                If txtComposer6 <> "" Then
                    lbl3.ForeColor = &HFF&       'red
                Else
                    lbl3.ForeColor = &H80000002 'blue
                End If
                
                If txtComposer7 <> "" Then
                    lbl4.ForeColor = &HFF&       'red
                Else
                    lbl4.ForeColor = &H80000002 'blue
                End If
                
                If txtComposer8 <> "" Then
                    lbl5.ForeColor = &HFF&       'red
                Else
                    lbl5.ForeColor = &H80000002 'blue
                End If
                
                If txtComposer9 <> "" Then
                    lbl6.ForeColor = &HFF&       'red
                Else
                    lbl6.ForeColor = &H80000002 'blue
                End If
                
                If txtComposer10 <> "" Then
                    lbl7.ForeColor = &HFF&       'red
                Else
                    lbl7.ForeColor = &H80000002 'blue
                End If
                
                If txtComposer11 <> "" Then
                    lbl8.ForeColor = &HFF&       'red
                Else
                    lbl8.ForeColor = &H80000002 'blue
                End If
                
                mnuImport.Enabled = True
            End If
        End If
    Else
        MsgBox "There is no Music Lineup to export." & vbCrLf & vbCrLf & "An exported line requirers an entry in the minutes box and also should include a composer entry.", vbOKOnly, "Time Entries Required"
    End If
End Sub

Private Sub mnuToolsExportLineups_Click()
    Dim iResponse As Integer

    If txtMinute4 <> "" Or txtMinute5 <> "" Or txtMinute6 <> "" Or txtMinute7 <> "" Or txtMinute8 <> "" _
    Or txtMinute9 <> "" Or txtMinute10 <> "" Or txtMinute11 <> "" Then
        
        iResponse = MsgBox("The Music Lineup shown on this page will delete and replace all existing data." _
         & vbCrLf & "on the Music Planning page lineup." _
        & vbCrLf & vbCrLf & "Continue with the export?", vbYesNo + vbInformation, "Caution")
        If iResponse = vbYes Then
        
        giTxtCopy = 1
        ck4Played.Value = 0
        ck5Played.Value = 0
        ck6Played.Value = 0
        ck7Played.Value = 0
        ck8Played.Value = 0
        ck9Played.Value = 0
        ck10Played.Value = 0
        ck11Played.Value = 0
        
       ' mnuToolsExportLineups.Checked = True
        
        If mnuToolsExportLineCopy.Checked = True Then
            Call mnuToolsExportLineCopy_Click
        End If
        frmPlanner!txtComposer1 = txtComposer4
        frmPlanner!txtMinute1 = txtMinute4
        frmPlanner!txtSecond1 = txtSecond4
        frmPlanner!txtComposition1 = ""
        frmPlanner!txtDisc1 = ""
        frmPlanner!txtTrack1 = ""
        
        frmPlanner!txtComposer2 = txtComposer5
        frmPlanner!txtMinute2 = txtMinute5
        frmPlanner!txtSecond2 = txtSecond5
        frmPlanner!txtComposition2 = ""
        frmPlanner!txtDisc2 = ""
        frmPlanner!txtTrack2 = ""
        frmPlanner!txtComposer3 = txtComposer6
        frmPlanner!txtMinute3 = txtMinute6
        frmPlanner!txtSecond3 = txtSecond6
        frmPlanner!txtComposition3 = ""
        frmPlanner!txtDisc3 = ""
        frmPlanner!txtTrack3 = ""
        
        frmPlanner!txtComposer4 = txtComposer7
        frmPlanner!txtMinute4 = txtMinute7
        frmPlanner!txtSecond4 = txtSecond7
        frmPlanner!txtComposition4 = ""
        frmPlanner!txtDisc4 = ""
        frmPlanner!txtTrack4 = ""
        
        frmPlanner!txtComposer5 = txtComposer8
        frmPlanner!txtMinute5 = txtMinute8
        frmPlanner!txtSecond5 = txtSecond8
        frmPlanner!txtComposition5 = ""
        frmPlanner!txtDisc5 = ""
        frmPlanner!txtTrack5 = ""
        
        frmPlanner!txtComposer6 = txtComposer9
        frmPlanner!txtMinute6 = txtMinute9
        frmPlanner!txtSecond6 = txtSecond9
        frmPlanner!txtComposition6 = ""
        frmPlanner!txtDisc6 = ""
        frmPlanner!txtTrack6 = ""
        
        frmPlanner!txtComposer7 = txtComposer10
        frmPlanner!txtMinute7 = txtMinute10
        frmPlanner!txtSecond7 = txtSecond10
        frmPlanner!txtComposition7 = ""
        frmPlanner!txtDisc7 = ""
        frmPlanner!txtTrack7 = ""
        
        frmPlanner!txtComposer8 = txtComposer11
        frmPlanner!txtMinute8 = txtMinute11
        frmPlanner!txtSecond8 = txtSecond11
        frmPlanner!txtComposition8 = ""
        frmPlanner!txtDisc8 = ""
        frmPlanner!txtTrack8 = ""
        mnuImport.Enabled = True
        
        If mnuToolsCD.Checked = True Then
            frmPlanner!txtCD1 = txtCD4
            frmPlanner!txtCD2 = txtCD5
            frmPlanner!txtCD3 = txtCD6
            frmPlanner!txtCD4 = txtCD7
            frmPlanner!txtCD5 = txtCD8
            frmPlanner!txtCD6 = txtCD9
            frmPlanner!txtCD7 = txtCD10
            frmPlanner!txtCD8 = txtCD11
        Else
            frmPlanner!txtCD1 = ""
            frmPlanner!txtCD2 = ""
            frmPlanner!txtCD3 = ""
            frmPlanner!txtCD4 = ""
            frmPlanner!txtCD5 = ""
            frmPlanner!txtCD6 = ""
            frmPlanner!txtCD7 = ""
            frmPlanner!txtCD8 = ""
        End If
        
        pAdd 'prevents "Time remain and planning times differ" message
        
        ElseIf iResponse = vbNo Then
           ' mnuToolsExportLineups.Checked = False
            Exit Sub
        End If
    Else
        MsgBox "There is no Music Lineup to export." & vbCrLf & vbCrLf & "Exported data requirers entries in at least a minutes box and also should include composer entries.", vbOKOnly, "Time Entries Required"
    End If

End Sub

Private Sub mnuToolsMemos_Click()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If

    If giMemo = 0 Then
        Dim prompt, AccessCode
        prompt = "" & vbCrLf & "Enter Access Code."
        AccessCode = InputBox$(prompt, "Access Code Required")

        If AccessCode = giAccess Then
            frmMemos.Show 'necessary to minimize
            giMemo = giMemo + 1
            mnuToolsMemos.Caption = "Memos..."
            frmPlanner!mnuToolsMemos.Caption = "Memos..."
        ElseIf AccessCode <> giAccess And AccessCode <> "" Then
            MsgBox AccessCode & " is an incorrect access code", vbOKOnly, "Incorrect Code"
        Else
        End If
    Else
        frmMemos.Show 'necessary to minimize
    End If
End Sub

Private Sub mnuToolsPrintPage_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "Time Remain")
    If iResponse = vbNo Then
        Exit Sub
    ElseIf iResponse = vbYes Then
       ' lblProgramRemain.ForeColor = &H80000008 'black
        lblProgramInfo.ForeColor = &H80000008
        Label9.ForeColor = &H80000008
        PrintForm
    End If
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

Private Sub mnuToolsCD_Click()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
    
    If mnuToolsCD.Checked = True Then
       mnuToolsCD.Checked = False 'tabstops activated if unchecked
       
        Label4.ForeColor = &H808080       'Annc label gray  'txtCD4.BackColor = &H80000016
        
        txtCD4.BackColor = &H80000016 'lite gray
        txtCD5.BackColor = &H80000016
        txtCD6.BackColor = &H80000016
        txtCD7.BackColor = &H80000016
        txtCD8.BackColor = &H80000016
        txtCD9.BackColor = &H80000016
        txtCD10.BackColor = &H80000016
        txtCD11.BackColor = &H80000016
                 
        txtCD4.Appearance = 0
        txtCD5.Appearance = 0
        txtCD6.Appearance = 0
        txtCD7.Appearance = 0
        txtCD8.Appearance = 0
        txtCD9.Appearance = 0
        txtCD10.Appearance = 0
        txtCD11.Appearance = 0
        
        txtCD4.BorderStyle = 0
        txtCD5.BorderStyle = 0
        txtCD6.BorderStyle = 0
        txtCD7.BorderStyle = 0
        txtCD8.BorderStyle = 0
        txtCD9.BorderStyle = 0
        txtCD10.BorderStyle = 0
        txtCD11.BorderStyle = 0
         
        txtCD4.TabStop = False
        txtCD5.TabStop = False
        txtCD6.TabStop = False
        txtCD7.TabStop = False
        txtCD8.TabStop = False
        txtCD9.TabStop = False
        txtCD10.TabStop = False
        txtCD11.TabStop = False
        
        txtCD4 = ""
        txtCD5 = ""
        txtCD6 = ""
        txtCD7 = ""
        txtCD8 = ""
        txtCD9 = ""
        txtCD10 = ""
        txtCD11 = ""
        
        txtCD4.Enabled = False
        txtCD5.Enabled = False
        txtCD6.Enabled = False
        txtCD7.Enabled = False
        txtCD8.Enabled = False
        txtCD9.Enabled = False
        txtCD10.Enabled = False
        txtCD11.Enabled = False
        
      Else
        mnuToolsCD.Checked = True 'checked, tabstops deactivated
        
        Label4.ForeColor = &H80&      'Annc label rust
        
        txtCD4.BackColor = &HFFFFFF    'white
        txtCD5.BackColor = &HFFFFFF
        txtCD6.BackColor = &HFFFFFF
        txtCD7.BackColor = &HFFFFFF
        txtCD8.BackColor = &HFFFFFF
        txtCD9.BackColor = &HFFFFFF
        txtCD10.BackColor = &HFFFFFF
        txtCD11.BackColor = &HFFFFFF
              
        txtCD4.TabStop = True
        txtCD5.TabStop = True
        txtCD6.TabStop = True
        txtCD7.TabStop = True
        txtCD8.TabStop = True
        txtCD9.TabStop = True
        txtCD10.TabStop = True
        txtCD11.TabStop = True
        
        txtCD4.Enabled = True
        txtCD5.Enabled = True
        txtCD6.Enabled = True
        txtCD7.Enabled = True
        txtCD8.Enabled = True
        txtCD9.Enabled = True
        txtCD10.Enabled = True
        txtCD11.Enabled = True
        
        txtCD4.Appearance = 1
        txtCD5.Appearance = 1
        txtCD6.Appearance = 1
        txtCD7.Appearance = 1
        txtCD8.Appearance = 1
        txtCD9.Appearance = 1
        txtCD10.Appearance = 1
        txtCD11.Appearance = 1
        
        txtCD4.BorderStyle = 1
        txtCD5.BorderStyle = 1
        txtCD6.BorderStyle = 1
        txtCD7.BorderStyle = 1
        txtCD8.BorderStyle = 1
        txtCD9.BorderStyle = 1
        txtCD10.BorderStyle = 1
        txtCD11.BorderStyle = 1
     End If
End Sub

Private Sub txtConvert_DblClick()
    txtConvert = ""
    Label8 = "enter min ---- then click"
    Frame6.Caption = "Convert min to sec"
    lblMinSec = "min"
    
End Sub

Private Sub txtComposer10_GotFocus()

    If txtComposer10 = "" Then
        txtComposer10 = "?"
    End If

    txtComposer10.SelStart = 0 'begin selection at start
    txtComposer10.SelLength = Len(txtComposer10)
End Sub

Private Sub txtComposer11_GotFocus()

    If txtComposer11 = "" Then
        txtComposer11 = "?"
    End If

    txtComposer11.SelStart = 0 'begin selection at start
    txtComposer11.SelLength = Len(txtComposer11)
End Sub

Private Sub txtComposer4_GotFocus()

    If txtComposer4 = "" Then
    txtComposer4 = "?"
    End If
    
    txtComposer4.SelStart = 0 'begin selection at start
    txtComposer4.SelLength = Len(txtComposer4)
End Sub

Private Sub txtComposer5_GotFocus()

    If txtComposer5 = "" Then
    txtComposer5 = "?"
    End If
    
    txtComposer5.SelStart = 0 'begin selection at start
    txtComposer5.SelLength = Len(txtComposer5)
End Sub

Private Sub txtComposer6_GotFocus()

    If txtComposer6 = "" Then
    txtComposer6 = "?"
    End If
    
    txtComposer6.SelStart = 0 'begin selection at start
    txtComposer6.SelLength = Len(txtComposer6)
End Sub

Private Sub txtComposer7_GotFocus()

    If txtComposer7 = "" Then
        txtComposer7 = "?"
    End If

    txtComposer7.SelStart = 0 'begin selection at start
    txtComposer7.SelLength = Len(txtComposer7)
End Sub

Private Sub txtComposer8_GotFocus()

    If txtComposer8 = "" Then
        txtComposer8 = "?"
    End If

    txtComposer8.SelStart = 0 'begin selection at start
    txtComposer8.SelLength = Len(txtComposer8)
End Sub

Private Sub txtComposer9_GotFocus()

    If txtComposer9 = "" Then
        txtComposer9 = "?"
    End If

    txtComposer9.SelStart = 0 'begin selection at start
    txtComposer9.SelLength = Len(txtComposer9)
End Sub

Private Sub txtConvert_LostFocus()

    If Label8 = "enter min ---- then click" Then
        If Val(txtConvert) > 15 Then
            MsgBox "Max entry is 15 minutes which is 900 seconds", _
            vbOKOnly, "Entry Greater than 15 Minutes"
            txtConvert = ""
            txtConvert.SetFocus
            Exit Sub
        End If
        
        ElseIf Label8 = "enter sec ---- then click" Then
        
        If Val(txtConvert) > 900 Then
            MsgBox "Max entry is 900 seconds which is 15 minutes", _
            vbOKOnly, "Entry Greater than 900 Seconds"
            txtConvert = ""
            txtConvert.SetFocus
            Exit Sub
        End If
     End If
End Sub

Private Sub txtMinAdj_Change()

    If txtMinAdj = "--" Then
        MsgBox txtMinAdj & "   Incorrect. Negative number must be preceeded by a single dash only.", vbOKOnly, "You Have Entered a Double-Dash  " & txtMinAdj
         txtMinAdj = ""
         txtMinAdj.SetFocus
         Exit Sub
    End If

    If Not IsNumeric(txtMinAdj) And txtMinAdj <> "-" And txtMinAdj <> "+" And txtMinAdj <> "" Then
         MsgBox txtMinAdj & "  is an incorrect entry. Entry must be a positive or negative number.", vbOKOnly, "Non-Numeric Entry"
         txtMinAdj = ""
         txtMinAdj.SetFocus
         Exit Sub
    End If
    
    If (Val(txtMinAdj) > 59) Or (Val(txtMinAdj) < -59) Then 'warning entry exceeds 59 minutes
        MsgBox "Entry may not exceed 59 minutes", 0, "Error"
        txtMinAdj = ""
        txtMinAdj.SetFocus
        Exit Sub
    End If

    If Val(txtSecAdj) < 0 And Val(txtMinAdj) > 0 Then
        txtSecAdj = ""
    ElseIf Val(txtSecAdj) > 0 And Val(txtMinAdj) < 0 Then
        txtSecAdj = ""
    End If
    
    If Val(txtMinAdj) > 0 And Val(txtSecAdj) < 0 Then 'if txtSecond11 is negative, changes minute entry to negative
        txtMinAdj = Format$(txtMinAdj, "-##")
    End If
    
    If Val(txtMinAdj) > 0 Or Val(txtSecAdj) > 0 Then
        Label29 = "computer clock slow"
        Label29.ForeColor = vbBlue
        'lblMinutesLeft.ForeColor = &H80000008   'black
        'lblSecondsLeft.ForeColor = &H80000008   'black
    ElseIf Val(txtMinAdj) < 0 Or Val(txtSecAdj) < 0 Then
        Label29 = "computer clock fast"
        Label29.ForeColor = vbBlue
        lblMinutesLeft.ForeColor = &H80000008   'black
        lblSecondsLeft.ForeColor = &H80000008   'black
    Else
        Label29 = "clock error adjustment"
        Label29.ForeColor = vbBlue
        'lblMinutesLeft.ForeColor = &H80&      'rust'&H80000008   'black
        'lblSecondsLeft.ForeColor = &H80&      'rust'&H80000008   'black
        Label29.ForeColor = &H80&
    End If
 
    pAnnounce
End Sub

Private Sub txtMinAdj_DblClick()

    txtMinAdj.Text = ""
    
    If txtMinAdj = "" And txtSecAdj = "" Then
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
    End If
    
End Sub

Private Sub txtMinAdj_GotFocus()
    txtMinAdj.SelStart = 0 'begin selection at start
    txtMinAdj.SelLength = Len(txtMinAdj)
    txtSecAdj.TabStop = True
End Sub

Private Sub txtMinAdj_LostFocus()
    If Val(txtMinAdj) = 0 And Val(txtSecAdj) = 0 Then
        txtSecAdj = ""
    End If
    
    If Val(txtMinAdj) < 0 Then
        txtSecAdj = "-"
    Else
        txtSecAdj = ""
    End If
    
    If Val(txtMinAdj) = 0 Then
        txtMinAdj = ""
    End If
    
    If txtMinAdj <> "" Then
        txtMinAdj.Text = Format(txtMinAdj, "0")
    End If
    
    If txtMinAdj <> "" Or txtSecAdj <> "" Then
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
    Else
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
    End If

End Sub

Private Sub txtSecAdj_Change()

    If txtSecAdj = "--" Then
        MsgBox txtSecAdj & "   Incorrect. Negative number must be preceeded by a single dash only.", vbOKOnly, "You Have Entered a Double-Dash  " & txtSecAdj
         txtSecAdj = ""
         txtSecAdj.SetFocus
         Exit Sub
    End If

    If Not IsNumeric(txtSecAdj) And txtSecAdj <> "-" And txtSecAdj <> "+" And txtSecAdj <> "" Then
        MsgBox txtSecAdj & "  is an incorrect entry. Entry must be a number betweem plus 59 and negative 59 seconds", vbOKOnly, "Non-Numeric Entry"
        txtSecAdj = ""
        txtSecAdj.SetFocus
        Exit Sub
    End If
    
    If (Val(txtSecAdj) > 59) Or (Val(txtSecAdj) < -59) Then 'warning entry exceeds 59 minutes
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecAdj = ""
        txtSecAdj.SetFocus
        Exit Sub
    End If

    If Val(txtMinAdj) < 0 And Val(txtSecAdj) > 0 And Val(txtSecAdj) > 5 Then 'reformats as minus number
         txtSecAdj = Format$(txtSecAdj, "-00")                      '>5 prevents premature change as entry begins
    End If
   
    If Val(txtSecAdj) > 0 Or Val(txtMinAdj) > 0 Then
        Label29 = "computer clock slow"
        Label29.ForeColor = vbBlue
 
    ElseIf Val(txtSecAdj) < 0 Or Val(txtMinAdj) < 0 Then
        Label29 = "computer clock fast"
        Label29.ForeColor = vbBlue
        lblSecondsLeft.ForeColor = &H80000008   'black
        lblMinutesLeft.ForeColor = &H80000008   'black
    Else
        Label29 = "clock error adjustment"
        Label29.ForeColor = vbBlue
        Label29.ForeColor = &H80&
    End If
    
    If txtSecAdj <> "" Then
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
    End If
 
    pAnnounce

End Sub

Private Sub Timer1_Timer()

    Dim Today As Variant
    Today = Now
    lblClock.Caption = Format(Today, "h:mm:ss ampm")
    
     'setting the computer system clock correction factor
    
    If fraAdjustTime.Visible = True Then
        Dim rightnow
        rightnow = Now
   
        Dim vSecAdj As Integer
        Dim vHourAdj As Integer
        Dim HourSecAdj As Integer
        
        If Second(rightnow) + Val(txtSecAdj) >= 0 And Second(rightnow) + Val(txtSecAdj) <= 60 Then
            vSecAdj = Second(rightnow) + Val(txtSecAdj)
            vMinAdj = Minute(rightnow) + Val(txtMinAdj)
            
        ElseIf Second(rightnow) + Val(txtSecAdj) >= 60 Then
            vSecAdj = (Second(rightnow) + Val(txtSecAdj) - 60)
             vMinAdj = (Minute(rightnow) + Val(txtMinAdj)) + 1
             
        ElseIf Second(rightnow) + Val(txtSecAdj) < 0 Then
      
            vSecAdj = (Second(rightnow) + 60 + Val(txtSecAdj))
            vMinAdj = Minute(rightnow) + Val(txtMinAdj) - 1
        End If
    
        If Val(vMinAdj) > 60 Then
            vMinAdj = vMinAdj - 60
        End If
        
        If Val(vMinAdj) < 0 Then
            vMinAdj = 60 + vMinAdj
        End If
        
        If vMinAdj = "60" Then
            vMinAdj = "00"
        End If
        
        If vSecAdj = 60 Then 'timer to show "00" rather than "60"
            vSecAdj = 0
        End If
        
        lblMinutesLeft.Caption = ":" & Format(vMinAdj, "0#")
        lblSecondsLeft.Caption = ":" & Format(vSecAdj, "0#")
    
    End If
    
End Sub

Private Sub txtAnnc10_Change()

    If ck10Played.Value = 0 Then
       If IsNumeric(txtAnnc10) Or txtAnnc10 = "" Or txtAnnc10 = "-" Then
            iAnnc10 = txtAnnc10
            iiAnnc10 = txtAnnc10
            pAnnounce
        Else
            MsgBox "You have entered a character other than a number." _
            & vbCrLf & vbCrLf & "Enter the number of seconds you plan to use announcing and/or back-announcing this item." _
            & vbCrLf & vbCrLf & "If you do not intend to announce/back-announce this selection, enter a zero.", vbOKOnly, "Entry Error"
            txtAnnc10 = ""
            txtAnnc10.SetFocus
            Exit Sub
        End If
    End If
    
    '-------
    If Val(txtAnnc10) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce the current selection." _
        & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & _
        "If no announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Greater than 300 Seconds"
        txtAnnc10 = ""
        txtAnnc10.SetFocus
        Exit Sub
    End If
    '-------
    If frmPlanner!txtComposer7 <> "" Or frmPlanner!txtMinute7 <> "" Then
        If ck10Played.Value = 0 Then
            frmPlanner!txtAnnc7 = txtAnnc10
        Else
            frmPlanner!txtAnnc7 = ""
        End If
    End If
    
    If Val(txtAnnc10) > 15 Then
        txtAnnc10.ToolTipText = "Overwrite to change announce time or double-click to reduce announce time by 15 sec"
    Else
        txtAnnc10.ToolTipText = "Overwrite to change announce time"
    End If

End Sub

Private Sub txtAnnc10_DblClick()

   If Val(txtAnnc10) > 10 Then
        txtAnnc10 = Format((Val(txtAnnc10) - 15), "##")
    End If
     
    txtAnnc10.SelStart = 0 'begin selection at start
    txtAnnc10.SelLength = Len(txtAnnc10)
    txtAnnc10.SetFocus
End Sub

Private Sub txtAnnc10_GotFocus()
    txtAnnc10.SelStart = 0 'begin selection at start
    txtAnnc10.SelLength = Len(txtAnnc10)
End Sub

Private Sub txtAnnc10_LostFocus()

    If txtAnnc10 <> "" And Check1(6).Value = 1 Then
        Check1(6).Value = 0
    End If
    
    If Val(txtAnnc10) < 0 Then
        txtAnnc10 = ""
    End If
    
    If txtAnnc10 <> "" And Val(txtAnnc10) > 0 Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
    End If
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, iBackAnnc, iAnnc4, iAnnc5, iAnnc6, iAnnc7, iAnnc8, iAnnc9, iAnnc10, iAnnc11
    Close #501
End Sub

Private Sub txtAnnc11_Change()

    If ck11Played.Value = 0 Then
       If IsNumeric(txtAnnc11) Or txtAnnc11 = "" Or txtAnnc11 = "-" Then
            iAnnc11 = txtAnnc11
            iiAnnc11 = txtAnnc11
            pAnnounce
        Else
            MsgBox "You have entered a character other than a number." _
            & vbCrLf & vbCrLf & "Enter the number of seconds you plan to use announcing and/or back-announcing this item." _
           & vbCrLf & vbCrLf & "If you do not intend to announce/back-announce this selection, enter a zero.", vbOKOnly, "Entry Error"
            txtAnnc11 = ""
            txtAnnc11.SetFocus
            Exit Sub
        End If
    End If
    
    '-------
    If Val(txtAnnc11) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce the current selection." _
        & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & _
        "If no announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Greater than 300 Seconds"
        txtAnnc11 = ""
        txtAnnc11.SetFocus
        Exit Sub
    End If
    '-------
    If frmPlanner!txtComposition8 <> "" Or frmPlanner!txtMinute8 <> "" Then 'And frmPlanner!chkNonCD8.Value = 0 Then
        If ck11Played.Value = 0 Then
            frmPlanner!txtAnnc8 = txtAnnc11
        Else
            frmPlanner!txtAnnc8 = ""
        End If
    End If
    
    If Val(txtAnnc11) > 15 Then
        txtAnnc11.ToolTipText = "Overwrite to change announce time or double-click to reduce announce time by 15 sec"
    Else
        txtAnnc11.ToolTipText = "Overwrite to change announce time"
    End If
End Sub

Private Sub txtAnnc11_DblClick()

   If Val(txtAnnc11) > 10 Then
        txtAnnc11 = Format((Val(txtAnnc11) - 15), "##")
    End If
     
    txtAnnc11.SelStart = 0 'begin selection at start
    txtAnnc11.SelLength = Len(txtAnnc11)
    txtAnnc11.SetFocus
End Sub

Private Sub txtAnnc11_GotFocus()
    txtAnnc11.SelStart = 0 'begin selection at start
    txtAnnc11.SelLength = Len(txtAnnc11)
End Sub

Private Sub txtAnnc11_LostFocus()

    If txtAnnc11 <> "" And Check1(7).Value = 1 Then
        Check1(7).Value = 0
    End If
    
    If Val(txtAnnc11) < 0 Then
        txtAnnc11 = ""
    End If
    
    If txtAnnc11 <> "" And Val(txtAnnc11) > 0 Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
    End If
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, iBackAnnc, iAnnc4, iAnnc5, iAnnc6, iAnnc7, iAnnc8, iAnnc9, iAnnc10, iAnnc11
    Close #501
End Sub

Private Sub txtAnnc4_Change()

    If ck4Played.Value = 0 Then
       If IsNumeric(txtAnnc4) Or txtAnnc4 = "" Then
            iAnnc4 = txtAnnc4 'for file
            iiAnnc4 = txtAnnc4 'for computation
            pAnnounce
        Else
            MsgBox "You have entered a character other than a number." _
            & vbCrLf & vbCrLf & "Enter the number of seconds you plan to use announcing and/or back-announcing this item." _
            & vbCrLf & vbCrLf & "If you do not intend to announce/back-announce this selection, enter a zero.", vbOKOnly, "Entry Error"
            txtAnnc4 = ""
            txtAnnc4.SetFocus
            Exit Sub
        End If
    End If
    
    '-------
    If Val(txtAnnc4) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce the current selection." _
        & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & _
        "If no announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Greater than 300 Seconds"
        txtAnnc4 = ""
        txtAnnc4.SetFocus
        Exit Sub
    End If
    '-------
    If frmPlanner!txtComposer1 <> "" Or frmPlanner!txtMinute1 <> "" Then
        If ck4Played.Value = 0 Then
            frmPlanner!txtAnnc1 = txtAnnc4
        Else
            frmPlanner!txtAnnc1 = ""
        End If
    End If
    
    If Val(txtAnnc4) > 15 Then
        txtAnnc4.ToolTipText = "Overwrite to change announce time or double-click to reduce announce time by 15 sec"
    Else
        txtAnnc4.ToolTipText = "Overwrite to change announce time"
    End If
End Sub

Private Sub txtAnnc4_DblClick()

   If Val(txtAnnc4) > 10 Then
        txtAnnc4 = Format((Val(txtAnnc4) - 15), "##")
    End If
    
    If txtAnnc4.Enabled = True Then
       If txtAnnc4 <> Val(txtIntro) And (txtBackAnnc = "" Or txtBackAnnc = "0") Then
           lblAnncTime = " Click to reset Annc Times to " & txtIntro & " sec "
       End If
    End If

    txtAnnc4.SelStart = 0 'begin selection at start
    txtAnnc4.SelLength = Len(txtAnnc4)
    txtAnnc4.SetFocus
   
    
End Sub

Private Sub txtAnnc4_GotFocus()
    txtAnnc4.SelStart = 0 'begin selection at start
    txtAnnc4.SelLength = Len(txtAnnc4)
End Sub

Private Sub txtAnnc4_LostFocus()

    If txtAnnc4 <> "" And Check1(0).Value = 1 Then
        Check1(0).Value = 0
    End If
    
    If txtAnnc4 <> "" Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
    End If
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, iBackAnnc, iAnnc4, iAnnc5, iAnnc6, iAnnc7, iAnnc8, iAnnc9, iAnnc10, iAnnc11
    Close #501
End Sub

Private Sub txtAnnc5_Change()

    If ck5Played.Value = 0 Then
        If IsNumeric(txtAnnc5) Or txtAnnc5 = "" Then
            iAnnc5 = txtAnnc5
            iiAnnc5 = txtAnnc5
            pAnnounce
        Else
            MsgBox "You have entered a character other than a number." _
            & vbCrLf & vbCrLf & "Enter the number of seconds you plan to use announcing and/or back-announcing this item." _
            & vbCrLf & vbCrLf & "If you do not intend to announce/back-announce this selection, enter a zero.", vbOKOnly, "Entry Error"
            txtAnnc5 = ""
            txtAnnc5.SetFocus
            Exit Sub
        End If
    End If
    
    '-------
    If Val(txtAnnc5) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce the current selection." _
        & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & _
        "If no announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Greater than 300 Seconds"
        txtAnnc5 = ""
        txtAnnc5.SetFocus
        Exit Sub
    End If
    '-------
    
    If frmPlanner!txtComposer2 <> "" Or frmPlanner!txtMinute2 <> "" Then
        If ck5Played.Value = 0 Then
            frmPlanner!txtAnnc2 = txtAnnc5
        Else
            frmPlanner!txtAnnc2 = ""
        End If
    End If
    
    If Val(txtAnnc5) > 15 Then
        txtAnnc5.ToolTipText = "Overwrite to change announce time or double-click to reduce announce time by 15 sec"
    Else
        txtAnnc5.ToolTipText = "Overwrite to change announce time"
    End If
End Sub

Private Sub txtAnnc5_DblClick()

   If Val(txtAnnc5) > 10 Then
        txtAnnc5 = Format((Val(txtAnnc5) - 15), "##")
    End If
     
    txtAnnc5.SelStart = 0 'begin selection at start
    txtAnnc5.SelLength = Len(txtAnnc5)
    txtAnnc5.SetFocus
End Sub

Private Sub txtAnnc5_GotFocus()
    txtAnnc5.SelStart = 0 'begin selection at start
    txtAnnc5.SelLength = Len(txtAnnc5)
End Sub

Private Sub txtAnnc5_LostFocus()

    If txtAnnc5 <> "" And Check1(1).Value = 1 Then
        Check1(1).Value = 0
    End If
    
    If txtAnnc5 <> "" Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
    End If
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, iBackAnnc, iAnnc4, iAnnc5, iAnnc6, iAnnc7, iAnnc8, iAnnc9, iAnnc10, iAnnc11
    Close #501
End Sub

Private Sub txtAnnc6_Change()

    If ck6Played.Value = 0 Then
       If IsNumeric(txtAnnc6) Or txtAnnc6 = "" Then
            iAnnc6 = txtAnnc6
            iiAnnc6 = txtAnnc6
            pAnnounce
        Else
            MsgBox "You have entered a character other than a number." _
            & vbCrLf & vbCrLf & "Enter the number of seconds you plan to use announcing and/or back-announcing this item." _
            & vbCrLf & vbCrLf & "If you do not intend to announce/back-announce this selection, enter a zero.", vbOKOnly, "Entry Error"
            txtAnnc6 = ""
            txtAnnc6.SetFocus
            Exit Sub
        End If
    End If
    
    '-------
    If Val(txtAnnc6) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce the current selection." _
        & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & _
        "If no announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Greater than 300 Seconds"
        txtAnnc6 = ""
        txtAnnc6.SetFocus
        Exit Sub
    End If
    '-------
    If frmPlanner!txtComposer3 <> "" Or frmPlanner!txtMinute3 <> "" Then
        If ck6Played.Value = 0 Then
            frmPlanner!txtAnnc3 = txtAnnc6
        Else
            frmPlanner!txtAnnc3 = ""
        End If
    End If
    
    If Val(txtAnnc6) > 15 Then
        txtAnnc6.ToolTipText = "Overwrite to change announce time or double-click to reduce announce time by 15 sec"
    Else
        txtAnnc6.ToolTipText = "Overwrite to change announce time"
    End If
End Sub

Private Sub txtAnnc6_DblClick()

   If Val(txtAnnc6) > 10 Then
        txtAnnc6 = Format((Val(txtAnnc6) - 15), "##")
    End If
     
    txtAnnc6.SelStart = 0 'begin selection at start
    txtAnnc6.SelLength = Len(txtAnnc6)
    txtAnnc6.SetFocus
End Sub

Private Sub txtAnnc6_GotFocus()
    txtAnnc6.SelStart = 0 'begin selection at start
    txtAnnc6.SelLength = Len(txtAnnc6)
End Sub

Private Sub txtAnnc6_LostFocus()

    If txtAnnc6 <> "" And Check1(2).Value = 1 Then
        Check1(2).Value = 0
    End If
    
    If txtAnnc6 <> "" Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
    End If
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, iBackAnnc, iAnnc4, iAnnc5, iAnnc6, iAnnc7, iAnnc8, iAnnc9, iAnnc10, iAnnc11
    Close #501
End Sub

Private Sub txtAnnc7_Change()

    If ck7Played.Value = 0 Then
        If IsNumeric(txtAnnc7) Or txtAnnc7 = "" Then
            iAnnc7 = txtAnnc7
            iiAnnc7 = txtAnnc7
            pAnnounce
        Else
            MsgBox "You have entered a character other than a number." _
            & vbCrLf & vbCrLf & "Enter the number of seconds you plan to use announcing and/or back-announcing this item." _
            & vbCrLf & vbCrLf & "If you do not intend to announce/back-announce this selection, enter a zero.", vbOKOnly, "Entry Error"
            txtAnnc7 = ""
            txtAnnc7.SetFocus
            Exit Sub
        End If
    End If
    
    '-------
    If Val(txtAnnc7) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce the current selection." _
        & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & _
        "If no announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Greater than 300 Seconds"
        txtAnnc7 = ""
        txtAnnc7.SetFocus
        Exit Sub
    End If
    '-------
    
    If frmPlanner!txtComposer4 <> "" Or frmPlanner!txtMinute4 <> "" Then
        If ck7Played.Value = 0 Then
            frmPlanner!txtAnnc4 = txtAnnc7
        Else
            frmPlanner!txtAnnc4 = ""
        End If
    End If
    
    If Val(txtAnnc7) > 15 Then
        txtAnnc7.ToolTipText = "Overwrite to change announce time or double-click to reduce announce time by 15 sec"
    Else
        txtAnnc7.ToolTipText = "Overwrite to change announce time"
    End If
End Sub

Private Sub txtAnnc7_DblClick()

   If Val(txtAnnc7) > 10 Then
        txtAnnc7 = Format((Val(txtAnnc7) - 15), "##")
    End If
     
    txtAnnc7.SelStart = 0 'begin selection at start
    txtAnnc7.SelLength = Len(txtAnnc7)
    txtAnnc7.SetFocus
End Sub

Private Sub txtAnnc7_GotFocus()
    txtAnnc7.SelStart = 0 'begin selection at start
    txtAnnc7.SelLength = Len(txtAnnc7)
End Sub

Private Sub txtAnnc7_LostFocus()

    If txtAnnc7 <> "" And Check1(3).Value = 1 Then
        Check1(3).Value = 0
    End If
    
    If txtAnnc7 <> "" Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
    End If
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, iBackAnnc, iAnnc4, iAnnc5, iAnnc6, iAnnc7, iAnnc8, iAnnc9, iAnnc10, iAnnc11
    Close #501
End Sub

Private Sub txtAnnc8_Change()

    If ck8Played.Value = 0 Then
       If IsNumeric(txtAnnc8) Or txtAnnc8 = "" Then
            iAnnc8 = txtAnnc8
            iiAnnc8 = txtAnnc8
            pAnnounce
        Else
            MsgBox "You have entered a character other than a number." _
            & vbCrLf & vbCrLf & "Enter the number of seconds you plan to use announcing and/or back-announcing this item." _
            & vbCrLf & vbCrLf & "If you do not intend to announce/back-announce this selection, enter a zero.", vbOKOnly, "Entry Error"
            txtAnnc8 = ""
            txtAnnc8.SetFocus
            Exit Sub
        End If
    End If
    
    '-------
    If Val(txtAnnc8) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce the current selection." _
        & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & _
        "If no announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Greater than 300 Seconds"
        txtAnnc8 = ""
        txtAnnc8.SetFocus
        Exit Sub
    End If
    '-------
    If frmPlanner!txtComposer5 <> "" Or frmPlanner!txtMinute5 <> "" Then
        If ck8Played.Value = 0 Then
            frmPlanner!txtAnnc5 = txtAnnc8
        Else
            frmPlanner!txtAnnc5 = ""
        End If
    End If
    
    If Val(txtAnnc8) > 15 Then
        txtAnnc8.ToolTipText = "Overwrite to change announce time or double-click to reduce announce time by 15 sec"
    Else
        txtAnnc8.ToolTipText = "Overwrite to change announce time"
    End If
End Sub

Private Sub txtAnnc8_DblClick()

   If Val(txtAnnc8) > 10 Then
        txtAnnc8 = Format((Val(txtAnnc8) - 15), "##")
    End If
     
    txtAnnc8.SelStart = 0 'begin selection at start
    txtAnnc8.SelLength = Len(txtAnnc8)
    txtAnnc8.SetFocus
End Sub

Private Sub txtAnnc8_GotFocus()
    txtAnnc8.SelStart = 0 'begin selection at start
    txtAnnc8.SelLength = Len(txtAnnc8)
End Sub

Private Sub txtAnnc8_LostFocus()

    If txtAnnc8 <> "" And Check1(4).Value = 1 Then
        Check1(4).Value = 0
    End If
    
    If txtAnnc8 <> "" Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
    End If
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, iBackAnnc, iAnnc4, iAnnc5, iAnnc6, iAnnc7, iAnnc8, iAnnc9, iAnnc10, iAnnc11
    Close #501
End Sub

Private Sub txtAnnc9_Change()

    If ck9Played.Value = 0 Then
       If IsNumeric(txtAnnc9) Or txtAnnc9 = "" Then
            iAnnc9 = txtAnnc9
            iiAnnc9 = txtAnnc9
            pAnnounce
        Else
            MsgBox "You have entered a character other than a number." _
            & vbCrLf & vbCrLf & "Enter the number of seconds you plan to use announcing and/or back-announcing this item." _
            & vbCrLf & vbCrLf & "If you do not intend to announce/back-announce this selection, enter a zero.", vbOKOnly, "Entry Error"
            txtAnnc9 = ""
            txtAnnc9.SetFocus
            Exit Sub
        End If
    End If
    
    '-------
    If Val(txtAnnc9) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce the current selection." _
        & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & _
        "If no announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Greater than 300 Seconds"
        txtAnnc9 = ""
        txtAnnc9.SetFocus
        Exit Sub
    End If
    '-------
    If frmPlanner!txtComposer6 <> "" Or frmPlanner!txtMinute6 <> "" Then
        If ck9Played.Value = 0 Then
            frmPlanner!txtAnnc6 = txtAnnc9
        Else
            frmPlanner!txtAnnc6 = ""
        End If
    End If
    
    If Val(txtAnnc9) > 15 Then
        txtAnnc9.ToolTipText = "Overwrite to change announce time or double-click to reduce announce time by 15 sec"
    Else
        txtAnnc9.ToolTipText = "Overwrite to change announce time"
    End If
End Sub

Private Sub txtAnnc9_DblClick()

   If Val(txtAnnc9) > 10 Then
        txtAnnc9 = Format((Val(txtAnnc9) - 15), "##")
    End If
     
    txtAnnc9.SelStart = 0 'begin selection at start
    txtAnnc9.SelLength = Len(txtAnnc9)
    txtAnnc9.SetFocus
End Sub

Private Sub txtAnnc9_GotFocus()
    txtAnnc9.SelStart = 0 'begin selection at start
    txtAnnc9.SelLength = Len(txtAnnc9)
End Sub

Private Sub txtAnnc9_LostFocus()

    If txtAnnc9 <> "" And Check1(5).Value = 1 Then
        Check1(5).Value = 0
    End If
    
    If txtAnnc9 <> "" Then
        cmdClearAnncTimes.Caption = "Clear Announce Times"
    End If
    
    Close #501
    Open "AnncTime.dat" For Output As #501
    Write #501, iBackAnnc, iAnnc4, iAnnc5, iAnnc6, iAnnc7, iAnnc8, iAnnc9, iAnnc10, iAnnc11
    Close #501
End Sub

Private Sub txtBackAnnc_Change()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
    
    Check1(0).Value = False

    If Not IsNumeric(txtBackAnnc) And txtBackAnnc <> "" Then
         MsgBox "You have entered the non-numeric character(s):  " & txtBackAnnc & vbCrLf & vbCrLf & _
         "Enter the number of seconds needed to back-announce the currently playing CD." & vbCrLf & vbCrLf & _
         "Entry can be up to 300 seconds, which is 5 minutes.", vbOKOnly, "CD Back Announce Entry Error"
         txtBackAnnc = "0"
         
         Exit Sub
    End If
'--------------
    miBackAnnc = 1
       
    If chkBackAnnc.Value = 0 Then
        If IsNumeric(txtBackAnnc) And Val(txtBackAnnc) < 301 Or txtBackAnnc = "" Then
        
            If Val(txtIntro) > 29 Then
                If Val(txtBackAnnc) = (Val((txtIntro) / 2) - 10) Then
                Label21.ToolTipText = ""
            End If
            
          End If
          
        Else
            MsgBox "Enter in seconds the estimated time needed to back-announce the current selection." _
            & vbCrLf & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes)." & vbCrLf & vbCrLf & _
            "If no back-announcement is planned, delete the entry or enter 0.", vbOKOnly, "Entry Error:  Non-Numeric or Greater than 300 Seconds"
            txtBackAnnc = ""
            txtBackAnnc.SetFocus
            Exit Sub
        End If
        
    ElseIf chkBackAnnc.Value = 1 Then
       
        txtBackAnnc.Text = ""
        
        If Val(txtIntro) > 29 Then
            chkBackAnnc.Caption = "Uncheck to set back-announce time to " & Format((Val((txtIntro) / 2) - 10), "##") & " sec "
        Else
             chkBackAnnc.Caption = "Back announce time is 0"
        End If
        
        Label21.ToolTipText = ""
    End If
'----------------


If (txtMinute4 <> "" Or txtSecond4 <> "") And (Check1(0).Value = 0) Then

    If (Val(txtBackAnnc) > 0) And Val(txtBackAnnc) < Val(txtIntro) Then
        txtAnnc4 = Val(txtIntro) - Val(txtBackAnnc)
    
    ElseIf Val(txtBackAnnc) >= Val(txtIntro) Then
        txtAnnc4 = "0"
    End If
    
End If

'---------------
    
    If txtBackAnnc <> "" And chkBackAnnc.Enabled = False Then
        chkBackAnnc.Enabled = True
    End If
        
    pAnnounce
End Sub

Private Sub txtBackAnnc_DblClick()

    If Val(txtBackAnnc) > 10 Then
        txtBackAnnc = Format((Val(txtBackAnnc) - 5), "##")
    Else
        txtBackAnnc = 0
        chkBackAnnc.Value = 1
    End If
    txtBackAnnc.SelStart = 0 'begin selection at start
    txtBackAnnc.SelLength = Len(txtBackAnnc)
    
    If chkBackAnnc.Enabled = False Then
        chkBackAnnc.Enabled = True
    End If
    
End Sub

Private Sub txtBackAnnc_GotFocus()
    txtBackAnnc.SelStart = 0 'begin selection at start
    txtBackAnnc.SelLength = Len(txtBackAnnc)
End Sub

Private Sub txtBackAnnc_LostFocus()
    If txtBackAnnc = "" Then
        txtBackAnnc = "0"
        txtAnnc4 = Val(txtIntro)
    End If
    pSetFocus
End Sub

Private Sub txtBlock_Change()
    'calls ADD feature after text block entry changes
    Dim iBlock As Integer
    If IsNumeric(txtBlock) Or txtBlock = "" Then
        iBlock = 1
    Else
        MsgBox "You have entered the non-numeric character:  " & txtBlock & vbCrLf & vbCrLf & _
        "Enter the total number of minutes you are programming, normally 60", 0, "Non-Numerical Entry"
        txtBlock.SelStart = 0 'begin selection at start
        txtBlock.SelLength = Len(txtBlock)
        txtBlock.SetFocus
        Exit Sub
    End If
    
    If Val(txtBlock) > 120.01 Then
        MsgBox "Planned Program Time entry cannot exceed 120 minutes (two hours)", vbOKOnly, "Entry Exceeds 120 Minute Limit"
        txtBlock = "60.0"
       Exit Sub
    End If
    
    If txtBlock.Text = "60.0" Then
        txtBlock.ForeColor = &H80000008  'black
    Else
        txtBlock.ForeColor = &H80& ' rust '&HC00000    'blue
    End If
        
    If txtBlock <> "" Then
        cmdClearBlock.Caption = "Cl&ear"
    ElseIf txtBlock = "" Then
        cmdClearBlock.Caption = ""
    End If
    lblRemain60 = txtBlock & " minutes planned"
    Call pAdd
 End Sub

Private Sub txtBlock_DblCliCk()
    txtBlock.SelStart = 0 'begin selection at start
    txtBlock.SelLength = Len(txtBlock)
End Sub

Private Sub txtBlock_LostFocus()

    If txtBlock = "" Or Val(txtBlock) < 1 Then
        MsgBox "Enter in minutes the total program time (music plus talk plus ID) planned for the hour. This normally is 60 minutes." _
         & vbCrLf & vbCrLf & _
        "If you are programming orher than 60 minutes (maximum 120 minutes), enter the total minutes planned.", vbOKOnly, "Planned Program Time Missing or Incorrect"
        txtBlock = "60.0"
        txtBlock.SelStart = 0 'begin selection at start
        txtBlock.SelLength = Len(txtBlock)
        Exit Sub
    End If
    
    If Not IsNumeric(txtBlock) Then
        txtBlock = "60.0"
    End If
    
    If txtBlock <> "" Then
        txtBlock.Text = Format$(txtBlock, "#0.0")
    End If

 End Sub

Private Sub txtCD10_Change()
    If mnuToolsExportLineCopy.Checked = True Then
        lbl7.BorderStyle = 0
    End If
End Sub

Private Sub txtCD10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtCD11_Change()
    If mnuToolsExportLineCopy.Checked = True Then
        lbl8.BorderStyle = 0
    End If
End Sub

Private Sub txtCD11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtCD4_Change()
    If mnuToolsExportLineCopy.Checked = True Then
        lbl1.BorderStyle = 0
    End If
End Sub

Private Sub txtCD4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtCD5_Change()
    If mnuToolsExportLineCopy.Checked = True Then
        lbl2.BorderStyle = 0
    End If
End Sub

Private Sub txtCD5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtCD6_Change()
    If mnuToolsExportLineCopy.Checked = True Then
        lbl3.BorderStyle = 0
    End If
End Sub

Private Sub txtCD6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtCD7_Change()
    If mnuToolsExportLineCopy.Checked = True Then
        lbl4.BorderStyle = 0
    End If
End Sub

Private Sub txtCD7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtCD8_Change()
    If mnuToolsExportLineCopy.Checked = True Then
        lbl5.BorderStyle = 0
    End If
End Sub

Private Sub txtCD8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtCD9_Change()
    If mnuToolsExportLineCopy.Checked = True Then
        lbl6.BorderStyle = 0
    End If
End Sub

Private Sub txtCD9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtCloseOut_Change()
    If Not IsNumeric(txtCloseOut) And txtCloseOut <> "" Then
         MsgBox "You have entered a non-numeric value", vbOKOnly, "Entry Error"
         txtCloseOut = "30"
         txtCloseOut.SetFocus
         Exit Sub
    End If

    miCloseOut = txtCloseOut
    
    Dim iCloseout As String
    iCloseout = Val(miCloseOut) / 60
    
    If Val(miCloseOut) < 60 Then
      ' 'Frame8.Caption = Format(miCloseOut, " ##") & " sec allocated for Closeout && ID"
       'Frame8.Caption = "Seconds Allocated for Closeout && ID"
       
    Else
       iCloseout = Val(miCloseOut) / 60
      ' 'Frame8.Caption = Format(iCloseout, " 0.#") & " min allocated for Closeout && ID"
    End If
   
    Label20.Caption = " • " & miCloseOut & " secs will remain for show closeout && station ID"
    
    Label15.Caption = "Time allocated for CLOSEOUT and station ID is " & txtCloseOut & " sec. To change, overwrite the entry in the 'time allocated for closeout && ID' box below. To save the change, click the 'Save Your Changes' button."
    
    chkAnnounce.Caption = "Check if you do Not want to include estimated announce times of " & txtIntro & _
    " sec for each selection and " & txtCloseOut & " sec for program closeout"
    
    If Val(txtCloseOut) >= 20 Then
        txtCloseOut.ToolTipText = "Double-click to reduce Closeout by 5 seconds"
    Else
        txtCloseOut.ToolTipText = ""
    End If
    
    pAnnounce
End Sub

Private Sub txtCloseOut_DblClick()
    If Val(txtCloseOut) > 10 Then
        txtCloseOut = Format((Val(txtCloseOut) - 5), "##")
        txtCloseOut.SelStart = 0 'begin selection at start
        txtCloseOut.SelLength = Len(txtCloseOut)
        txtCloseOut.SetFocus
    ElseIf Val(txtCloseOut) <= 10 Then
        txtCloseOut = frmDefaults!txtClose
    End If
End Sub

Private Sub txtCloseOut_GotFocus()
    txtCloseOut.SelStart = 0 'begin selection at start
    txtCloseOut.SelLength = Len(txtCloseOut)
End Sub

Private Sub txtCloseOut_LostFocus()
    If txtCloseOut = "" Then
        txtCloseOut = "30"
    End If
End Sub

Private Sub txtComposer10_Change()

    If txtComposer10 <> "" Then
        ck10Played.Enabled = True
        mnuExport.Enabled = True
    Else
        ck10Played.Enabled = False
    End If
    
   If mnuToolsExportLineCopy.Checked = True Then
        lbl7.BorderStyle = 0
    End If
    
    If txtMinute4 <> "" And Len(txtComposer4) < 3 Or txtMinute5 <> "" And Len(txtComposer5) < 3 Or txtMinute6 <> "" And Len(txtComposer6) < 3 _
    Or txtMinute7 <> "" And Len(txtComposer7) < 3 Or txtMinute8 <> "" And Len(txtComposer8) < 3 Or txtMinute9 <> "" And Len(txtComposer9) < 3 _
    Or txtMinute10 <> "" And Len(txtComposer10) < 3 Or txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
    Else
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdRestoreEntries.ToolTipText = ""
    End If
    
    If txtComposer10 = "" Then
        txtMinute10 = ""
        txtSecond10 = ""
        txtCD10 = ""
        txtAnnc10 = ""
        iAnnc10 = ""
        
        lbl7.BorderStyle = 0
        cmdRestoreEntries.Enabled = True
        pAnnounce
    End If
End Sub

Private Sub txtComposer10_DblClick()

    If txtComposer10 <> "" And txtMinute10 <> "" Then 'stores data
        mtxtComposer = txtComposer10
        mtxtMin = txtMinute10
        mtxtSec = txtSecond10
     End If
        
    If txtComposer10 <> "" Then
        txtMinute10 = ""
        txtSecond10 = ""
        txtComposer10 = ""
        txtAnnc10 = ""
        Check1(6).Value = 0
        txtComposer10.BackColor = &H80000005  'white
        
    ElseIf txtComposer10 = "" Then ' fills in from stored data
        txtComposer10 = mtxtComposer
        txtMinute10 = mtxtMin
        txtSecond10 = mtxtSec
    End If
    
End Sub

Private Sub txtComposer10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtComposer10_LostFocus()

    If txtComposer10 = "?" Then
        txtComposer10 = ""
    End If

    txtComposer10.SelStart = 0
    
    If Len(txtComposer10) > 0 And Len(txtComposer10) <= 2 Then
        txtComposer10.BackColor = &H80000018  'tool tip yellow
    Else
        txtComposer10.BackColor = &H80000005  'white
    End If
End Sub

Private Sub txtComposer11_Change()
 
    If txtComposer11 <> "" Then
        ck11Played.Enabled = True
        mnuExport.Enabled = True
    Else
        ck11Played.Enabled = False
    End If

    If mnuToolsExportLineCopy.Checked = True Then
        lbl8.BorderStyle = 0
    End If
    
    If txtMinute4 <> "" And Len(txtComposer4) < 3 Or txtMinute5 <> "" And Len(txtComposer5) < 3 Or txtMinute6 <> "" And Len(txtComposer6) < 3 _
    Or txtMinute7 <> "" And Len(txtComposer7) < 3 Or txtMinute8 <> "" And Len(txtComposer8) < 3 Or txtMinute9 <> "" And Len(txtComposer9) < 3 _
    Or txtMinute10 <> "" And Len(txtComposer10) < 3 Or txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
    Else
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdRestoreEntries.ToolTipText = ""
    End If

    If txtComposer11 = "" Then
        txtMinute11 = ""
        txtSecond11 = ""
        txtCD11 = ""
        txtAnnc11 = ""
        iAnnc11 = ""
      
        lbl8.BorderStyle = 0
        cmdRestoreEntries.Enabled = True
        pAnnounce
    End If
End Sub

Private Sub txtComposer11_DblClick()

    If txtComposer11 <> "" And txtMinute11 <> "" Then 'stores data
        mtxtComposer = txtComposer11
        mtxtMin = txtMinute11
        mtxtSec = txtSecond11
     End If
        
    If txtComposer11 <> "" Then
        txtMinute11 = ""
        txtSecond11 = ""
        txtComposer11 = ""
        txtAnnc11 = ""
        Check1(7).Value = 0
        txtComposer11.BackColor = &H80000005  'white
        
    ElseIf txtComposer11 = "" Then ' fills in from stored data
        txtComposer11 = mtxtComposer
        txtMinute11 = mtxtMin
        txtSecond11 = mtxtSec
    End If
    
End Sub

Private Sub txtComposer11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtComposer11_LostFocus()

    If txtComposer11 = "?" Then
        txtComposer11 = ""
    End If

    txtComposer11.SelStart = 0
    
    If Len(txtComposer11) > 0 And Len(txtComposer11) <= 2 Then
        txtComposer11.BackColor = &H80000018  'tool tip yellow
    Else
        txtComposer11.BackColor = &H80000005  'white
    End If
End Sub

Private Sub txtComposer4_Change()

    If chkAnnounce.Value = 0 And (txtIntroSetting.Text = "" Or txtCloseOut.Text = "" Or txtSpotLength.Text = "") Then
        MsgBox "Estimated Announce Time have not been entered. From 'Announce-Times' Menu select" _
        & vbCrLf & "'Set Estimated Announce Times' and enter the requested times or select 'Use Defaults'.", vbOKOnly, "Need Estimated Announce Times"
        fraIntro.Visible = True
        Exit Sub
    End If

    If txtComposer4 <> "" Then
        ck4Played.Enabled = True
        mnuExport.Enabled = True
    Else
        ck4Played.Enabled = False
    End If

    If mnuToolsExportLineCopy.Checked = True Then
        lbl1.BorderStyle = 0
    End If
    
    If txtMinute4 <> "" And Len(txtComposer4) < 3 Or txtMinute5 <> "" And Len(txtComposer5) < 3 Or txtMinute6 <> "" And Len(txtComposer6) < 3 _
    Or txtMinute7 <> "" And Len(txtComposer7) < 3 Or txtMinute8 <> "" And Len(txtComposer8) < 3 Or txtMinute9 <> "" And Len(txtComposer9) < 3 _
    Or txtMinute10 <> "" And Len(txtComposer10) < 3 Or txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
    Else
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdRestoreEntries.ToolTipText = ""
    End If
    
    If txtComposer4 = "" Then
        txtMinute4 = ""
        txtSecond4 = ""
        txtCD4 = ""
        txtAnnc4 = ""
        iAnnc4 = ""

        lbl1.BorderStyle = 0
        pAnnounce
        cmdRestoreEntries.Enabled = True
    End If
    
End Sub

Private Sub txtComposer4_DblCliCk()
   
    If txtComposer4 <> "" And txtMinute4 <> "" Then 'stores data
        mtxtComposer = txtComposer4
        mtxtMin = txtMinute4
        mtxtSec = txtSecond4
     End If
        
    If txtComposer4 <> "" Then
        txtMinute4 = ""
        txtSecond4 = ""
        txtComposer4 = ""
        txtAnnc4 = ""
        Check1(0).Value = 0
        txtComposer4.BackColor = &H80000005  'white
        
    ElseIf txtComposer4 = "" Then ' fills in from stored data
        txtComposer4 = mtxtComposer
        txtMinute4 = mtxtMin
        txtSecond4 = mtxtSec
    End If
    
End Sub

Private Sub txtComposer4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtComposer4_LostFocus()

    If txtComposer4 = "?" Then
        txtComposer4 = ""
    End If
    
    txtComposer4.SelStart = 0

    If Len(txtComposer4) > 0 And Len(txtComposer4) <= 2 Then
        txtComposer4.BackColor = &H80000018  'tool tip yellow
    Else
        txtComposer4.BackColor = &H80000005  'white
    End If
 
End Sub

Private Sub txtComposer5_Change()

    If txtComposer5 <> "" Then
        ck5Played.Enabled = True
        mnuExport.Enabled = True
    Else
        ck5Played.Enabled = False
    End If

    If mnuToolsExportLineCopy.Checked = True Then
        lbl2.BorderStyle = 0
    End If

    If txtMinute4 <> "" And Len(txtComposer4) < 3 Or txtMinute5 <> "" And Len(txtComposer5) < 3 Or txtMinute6 <> "" And Len(txtComposer6) < 3 _
    Or txtMinute7 <> "" And Len(txtComposer7) < 3 Or txtMinute8 <> "" And Len(txtComposer8) < 3 Or txtMinute9 <> "" And Len(txtComposer9) < 3 _
    Or txtMinute10 <> "" And Len(txtComposer10) < 3 Or txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
    Else
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdRestoreEntries.ToolTipText = ""
    End If
        
    If txtComposer5 = "" Then
        txtMinute5 = ""
        txtSecond5 = ""
        txtCD5 = ""
        txtAnnc5 = ""
        iAnnc5 = ""
       
        lbl2.BorderStyle = 0
        cmdRestoreEntries.Enabled = True
        pAnnounce
    End If
    
End Sub

Private Sub txtComposer5_DblCliCk()

     If txtComposer5 <> "" And txtMinute5 <> "" Then 'stores data
        mtxtComposer = txtComposer5
        mtxtMin = txtMinute5
        mtxtSec = txtSecond5
     End If
        
    If txtComposer5 <> "" Then
        txtMinute5 = ""
        txtSecond5 = ""
        txtComposer5 = ""
        txtAnnc5 = ""
        Check1(1).Value = 0
        txtComposer5.BackColor = &H80000005  'white
        
    ElseIf txtComposer5 = "" Then ' fills in from stored data
        txtComposer5 = mtxtComposer
        txtMinute5 = mtxtMin
        txtSecond5 = mtxtSec
    End If
End Sub

Private Sub txtComposer5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtComposer5_LostFocus()

    If txtComposer5 = "?" Then
        txtComposer5 = ""
    End If

    txtComposer5.SelStart = 0
    
    If Len(txtComposer5) > 0 And Len(txtComposer5) <= 2 Then
        txtComposer5.BackColor = &H80000018  'tool tip yellow
    Else
        txtComposer5.BackColor = &H80000005  'white
    End If
End Sub

Private Sub txtComposer6_Change()

    If txtComposer6 <> "" Then
        ck6Played.Enabled = True
        mnuExport.Enabled = True
    Else
        ck6Played.Enabled = False
    End If

    If mnuToolsExportLineCopy.Checked = True Then
        lbl3.BorderStyle = 0
    End If
    
    If txtMinute4 <> "" And Len(txtComposer4) < 3 Or txtMinute5 <> "" And Len(txtComposer5) < 3 Or txtMinute6 <> "" And Len(txtComposer6) < 3 _
    Or txtMinute7 <> "" And Len(txtComposer7) < 3 Or txtMinute8 <> "" And Len(txtComposer8) < 3 Or txtMinute9 <> "" And Len(txtComposer9) < 3 _
    Or txtMinute10 <> "" And Len(txtComposer10) < 3 Or txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
    Else
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdRestoreEntries.ToolTipText = ""
    End If

    If txtComposer6 = "" Then
        txtMinute6 = ""
        txtSecond6 = ""
        txtCD6 = ""
        txtAnnc6 = ""
        iAnnc6 = ""
        
        lbl3.BorderStyle = 0
        cmdRestoreEntries.Enabled = True
        pAnnounce
    End If
End Sub

Private Sub txtComposer6_DblCliCk()

    If txtComposer6 <> "" And txtMinute6 <> "" Then 'stores data
        mtxtComposer = txtComposer6
        mtxtMin = txtMinute6
        mtxtSec = txtSecond6
     End If
        
    If txtComposer6 <> "" Then
        txtMinute6 = ""
        txtSecond6 = ""
        txtComposer6 = ""
        txtAnnc6 = ""
        Check1(2).Value = 0
        txtComposer6.BackColor = &H80000005  'white
        
    ElseIf txtComposer6 = "" Then ' fills in from stored data
        txtComposer6 = mtxtComposer
        txtMinute6 = mtxtMin
        txtSecond6 = mtxtSec
    End If
    
End Sub

Private Sub txtComposer6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
       KeyAscii = 39
    End If
End Sub

Private Sub txtComposer6_LostFocus()

    If txtComposer6 = "?" Then
        txtComposer6 = ""
    End If

    txtComposer6.SelStart = 0
    
    If Len(txtComposer6) > 0 And Len(txtComposer6) <= 2 Then
        txtComposer6.BackColor = &H80000018  'tool tip yellow
    Else
        txtComposer6.BackColor = &H80000005  'white
    End If
End Sub

Private Sub txtComposer7_Change()

    If txtComposer7 <> "" Then
        ck7Played.Enabled = True
        mnuExport.Enabled = True
    Else
        ck7Played.Enabled = False
    End If

    If mnuToolsExportLineCopy.Checked = True Then
        lbl4.BorderStyle = 0
    End If
    
    If txtMinute4 <> "" And Len(txtComposer4) < 3 Or txtMinute5 <> "" And Len(txtComposer5) < 3 Or txtMinute6 <> "" And Len(txtComposer6) < 3 _
    Or txtMinute7 <> "" And Len(txtComposer7) < 3 Or txtMinute8 <> "" And Len(txtComposer8) < 3 Or txtMinute9 <> "" And Len(txtComposer9) < 3 _
    Or txtMinute10 <> "" And Len(txtComposer10) < 3 Or txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
    Else
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdRestoreEntries.ToolTipText = ""
    End If

    If txtComposer7 = "" Then
        txtMinute7 = ""
        txtSecond7 = ""
        txtCD7 = ""
        txtAnnc7 = ""
        iAnnc7 = ""
        
        lbl4.BorderStyle = 0
        cmdRestoreEntries.Enabled = True
        pAnnounce
    End If
End Sub

Private Sub txtComposer7_DblCliCk()

    If txtComposer7 <> "" And txtMinute7 <> "" Then 'stores data
        mtxtComposer = txtComposer7
        mtxtMin = txtMinute7
        mtxtSec = txtSecond7
     End If
        
    If txtComposer7 <> "" Then
        txtMinute7 = ""
        txtSecond7 = ""
        txtComposer7 = ""
        txtAnnc7 = ""
        Check1(3).Value = 0
        txtComposer7.BackColor = &H80000005  'white
        
    ElseIf txtComposer7 = "" Then ' fills in from stored data
        txtComposer7 = mtxtComposer
        txtMinute7 = mtxtMin
        txtSecond7 = mtxtSec
    End If
    
End Sub


Private Sub txtComposer7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtComposer7_LostFocus()

    If txtComposer7 = "?" Then
        txtComposer7 = ""
    End If

    txtComposer7.SelStart = 0
    
    If Len(txtComposer7) > 0 And Len(txtComposer7) <= 2 Then
        txtComposer7.BackColor = &H80000018  'tool tip yellow
    Else
        txtComposer7.BackColor = &H80000005  'white
    End If
End Sub

Private Sub txtComposer8_Change()

    If txtComposer8 <> "" Then
        ck8Played.Enabled = True
        mnuExport.Enabled = True
    Else
        ck8Played.Enabled = False
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl5.BorderStyle = 0
    End If
    
    If txtMinute4 <> "" And Len(txtComposer4) < 3 Or txtMinute5 <> "" And Len(txtComposer5) < 3 Or txtMinute6 <> "" And Len(txtComposer6) < 3 _
    Or txtMinute7 <> "" And Len(txtComposer7) < 3 Or txtMinute8 <> "" And Len(txtComposer8) < 3 Or txtMinute9 <> "" And Len(txtComposer9) < 3 _
    Or txtMinute10 <> "" And Len(txtComposer10) < 3 Or txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
    Else
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdRestoreEntries.ToolTipText = ""
    End If
    
    If txtComposer8 = "" Then
        txtMinute8 = ""
        txtSecond8 = ""
        txtCD8 = ""
        txtAnnc8 = ""
        iAnnc8 = ""
       
        lbl5.BorderStyle = 0
        cmdRestoreEntries.Enabled = True
        pAnnounce
    End If
End Sub

Private Sub txtComposer8_DblCliCk()

    If txtComposer8 <> "" And txtMinute8 <> "" Then 'stores data
        mtxtComposer = txtComposer8
        mtxtMin = txtMinute8
        mtxtSec = txtSecond8
     End If
        
    If txtComposer8 <> "" Then
        txtMinute8 = ""
        txtSecond8 = ""
        txtComposer8 = ""
        txtAnnc8 = ""
        Check1(4).Value = 0
        txtComposer8.BackColor = &H80000005  'white
        
    ElseIf txtComposer8 = "" Then ' fills in from stored data
        txtComposer8 = mtxtComposer
        txtMinute8 = mtxtMin
        txtSecond8 = mtxtSec
    End If
End Sub

Private Sub txtComposer8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtComposer8_LostFocus()

    If txtComposer8 = "?" Then
        txtComposer8 = ""
    End If

    txtComposer8.SelStart = 0
    
    If Len(txtComposer8) > 0 And Len(txtComposer8) <= 2 Then
        txtComposer8.BackColor = &H80000018  'tool tip yellow
    Else
        txtComposer8.BackColor = &H80000005  'white
    End If
End Sub

Private Sub txtComposer9_Change()

    If txtComposer9 <> "" Then
        ck9Played.Enabled = True
        mnuExport.Enabled = True
    Else
        ck9Played.Enabled = False
    End If

    If mnuToolsExportLineCopy.Checked = True Then
        lbl6.BorderStyle = 0
    End If
    
    If txtMinute4 <> "" And Len(txtComposer4) < 3 Or txtMinute5 <> "" And Len(txtComposer5) < 3 Or txtMinute6 <> "" And Len(txtComposer6) < 3 _
    Or txtMinute7 <> "" And Len(txtComposer7) < 3 Or txtMinute8 <> "" And Len(txtComposer8) < 3 Or txtMinute9 <> "" And Len(txtComposer9) < 3 _
    Or txtMinute10 <> "" And Len(txtComposer10) < 3 Or txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
    Else
        cmdRestoreEntries.Caption = "&Restore Lineup"
        cmdRestoreEntries.BackColor = &H8000000F
        cmdRestoreEntries.ToolTipText = ""
    End If
    
    If txtComposer9 = "" Then
        txtMinute9 = ""
        txtSecond9 = ""
        txtCD9 = ""
        txtAnnc9 = ""
        iAnnc9 = ""
       
        lbl6.BorderStyle = 0
        cmdRestoreEntries.Enabled = True
        pAnnounce
    End If
End Sub

Private Sub txtComposer9_DblCliCk()

    If txtComposer9 <> "" And txtMinute9 <> "" Then 'stores data
        mtxtComposer = txtComposer9
        mtxtMin = txtMinute9
        mtxtSec = txtSecond9
     End If
        
    If txtComposer9 <> "" Then
        txtMinute9 = ""
        txtSecond9 = ""
        txtComposer9 = ""
        txtAnnc9 = ""
        Check1(5).Value = 0
        txtComposer9.BackColor = &H80000005  'white
        
    ElseIf txtComposer9 = "" Then ' fills in from stored data
        txtComposer9 = mtxtComposer
        txtMinute9 = mtxtMin
        txtSecond9 = mtxtSec
    End If
    
End Sub

Private Sub txtComposer9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtComposer9_LostFocus()

    If txtComposer9 = "?" Then
        txtComposer9 = ""
    End If

    txtComposer9.SelStart = 0
    
    If Len(txtComposer9) > 0 And Len(txtComposer9) <= 2 Then
        txtComposer9.BackColor = &H80000018  'tool tip yellow
    Else
        txtComposer9.BackColor = &H80000005  'white
    End If
End Sub

Private Sub txtIntro_DblClick()
On Error GoTo HandleErrors

    If Val(txtIntro) > 10 Then
        txtIntro = Format((Val(txtIntro) - 10), "##")
        txtIntro.SelStart = 0 'begin selection at start
        txtIntro.SelLength = Len(txtIntro)
        txtIntro.SetFocus
'-----------
    ElseIf Val(txtIntro) <= 10 Then
    
        Dim PlanTime, IntroOut, sClose, Spot As Integer
        Open "Times.dat" For Input As #23
        Input #23, PlanTime, IntroOut, sClose, Spot
        Close #23
        txtIntro = IntroOut
    Else
        txtIntro = "50"
    End If
    
    If chkBackAnnc.Value = 0 Then
    
        If Format(Val(txtBackAnnc), "##") = Format((Val((txtIntro) / 2) - 10), "##") Then
      
            Label21.ToolTipText = ""
        Else
                          
            If chkBackAnnc = 0 Then
                Label21.ToolTipText = " Double-click to reset back-announce time to " & Format((Val((txtIntro) / 2) - 10), "##") & " sec "
            Else
                Label21.ToolTipText = ""
            End If
        
        End If
    Else
    
        If Val(txtIntro) > 29 Then
            chkBackAnnc.Caption = "Uncheck to set back-announce time to " & Format((Val((txtIntro) / 2) - 10), "##") & " sec "
        Else
             chkBackAnnc.Caption = "Back announce time is 0"
        End If
                
    End If
    
    txtIntro.SelStart = 0 'begin selection at start
    txtIntro.SelLength = Len(txtIntro)
    txtIntro.SetFocus
    
    pAnnounce
HandleErrors:
End Sub

Private Sub txtIntroSetting_Change()
    If txtIntroSetting <> "" Then
        txtIntro = txtIntroSetting
    End If
End Sub

Private Sub txtIntroSetting_GotFocus()
    txtIntroSetting.SelStart = 0 'begin selection at start
    txtIntroSetting.SelLength = Len(txtIntroSetting)
End Sub

Private Sub txtIntroSetting_LostFocus()
    If Not IsNumeric(txtIntroSetting) Or txtIntroSetting = "" Then
        txtIntroSetting = "50"
    Else
        txtIntroSetting.Text = Format$(txtIntroSetting, "##")
    End If
End Sub

Private Sub txtIntro_Change()

    If Not IsNumeric(txtIntro) And txtIntro <> "" Then
         MsgBox "You have entered the non-numeric character:  " & txtIntro & vbCrLf & vbCrLf & _
         "Enter the average number of seconds allocated to each" & vbCrLf & "music selection for its introduction plus back-announce", vbOKOnly, "Entry Error"
         txtIntro = "50"
         txtIntro.SetFocus
         Exit Sub
    End If
    
    If Val(txtIntro) > 300 Then
        MsgBox "Enter in seconds the estimated time needed to Intro and Back-Announce each music selection." _
        & vbCrLf & vbCrLf & "Entry can range from 0 to a maximum of 300 seconds (which is 5 minutes).", _
        vbOKOnly, "Entry Greater than 300 Seconds"
        txtIntro = ""
        txtIntro.SetFocus
        Exit Sub
    End If
      
    chkAnnounce.Caption = "Check if you do Not want to include estimated announce times of " & txtIntro & _
    " sec for each selection and " & txtCloseOut & " sec for program closeout"
   
    frmPlanner!lblAnnc.Caption = "Intro/Back-Annc " & frmTimeRemain!txtIntro & " sec"
'------
    If Val(txtIntro) > 29 And chkBackAnnc.Visible = True Then
        txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
    Else
        txtBackAnnc = "0"
    End If
'------

    If txtAnnc4 <> "" Then
        txtAnnc4 = Val(txtIntro) - Val(txtBackAnnc)
    End If
    
    If txtMinute5 <> "" Then
        txtAnnc5.Text = txtIntro
    End If
    
    If txtMinute6 <> "" Then
        txtAnnc6.Text = txtIntro
    End If
    
    If txtMinute7 <> "" Then
        txtAnnc7.Text = txtIntro
    End If
    
    If txtMinute8 <> "" Then
        txtAnnc8.Text = txtIntro
    End If
    
    If txtMinute9 <> "" Then
        txtAnnc9.Text = txtIntro
    End If
    
    If txtMinute10 <> "" Then
        txtAnnc10.Text = txtIntro
    End If
    
    If txtMinute11 <> "" Then
        txtAnnc11.Text = txtIntro
    End If
    
    pAnnounce
  End Sub

Private Sub txtIntro_GotFocus()
    txtIntro.SelStart = 0 'begin selection at start
    txtIntro.SelLength = Len(txtIntro)
    lblAverageTime.ForeColor = &H80&
End Sub

Private Sub txtIntro_LostFocus()

    If txtIntro = "" Then
        txtIntro = "50"
        txtIntroSetting = "50"
    End If
    
    If chkBackAnnc.Value = 0 Then
    
        If Format(Val(txtBackAnnc), "##") = Format((Val((txtIntro) / 2) - 10), "##") Then
        
            Label21.ToolTipText = ""
        Else
                      
            If chkBackAnnc.Value = 0 Then
                Label21.ToolTipText = " Double-click to reset back-announce time to " & Format((Val((txtIntro) / 2) - 10), "##") & " sec "
            Else
                Label21.ToolTipText = ""
            End If
            
        End If
    Else
        
        If Val(txtIntro) > 29 Then
            chkBackAnnc.Caption = "Uncheck to set back-announce time to " & Format((Val((txtIntro) / 2) - 10), "##") & " sec "
        Else
             chkBackAnnc.Caption = "Back announce time is 0"
        End If
        
    End If
        
    If frmPlanner!mnuLinkTimeRemain.Checked = True Then
        If frmPlanner!txtAnnc1 <> txtIntro And frmPlanner!txtAnnc2 <> txtIntro Then
            frmPlanner!fraAnnc.ForeColor = &HC00000  'blue, indicates announce times do not agree
            frmPlanner!fraAnnc.ToolTipText = "Double-Click here to set Annc times to " & txtIntro & " sec average for Intro & Back-Announce"
        End If
    End If
    
   lblAverageTime.ForeColor = &H404040
   pSetFocus
End Sub

Private Sub txtMinute1_Change()
   
    Dim sSeconds As String 'to enter 0 seconds if txtSecond1 = ""
    If txtSecond1 <> "" Then
        sSeconds = txtSecond1
    Else
        sSeconds = "00"
    End If
 
    If txtMinute1 = "" And txtSecond1 = "00" Then
        txtSecond1 = ""
    End If
    
    'clock ending time displayed
    If Val(txtMinute1) > 59 Then
        MsgBox "Entry may not exceed 59 minutes.", vbOKOnly, "Minute Entry Error"
        txtMinute1 = ""
        txtMinute1.SetFocus
    Exit Sub
    End If
    
    If (txtMinute1 <> "" Or txtSecond1 <> "") And (txtMinute2 <> "" Or txtSecond2 > "00") Then
        Label12.Visible = True
        imgOnAirSign.Visible = True
        imgDisc.Visible = True
        Label13.Visible = True
    Else
        Label12.Visible = False
        imgOnAirSign.Visible = False
        imgDisc.Visible = False
        Label13.Visible = False
    End If
    
  '  lblHour.Visible = False 'manual entry clears setting time with system clock
  '  Shape1.Visible = True
    mCurrentTime = 0
    
    If Val(txtMinute1) < "55" Then
        lblCurrentTime.Alignment = 0 'left
        lblCurrentTime.ForeColor = vbBlack
        lblCurrentTime.Caption = " 1. Click the 'Set Current Time' button to enter the current time:"
    Else
        lblCurrentTime.ForeColor = &HC00000 'blue
        lblCurrentTime.Alignment = 1 'right
        
        If txtMinute1 >= "55" And txtMinute1 <= "57" Then
            lblCurrentTime.Caption = "Approaching the end of the hour"
            lblCurrentTime.ToolTipText = "Double-Click text line if timing for the following hour of programming begins in this hour hour at time " & txtMinute1 & " min " & sSeconds & " sec."
        ElseIf txtMinute1 = "58" Or txtMinute1 = "59" Then
            lblCurrentTime.Caption = "Double-Click here if the next hour's program begins at this time:"
        End If
        
    End If
    
    cmdSetTime.Visible = False
    Frame7.Caption = "Set Current Time"
    
    If txtMinute1 <> "" Then
        cmdClearTimes.Caption = "&Clear Currrent Time"
    Else
        cmdClearTimes.Caption = "Clear Times"
    End If
    
    If chkStopWatch.Value = 0 Then
    
        If txtMinAdj = "" And txtSecAdj = "" Then
            cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
        Else
             cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
        End If
        
    Else
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Stopwatch) F5"
    End If

    If frmPlanner!mnuLinkTimeRemain.Checked = True And (txtMinute1 <> "" Or txtSecond1 <> "") Then
        
        If lblCurrentTime.ForeColor = vbRed Then
            frmPlanner!staStatus.Panels(2) = "Timing begin in previous hour at time " & txtMinute1 & " min " & sSeconds & " sec"
        Else
            frmPlanner!staStatus.Panels(2) = "* Run time began at " & txtMinute1 & " min " & sSeconds & " sec past the hour"
        End If
    
    Else
        frmPlanner!staStatus.Panels(2) = ""
    End If
    txtSecond1.TabStop = True
    
    If txtSpotsS <> "" And txtSpotsS <> "0" And txtSecond3.Visible = True Then
        lblS.Visible = True
        lblSspot.Visible = True
        Label31.Visible = True
        If txtSpotsS = "1" Then
            lblS.Caption = txtSpotsS & " spot"
            Label31.Caption = txtSpotsS & " spot"
        Else
            lblS.Caption = txtSpotsS & " spots"
            Label31.Caption = txtSpotsS & " spots"
        End If

    Else
       lblS.Visible = False
       lblSspot.Visible = False
       Label31.Visible = False
       lblS.Caption = ""
       Label31.Caption = ""
    End If
  
    If lblLinked.Visible = True Then
        lblLinked_DblClick
    End If

    pAnnounce
End Sub

Private Sub txtMinute1_DblClick()
    txtMinute1 = ""
    txtSecond1 = ""
    lblHour = ""
    lblHour.Visible = False
    Shape1.Width = 840 'normal
    Shape1.Left = 5715
End Sub

Private Sub txtMinute1_GotFocus()
    txtMinute1.SelStart = 0 'begin selection at start
    txtMinute1.SelLength = Len(txtMinute1)
End Sub

Private Sub txtMinute1_LostFocus()
    If txtMinute1 = "" And txtSecond1 = "" Then
        Label12.Visible = False
        imgOnAirSign.Visible = False
        imgDisc.Visible = False
        Label13.Visible = False
    End If
    imgHand.Visible = False
End Sub

Private Sub txtMinute10_Change()

    If txtMinute10 <> "" Then
        lbl7.ForeColor = &H80& 'rust
    Else
        lbl7.ForeColor = &H808080 'gray
    End If

    If txtMinute10 <> "" Then
        txtMinute10.ToolTipText = " Double-Click to clear "
    Else
        txtMinute10.ToolTipText = "Enter CD playing time (minutes)"
    End If

    If txtMinute10 <> "" And Annc10 = "" And Check1(6).Value = 0 Then
        txtAnnc10 = txtIntro
    Else
        txtAnnc10 = Annc10
    End If
    
    If txtMinute10 = "" Then
        txtAnnc10 = ""
    End If

    If ck10Played.Value = 0 Then
       iMinute10 = txtMinute10
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl7.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtMinute10 <> frmPlanner!txtMinute7 Then
        txtMinute10.ForeColor = vbRed
        frmPlanner!txtMinute7.ForeColor = vbRed
    Else
        txtMinute10.ForeColor = &H80000012 'black
        frmPlanner!txtMinute7.ForeColor = &H80000012 'black
    End If
    
    If txtMinute10 <> "" And txtComposer10 <> "" And ck10Played.Enabled = False Then
        ck10Played.Enabled = True
    End If
    
    If txtMinute10 <> "" Or txtSecond10 <> "" Then
        Check1(6).Enabled = True
    Else
        Check1(6).Value = 0
        Check1(6).Enabled = False
    End If
    
    If txtMinute10 <> "" And txtComposer10 = "" Then
        txtComposer10 = "G"
    End If
    
    If ck10Played.Value = 0 And txtMinute10 = "" And txtSecond10 = "" Then
        txtComposer10 = ""
        txtComposer10.BackColor = &H80000005 ' white
    End If
    
    If txtMinute10 = "" And txtSecond10 = "" Then
        txtAnnc10.Text = ""
        txtAnnc10.Enabled = False
    Else
        txtAnnc10.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtMinute10_DblClick()
    txtMinute10 = ""
    txtSecond10 = ""
    
    If txtComposer10 <> "" Then
        ck10Played.Enabled = False
    End If
End Sub

Private Sub txtMinute10_GotFocus()
    Annc10 = ""
    txtMinute10.SelStart = 0 'begin selection at start
    txtMinute10.SelLength = Len(txtMinute10)
End Sub

Private Sub txtMinute10_LostFocus()
    If txtMinute10 = "" And txtSecond10 = "00" Then
        txtSecond10 = ""
    End If
    
    If txtMinute10 = "" Then
        Check1(6).Value = 0
    End If
    
    If txtMinute10 <> "" And Len(txtComposer10) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
        cmdRestoreEntries.ToolTipText = " A trial entry is any entry that contains less than 3 characters in the Music Lineup text box "
        
    ElseIf Len(txtComposer10) >= 3 And cmdRestoreEntries.Caption <> "&Remove Trial Entries" Then
      cmdRestoreEntries.Enabled = False
    End If
End Sub

Private Sub txtMinute11_Change()

    If txtMinute11 <> "" Then
        lbl8.ForeColor = &H80& 'rust
    Else
        lbl8.ForeColor = &H808080 'gray
    End If

    If txtMinute11 <> "" Then
        txtMinute11.ToolTipText = " Double-Click to clear "
    Else
        txtMinute11.ToolTipText = "Enter CD playing time (minutes)"
    End If

    If txtMinute11 <> "" And Annc11 = "" And Check1(7).Value = 0 Then
        txtAnnc11 = txtIntro
    Else
        txtAnnc11 = Annc11
    End If
    
    If txtMinute11 = "" Then
        txtAnnc11 = ""
    End If

    If ck11Played.Value = 0 Then
       iMinute11 = txtMinute11
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl7.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtMinute11 <> frmPlanner!txtMinute8 Then
        txtMinute11.ForeColor = vbRed
        frmPlanner!txtMinute8.ForeColor = vbRed
    Else
        txtMinute11.ForeColor = &H80000012 'black
        frmPlanner!txtMinute8.ForeColor = &H80000012 'black
    End If
    
    If txtMinute11 <> "" And txtComposer11 <> "" And ck11Played.Enabled = False Then
        ck10Played.Enabled = True
    End If
    
    If txtMinute11 <> "" Or txtSecond11 <> "" Then
        Check1(7).Enabled = True
    Else
        Check1(7).Value = 0
        Check1(7).Enabled = False
    End If
    
    If txtMinute11 <> "" And txtComposer11 = "" Then
        txtComposer11 = "H"
    End If
    
    If ck11Played.Value = 0 And txtMinute11 = "" And txtSecond11 = "" Then
        txtComposer11 = ""
        txtComposer11.BackColor = &H80000005 ' white
    End If
    
    If txtMinute11 = "" And txtSecond11 = "" Then
        txtAnnc11.Text = ""
        txtAnnc11.Enabled = False
    Else
        txtAnnc11.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtMinute11_DblClick()
    txtMinute11 = ""
    txtSecond11 = ""
    
    If txtComposer11 <> "" Then
        ck11Played.Enabled = False
    End If
End Sub

Private Sub txtMinute11_GotFocus()
    Annc11 = ""
    txtMinute11.SelStart = 0 'begin selection at start
    txtMinute11.SelLength = Len(txtMinute11)
End Sub

Private Sub txtMinute11_LostFocus()
    If txtMinute11 = "" And txtSecond11 = "00" Then
        txtSecond11 = ""
    End If
    
    If txtMinute11 = "" Then
        Check1(7).Value = 0
    End If

    If txtMinute11 <> "" And Len(txtComposer11) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
        cmdRestoreEntries.ToolTipText = " A trial entry is any entry that contains less than 3 letters in the Music Lineup text box "
        
    ElseIf Len(txtComposer11) >= 3 And cmdRestoreEntries.Caption <> "&Remove Trial Entries" Then
      cmdRestoreEntries.Enabled = False
    End If
End Sub

Private Sub txtMinute2_Change()

    Dim iMinute2 As Integer
    Dim iSecond2 As Integer

    If txtMinute2 <> "" Then
       iMinute2 = Val(txtMinute2)
    Else
       iMinute2 = 0
    End If

    If iMinute2 > 0 Or iSecond2 > 0 Then

        If txtMinute1 <> "" Or txtSecond1 <> "" Then
            Label12.Visible = True
            imgOnAirSign.Visible = True
            imgDisc.Visible = True
            Label13.Visible = True
        Else
            Label12.Visible = False
            imgOnAirSign.Visible = False
            imgDisc.Visible = False
            Label13.Visible = False
        End If

        lblEndTime.Visible = True

    Else
        txtBackAnnc.Visible = False
        chkBackAnnc.Visible = False

        If chkBackAnnc.Value = 1 Then
            chkBackAnnc.Value = 0
        End If
        txtBackAnnc = "0"
        Label21.Visible = False
        Label12.Visible = False
        imgOnAirSign.Visible = False
        imgDisc.Visible = False
        Label13.Visible = False
        lblEndTime.Visible = False
    End If
'----
  'changes caption for Music Lineup frame
    If txtMinute2 = "" Then
        Frame1.Caption = "Music Lineup"

        If txtMinute1 <> "" Then
            cmdClearTimes.Caption = "&Clear Curremt Time"
        Else
            cmdClearTimes.Caption = "Clear Times"
        End If
    ElseIf txtMinute2 <> "" Then
        Frame1.Caption = "Additional Music Lineup"
        cmdClearTimes.Caption = "&Clear CD Time"
    End If
    
    If txtSpotsS <> "" And txtSpotsS <> "0" Then
        
        If txtSecond3.Visible = False Then
            lblS.Visible = False
        Else
            lbl5.Visible = True
        End If
    
        lblSspot.Visible = True
        Label31.Visible = True
        If txtSpotsS = "1" Then
            lblS.Caption = txtSpotsS & " spot"
            Label31.Caption = txtSpotsS & " spot"
        Else
            lblS.Caption = txtSpotsS & " spots"
            Label31.Caption = txtSpotsS & " spots"
        End If
    Else
       lblS.Visible = False
       lblSspot.Visible = False
       Label31.Visible = False
       lblS.Caption = ""
       Label31.Caption = ""
    End If
    
    If lblLinked.Visible = True Then
        lblLinked_DblClick
    End If

   pAnnounce
End Sub

Private Sub txtMinute2_DblClick()
    txtMinute2 = ""
    txtSecond2 = ""
End Sub

Private Sub txtMinute2_GotFocus()
    txtMinute2.SelStart = 0 'begin selection at start
    txtMinute2.SelLength = Len(txtMinute2)
End Sub

Private Sub txtMinute2_LostFocus()
    If txtMinute2.Text = "" And txtSecond2 = "00" Then
        txtSecond2.Text = ""
    End If
    
    If txtMinute2.Text = "" And txtSecond2 = "" Then
        imgDisc.Visible = False
    End If
    
    Label6.BackColor = &H8000000F 'black
    Label6.BorderStyle = 0
End Sub

Private Sub txtMinute3_Change()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If
    
    If (Val(txtMinute3) > 59) Then
        MsgBox "Estimated announce time may not exceed 59 minutes.", 0, "Entry Exceeds 59 Minutes"
        txtMinute3 = ""
        txtMinute3.SetFocus
        Exit Sub
    End If

   If lblLinked.Visible = True Then 'break link to frmPlanner
        lblLinked_DblClick
    End If
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
    Check1(6).Value = 0
    Check1(7).Value = 0
  
    If txtMinute3 <> "" Then
    
        Check1(0).Enabled = False
        Check1(1).Enabled = False
        Check1(2).Enabled = False
        Check1(3).Enabled = False
        Check1(4).Enabled = False
        Check1(5).Enabled = False
        Check1(6).Enabled = False
        Check1(7).Enabled = False
 
        txtSpotsS.Visible = True
        lblSpots.Visible = True
        txtSpotLength.Visible = True
        Label3.Visible = True
  
        miAnncTime = 0
        giSpots = 0
        
        Frame8.Enabled = False
        'Frame8.Caption = ""
        
        txtBackAnnc.Visible = False
        chkBackAnnc.Value = 0
        chkBackAnnc.Visible = False
        Label21.Visible = False
       
        cmdClearAnncTimes.Enabled = False
    
        Label17.Visible = False
        Label24.Visible = False
        Label26.Visible = False
        txtCloseOut.Visible = False

        If txtMinute3 <> "" Then
           ' 'Label9.Alignment = 2
            Label9.Caption = "Estimated Announce Time (Double-click here to clear all. Double-click minutes to clear minutes, seconds to clear seconds)"
        Else
           ' 'Label9.Alignment = 2
            Label9.Caption = "Estimated Announce Time (double-click to clear)"
        End If
        
         If frmPlanner!mnuLinkTimeRemain.Checked = True Then
            frmPlanner!txtSpots.Visible = False
            frmPlanner!Shape1.Visible = False
            frmPlanner!lblSpotLength.Visible = False
            frmPlanner!lblAnnc.Visible = False
            frmPlanner!lblDate2.Visible = True
        End If
            
        chkAnnounce.Enabled = False
        chkAnnounce.Visible = False
        fraAnnounce.Caption = ""
  
    ElseIf txtMinute3 = "" And txtSecond3 = "" Then
                
        If chkAnnounce.Value = 1 Then
            txtSpotsS = ""
            
            txtSpotsS.Visible = False
            lblSpots.Visible = False
            txtSpotLength.Visible = False
            Label3.Visible = False
        Else
            
            txtIntro.Visible = True
            Frame12.Height = 1200
            txtAnnc4.Visible = True
            txtAnnc5.Visible = True
            txtAnnc6.Visible = True
            txtAnnc7.Visible = True
            txtAnnc8.Visible = True
            txtAnnc9.Visible = True
            txtAnnc10.Visible = True
            txtAnnc11.Visible = True
            
            Check1(0).Enabled = True
            Check1(1).Enabled = True
            Check1(2).Enabled = True
            Check1(3).Enabled = True
            Check1(4).Enabled = True
            Check1(5).Enabled = True
            Check1(6).Enabled = True
            Check1(7).Enabled = True
     
            txtBackAnnc.Enabled = True
            
            If txtMinute2 <> "" Then
            
                If Val(txtIntro) > 29 Then
                    txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
                Else
                    txtBackAnnc = "0"
                End If
                
            End If
            
            If txtCloseOut = "" Then
                txtCloseOut = "0"
            End If
         
            Label17.Visible = True
            Label21.Enabled = True
            Label22.Visible = True
            
            Label24.Visible = True
            Label26.Visible = True
            txtCloseOut.Visible = True
           
            lblAnncTime.Visible = True
            lblAverageTime.Visible = True
            
           ' 'Label9.Alignment = 0
            Label9.Caption = "You can replace the program's estimated announce time with your estimate of the announce time you will need:"
            
            Frame8.Enabled = True
            miCloseOut = Val(txtCloseOut)
''            'Frame8.Caption = Format(miCloseOut, " ##") & " sec allocated for Closeout && ID"
'            'Frame8.Caption = "Seconds Allocated for Closeout && ID"
            
            cmdClearAnncTimes.Enabled = True
            
            If lblEndTime.Visible = True Then
                txtBackAnnc.Visible = True
                chkBackAnnc.Value = 0
                chkBackAnnc.Visible = True
                Label21.Visible = True
            End If
            
            If frmPlanner!mnuLinkTimeRemain.Checked = True Then
                frmPlanner!txtSpots.Visible = True
                frmPlanner!Shape1.Visible = True
                frmPlanner!lblSpotLength.Visible = True
                frmPlanner!lblAnnc.Visible = True
                frmPlanner!lblDate2.Visible = False
            End If
            chkBackAnnc.Value = 0
        End If
        
        chkAnnounce.Enabled = True
        chkAnnounce.Value = 0
        chkAnnounce.Visible = True
        fraAnnounce.ForeColor = &H80& 'rust
        fraAnnounce.Caption = "Estimated Announce Time"
    End If
    
    '--------
    If txtSpotsS <> "" And txtSpotsS <> "0" And txtSecond3.Visible = True Then
        lblS.Visible = True
        
        lblSspot.Visible = True
        Label31.Visible = True
        If txtSpotsS = "1" Then
            lblS.Caption = txtSpotsS & " spot"
            Label31.Caption = txtSpotsS & " spot"
        Else
            lblS.Caption = txtSpotsS & " spots"
            Label31.Caption = txtSpotsS & " spots"
        End If
    Else
       lblS.Visible = False
       lblSspot.Visible = False
       Label31.Visible = False
       lblS.Caption = ""
       Label31.Caption = ""

    End If
    '--------
    pAnnounce
End Sub
Private Sub txtMinute3_DblClick()
    txtMinute3 = ""
    If txtSecond3 = "00" Or txtSecond3 = "0" Then
        txtSecond3 = ""
        
        Check1(0).Enabled = True
        Check1(1).Enabled = True
        Check1(2).Enabled = True
        Check1(3).Enabled = True
        Check1(4).Enabled = True
        Check1(5).Enabled = True
        Check1(6).Enabled = True
        Check1(7).Enabled = True
    End If
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
    Check1(6).Value = 0
    Check1(7).Value = 0
    
End Sub
Private Sub txtMinute3_GotFocus()
    txtMinute3.SelStart = 0 'begin selection at start
    txtMinute3.SelLength = Len(txtMinute3) 'selects # of characters
    txtSecond3.TabStop = True
End Sub
Private Sub txtMinute3_LostFocus()
    If txtMinute3 <> "" Then
        chkBackAnnc.Value = 0
    End If
End Sub

Private Sub txtMinute4_Change()

    If txtMinute4 <> "" Then
        lbl1.ForeColor = &H80& 'rust
    Else
        lbl1.ForeColor = &H808080 'gray
    End If

    If txtMinute4 <> "" Then
        txtMinute4.ToolTipText = " Double-Click to clear "
    Else
        txtMinute4.ToolTipText = "Enter CD playing time (minutes)"
    End If
    
    txtAnnc4 = Val(txtIntro) - Val(txtBackAnnc)
    
    If txtAnnc4 = Val(txtIntro) Then
        lblAnncTime = " Annc Time"
    End If

    If ck4Played.Value = 0 Then
        iMinute4 = txtMinute4
    End If
    
    If txtMinute4 = "" Then
        txtAnnc4 = ""
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl1.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtMinute4 <> frmPlanner!txtMinute1 Then
        txtMinute4.ForeColor = vbRed
        frmPlanner!txtMinute1.ForeColor = vbRed
    Else
        txtMinute4.ForeColor = &H80000012 'black
        frmPlanner!txtMinute1.ForeColor = &H80000012 'black
    End If
    
    If txtMinute4 <> "" And txtComposer4 <> "" And ck4Played.Enabled = False Then
        ck4Played.Enabled = True
    End If
    
    If txtMinute3 = "" And txtSecond3 = "" Then
        If txtMinute4 <> "" Or txtSecond4 <> "" Then
            Check1(0).Enabled = True
        Else
            Check1(0).Value = 0
            Check1(0).Enabled = False
        End If
    Else
        Check1(0).Enabled = False
    End If
    
    If txtMinute4 <> "" And txtComposer4 = "" Then
        txtComposer4 = "A"
    End If

    If ck4Played.Value = 0 And txtMinute4 = "" And txtSecond4 = "" Then
        txtComposer4 = ""
        txtComposer4.BackColor = &H80000005 ' white
    End If

    If txtMinute4 = "" And txtSecond4 = "" Then
        txtAnnc4.Text = ""
        txtAnnc4.Enabled = False
    Else
        txtAnnc4.Enabled = True
    End If
    
    If chkAnnounce.Value = 0 And txtMinute2 <> "" And txtMinute4 <> "" And txtMinute3 = "" And txtSecond3 = "" And F4Link = 0 Then
        txtBackAnnc.Visible = True
        chkBackAnnc.Visible = True
        Label21.Visible = True
        
        If Val(txtIntro) > 29 Then
            txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
        Else
            txtBackAnnc = "0"
        End If
    ElseIf txtMinute4 = "" Then
     
        txtBackAnnc.Visible = False
        chkBackAnnc.Visible = False
        Label21.Visible = False
        txtBackAnnc = ""
    End If

    pAnnounce
End Sub

Private Sub txtMinute4_DblClick()
    txtMinute4 = ""
    txtSecond4 = ""
       
    If txtComposer4 <> "" Then
        ck4Played.Enabled = False
    End If

End Sub

Private Sub txtMinute4_GotFocus()
    Annc4 = ""
    txtMinute4.SelStart = 0 'begin selection at start
    txtMinute4.SelLength = Len(txtMinute4)
End Sub

Private Sub txtMinute4_LostFocus()
    If txtMinute4 = "" And txtSecond4 = "00" Then
        txtSecond4 = ""
    End If
    
    If txtMinute4 = "" Then
        Check1(0).Value = 0
    End If
    
    If txtMinute4 <> "" And Len(txtComposer4) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
       cmdRestoreEntries.Enabled = True
       cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
       cmdRestoreEntries.Caption = "&Remove Trial Entries"
       cmdRestoreEntries.ToolTipText = " A trial entry is any entry that contains less than 3 characters in the Music Lineup text box "
       
    ElseIf Len(txtComposer4) >= 3 And cmdRestoreEntries.Caption <> "&Remove Trial Entries" Then
      cmdRestoreEntries.Enabled = False
    End If
    
End Sub

Private Sub txtMinute5_Change()

    If txtMinute5 <> "" Then
        lbl2.ForeColor = &H80& 'rust
    Else
        lbl2.ForeColor = &H808080 'gray
    End If

    If txtMinute5 <> "" Then
        txtMinute5.ToolTipText = " Double-Click to clear "
    Else
        txtMinute5.ToolTipText = "Enter CD playing time (minutes)"
    End If
    
    If txtMinute5 <> "" And Annc5 = "" And Check1(1).Value = 0 Then
        txtAnnc5 = txtIntro
    Else
        txtAnnc5 = Annc5
    End If

    If txtMinute5 = "" Then
        txtAnnc5 = ""
    End If

    If ck5Played.Value = 0 Then
        iMinute5 = txtMinute5 'Val(txtMinute5)
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl2.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtMinute5 <> frmPlanner!txtMinute2 Then
        txtMinute5.ForeColor = vbRed
        frmPlanner!txtMinute2.ForeColor = vbRed
    Else
        txtMinute5.ForeColor = &H80000012 'black
        frmPlanner!txtMinute2.ForeColor = &H80000012 'black
    End If
    
    If txtMinute5 <> "" And txtComposer5 <> "" And ck5Played.Enabled = False Then
        ck5Played.Enabled = True
    End If
    
If txtMinute3 = "" And txtSecond3 = "" Then
     If txtMinute5 <> "" Or txtSecond5 <> "" Then
        Check1(1).Enabled = True
    Else
        Check1(1).Value = 0
        Check1(1).Enabled = False
    End If
Else
    Check1(1).Enabled = False
End If
    
    If txtMinute5 <> "" And txtComposer5 = "" Then
        txtComposer5 = "B"
    End If
    
    If ck5Played.Value = 0 And txtMinute5 = "" And txtSecond5 = "" Then
        txtComposer5 = ""
        txtComposer5.BackColor = &H80000005 ' white
    End If
    
    If txtMinute5 = "" And txtSecond5 = "" Then
        txtAnnc5.Text = ""
        txtAnnc5.Enabled = False
    Else
        txtAnnc5.Enabled = True
    End If
        
    pAnnounce
End Sub

Private Sub txtMinute5_DblClick()
    txtMinute5 = ""
    txtSecond5 = ""
    
    If txtComposer5 <> "" Then
        ck5Played.Enabled = False
    End If
End Sub

Private Sub txtMinute5_GotFocus()
    Annc5 = ""
    txtMinute5.SelStart = 0 'begin selection at start
    txtMinute5.SelLength = Len(txtMinute5)
End Sub

Private Sub txtMinute5_LostFocus()
    If txtMinute5 = "" And txtSecond5 = "00" Then
        txtSecond5 = ""
    End If
    
    If txtMinute5 = "" Then
        Check1(1).Value = 0
    End If
    
    If txtMinute5 <> "" And Len(txtComposer5) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
        cmdRestoreEntries.ToolTipText = " A trial entry is any entry that contains less than 3 characters in the Music Lineup text box "
    
    ElseIf Len(txtComposer5) >= 3 And cmdRestoreEntries.Caption <> "&Remove Trial Entries" Then
      cmdRestoreEntries.Enabled = False
    End If
End Sub

Private Sub txtMinute6_Change()

    If txtMinute6 <> "" Then
        lbl3.ForeColor = &H80& 'rust
    Else
        lbl3.ForeColor = &H808080 'gray
    End If

    If txtMinute6 <> "" Then
        txtMinute6.ToolTipText = " Double-Click to clear "
    Else
        txtMinute6.ToolTipText = "Enter CD playing time (minutes)"
    End If

    If txtMinute6 <> "" And Annc6 = "" And Check1(2).Value = 0 Then
        txtAnnc6 = txtIntro
    Else
        txtAnnc6 = Annc6
    End If
    
    If txtMinute6 = "" Then
        txtAnnc6 = ""
    End If
    
    If ck6Played.Value = 0 Then
        iMinute6 = txtMinute6
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl3.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtMinute6 <> frmPlanner!txtMinute3 Then
        txtMinute6.ForeColor = vbRed
        frmPlanner!txtMinute3.ForeColor = vbRed
    Else
        txtMinute6.ForeColor = &H80000012 'black
        frmPlanner!txtMinute3.ForeColor = &H80000012 'black
    End If
    
    If txtMinute6 <> "" And txtComposer6 <> "" And ck6Played.Enabled = False Then
        ck6Played.Enabled = True
    End If
    
If txtMinute3 = "" And txtSecond3 = "" Then
    If txtMinute6 <> "" Or txtSecond6 <> "" Then
        Check1(2).Enabled = True
    Else
        Check1(2).Value = 0
        Check1(2).Enabled = False
    End If
Else
    Check1(2).Enabled = False
End If
    
    If txtMinute6 <> "" And txtComposer6 = "" Then
        txtComposer6 = "C"
    End If
    
    If ck6Played.Value = 0 And txtMinute6 = "" And txtSecond6 = "" Then
        txtComposer6 = ""
        txtComposer6.BackColor = &H80000005 ' white
    End If
    
    If txtMinute6 = "" And txtSecond6 = "" Then
        txtAnnc6.Text = ""
        txtAnnc6.Enabled = False
    Else
        txtAnnc6.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtMinute6_DblClick()
    txtMinute6 = ""
    txtSecond6 = ""
    
    If txtComposer6 <> "" Then
        ck6Played.Enabled = False
    End If
End Sub

Private Sub txtMinute6_GotFocus()
    Annc6 = ""
    txtMinute6.SelStart = 0 'begin selection at start
    txtMinute6.SelLength = Len(txtMinute6)
End Sub

Private Sub txtMinute6_LostFocus()
    If txtMinute6 = "" And txtSecond6 = "00" Then
        txtSecond6 = ""
    End If
    
    If txtMinute6 = "" Then
        Check1(2).Value = 0
    End If
    
    If txtMinute6 <> "" And Len(txtComposer6) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
        cmdRestoreEntries.ToolTipText = " A trial entry is any entry that contains less than 3 characters in the Music Lineup text box "
    
    ElseIf Len(txtComposer6) >= 3 And cmdRestoreEntries.Caption <> "&Remove Trial Entries" Then
      cmdRestoreEntries.Enabled = False
    End If
End Sub

Private Sub txtMinute7_Change()

    If txtMinute7 <> "" Then
        lbl4.ForeColor = &H80& 'rust
    Else
        lbl4.ForeColor = &H808080 'gray
    End If

    If txtMinute7 <> "" Then
        txtMinute7.ToolTipText = " Double-Click to clear "
    Else
        txtMinute7.ToolTipText = "Enter CD playing time (minutes)"
    End If

    If txtMinute7 <> "" And Annc7 = "" And Check1(3).Value = 0 Then
        txtAnnc7 = txtIntro
    Else
        txtAnnc7 = Annc7
    End If
    
    If txtMinute7 = "" Then
        txtAnnc7 = ""
    End If
    
    If ck7Played.Value = 0 Then
        iMinute7 = txtMinute7
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl4.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtMinute7 <> frmPlanner!txtMinute4 Then
        txtMinute7.ForeColor = vbRed
        frmPlanner!txtMinute4.ForeColor = vbRed
    Else
        txtMinute7.ForeColor = &H80000012 'black
        frmPlanner!txtMinute4.ForeColor = &H80000012 'black
    End If
    
    If txtMinute7 <> "" And txtComposer7 <> "" And ck7Played.Enabled = False Then
        ck7Played.Enabled = True
    End If
    
If txtMinute3 = "" And txtSecond3 = "" Then
    If txtMinute7 <> "" Or txtSecond7 <> "" Then
        Check1(3).Enabled = True
    Else
        Check1(3).Value = 0
        Check1(3).Enabled = False
    End If
Else
    Check1(3).Enabled = False
End If
    
    If txtMinute7 <> "" And txtComposer7 = "" Then
        txtComposer7 = "D"
    End If
    
    If ck7Played.Value = 0 And txtMinute7 = "" And txtSecond7 = "" Then
        txtComposer7 = ""
        txtComposer7.BackColor = &H80000005 ' white
    End If
    
    If txtMinute7 = "" And txtSecond7 = "" Then
        txtAnnc7.Text = ""
        txtAnnc7.Enabled = False
    Else
        txtAnnc7.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtMinute7_DblClick()
    txtMinute7 = ""
    txtSecond7 = ""
    
    If txtComposer7 <> "" Then
        ck7Played.Enabled = False
    End If
End Sub

Private Sub txtMinute7_GotFocus()
    Annc7 = ""
    txtMinute7.SelStart = 0 'begin selection at start
    txtMinute7.SelLength = Len(txtMinute7)
End Sub

Private Sub txtMinute7_LostFocus()
    If txtMinute7 = "" And txtSecond7 = "00" Then
        txtSecond7 = ""
    End If
    
    If txtMinute7 = "" Then
        Check1(3).Value = 0
    End If
    
    If txtMinute7 <> "" And Len(txtComposer7) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
        cmdRestoreEntries.ToolTipText = " A trial entry is any entry that contains less than 3 characters in the Music Lineup text box "
    
    ElseIf Len(txtComposer7) >= 3 And cmdRestoreEntries.Caption <> "&Remove Trial Entries" Then
      cmdRestoreEntries.Enabled = False
    End If
End Sub

Private Sub txtMinute8_Change()

    If txtMinute8 <> "" Then
        lbl5.ForeColor = &H80& 'rust
    Else
        lbl5.ForeColor = &H808080 'gray
    End If

    If txtMinute8 <> "" And Annc8 = "" And Check1(4).Value = 0 Then
        txtAnnc8 = txtIntro
    Else
        txtAnnc8 = Annc8
    End If
    
    If txtMinute8 = "" Then
        txtAnnc8 = ""
    End If
    
    If ck8Played.Value = 0 Then
        iMinute8 = txtMinute8
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl5.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtMinute8 <> frmPlanner!txtMinute5 Then
        txtMinute8.ForeColor = vbRed
        frmPlanner!txtMinute5.ForeColor = vbRed
    Else
        txtMinute8.ForeColor = &H80000012 'black
        frmPlanner!txtMinute5.ForeColor = &H80000012 'black
    End If
    
    If txtMinute8 <> "" And txtComposer8 <> "" And ck8Played.Enabled = False Then
        ck8Played.Enabled = True
    End If
    
If txtMinute3 = "" And txtSecond3 = "" Then
    If txtMinute8 <> "" Or txtSecond8 <> "" Then
        Check1(4).Enabled = True
    Else
        Check1(4).Value = 0
        Check1(4).Enabled = False
    End If
Else
    Check1(4).Enabled = False
End If
    
    If txtMinute8 <> "" And txtComposer8 = "" Then
        txtComposer8 = "E"
    End If
    
    If ck8Played.Value = 0 And txtMinute8 = "" And txtSecond8 = "" Then
        txtComposer8 = ""
        txtComposer8.BackColor = &H80000005 ' white
    End If
    
    If txtMinute8 = "" And txtSecond8 = "" Then
        txtAnnc8.Text = ""
        txtAnnc4.Enabled = False
    Else
        txtAnnc8.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtMinute8_DblClick()
    txtMinute8 = ""
    txtSecond8 = ""
    
    If txtComposer8 <> "" Then
        ck8Played.Enabled = False
    End If
End Sub

Private Sub txtMinute8_GotFocus()
    Annc8 = ""
    txtMinute8.SelStart = 0 'begin selection at start
    txtMinute8.SelLength = Len(txtMinute8)
End Sub

Private Sub txtMinute8_LostFocus()
    If txtMinute8 = "" And txtSecond8 = "00" Then
        txtSecond8 = ""
    End If
    
    If txtMinute8 = "" Then
        Check1(4).Value = 0
    End If
    
    If txtMinute8 <> "" And Len(txtComposer8) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
        cmdRestoreEntries.ToolTipText = " A trial entry is any entry that contains less than 3 characters in the Music Lineup text box "
        
    ElseIf Len(txtComposer8) >= 3 And cmdRestoreEntries.Caption <> "&Remove Trial Entries" Then
      cmdRestoreEntries.Enabled = False
    End If
End Sub

Private Sub txtMinute9_Change()

    If txtMinute9 <> "" Then
        lbl6.ForeColor = &H80& 'rust
    Else
        lbl6.ForeColor = &H808080 'gray
    End If

    If txtMinute9 <> "" Then
        lbl6.ForeColor = &H80& 'rust
    Else
        lbl6.ForeColor = &H808080 'gray
    End If

    If txtMinute9 <> "" Then
        txtMinute9.ToolTipText = " Double-Click to clear "
    Else
        txtMinute9.ToolTipText = "Enter CD playing time (minutes)"
    End If

    If txtMinute9 <> "" Then
        txtMinute9.ToolTipText = " Double-Click to clear "
    Else
        txtMinute9.ToolTipText = "Enter CD playing time (minutes)"
    End If

    If txtMinute9 <> "" And Annc9 = "" And Check1(5).Value = 0 Then
        txtAnnc9 = txtIntro
    Else
        txtAnnc9 = Annc9
    End If
    
    If txtMinute9 = "" Then
        txtAnnc9 = ""
    End If
    
    If ck9Played.Value = 0 Then
      iMinute9 = txtMinute9
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl6.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtMinute9 <> frmPlanner!txtMinute6 Then
        txtMinute9.ForeColor = vbRed
        frmPlanner!txtMinute6.ForeColor = vbRed
    Else
        txtMinute9.ForeColor = &H80000012 'black
        frmPlanner!txtMinute6.ForeColor = &H80000012 'black
    End If
    
    If txtMinute9 <> "" And txtComposer9 <> "" And ck9Played.Enabled = False Then
        ck9Played.Enabled = True
    End If
    
If txtMinute3 = "" And txtSecond3 = "" Then
    If txtMinute9 <> "" Or txtSecond9 <> "" Then
        Check1(5).Enabled = True
    Else
        Check1(5).Value = 0
        Check1(5).Enabled = False
    End If
Else
    Check1(5).Enabled = False
End If
    
    If txtMinute9 <> "" And txtComposer9 = "" Then
        txtComposer9 = "F"
    End If
    
    If ck9Played.Value = 0 And txtMinute9 = "" And txtSecond9 = "" Then
        txtComposer9 = ""
        txtComposer9.BackColor = &H80000005 ' white
    End If
    
    If txtMinute9 = "" And txtSecond9 = "" Then
        txtAnnc9.Text = ""
        txtAnnc9.Enabled = False
    Else
        txtAnnc9.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtMinute9_DblClick()
    txtMinute9 = ""
    txtSecond9 = ""
    
    If txtComposer9 <> "" Then
        ck9Played.Enabled = False
    End If
End Sub

Private Sub txtMinute9_GotFocus()
    Annc9 = ""
    txtMinute9.SelStart = 0 'begin selection at start
    txtMinute9.SelLength = Len(txtMinute9)
End Sub

Private Sub txtMinute9_LostFocus()
    If txtMinute9 = "" And txtSecond9 = "00" Then
        txtSecond9 = ""
    End If
    
    If txtMinute9 = "" Then
        Check1(5).Value = 0
    End If
    
    If txtMinute9 <> "" And Len(txtComposer9) < 3 And frmPlanner!mnuLinkTimeRemain.Checked = False Then
        cmdRestoreEntries.Enabled = True
        cmdRestoreEntries.BackColor = &H80000018  'tool tip yellow'&HFFFFFF    'white
        cmdRestoreEntries.Caption = "&Remove Trial Entries"
        cmdRestoreEntries.ToolTipText = " A trial entry is any entry that contains less than 3 characters in the Music Lineup text box "
        
    ElseIf Len(txtComposer9) >= 3 And cmdRestoreEntries.Caption <> "&Remove Trial Entries" Then
      cmdRestoreEntries.Enabled = False
    End If
End Sub

Private Sub txtSecAdj_DblClick()
    txtSecAdj.Text = ""
    
    If txtMinAdj = "" And txtSecAdj = "" Then
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
    End If
End Sub

Private Sub txtSecAdj_GotFocus()

    If Val(txtMinAdj) >= 0 Then 'begin selection at start
        txtSecAdj.SelStart = 0 'begin selection at start
        txtSecAdj.SelLength = Len(txtSecAdj)
    Else
        txtSecAdj.SelStart = 1
    End If
End Sub

Private Sub txtSecAdj_LostFocus()

    If txtSecAdj <> "" Then
        txtSecAdj.Text = Format(txtSecAdj, "00")
    End If
 
    txtSecAdj.TabStop = False

    If Val(txtMinAdj) < 0 And Val(txtSecAdj) > 0 Then
         txtSecAdj = Format$(txtSecAdj, "-00")
    End If

    If Val(txtMinAdj) > 0 And txtSecAdj = "" Then
        txtSecAdj = "00"
    End If

   If txtMinAdj.Text = "" And Val(txtSecAdj) = 0 Or txtSecAdj = "-" Then
        txtSecAdj = ""
    End If
      
    If Val(txtSecAdj) < 0 And Val(txtMinAdj) > 0 Then
       
       MsgBox "You have entered Minutes as a POSITIVE number and Seconds as a NEGATIVE number" & _
       vbCrLf & vbCrLf & "The two entries must agree as to case.", vbOKOnly, "Case MisMatch"
         txtSecAdj = ""
         
    ElseIf Val(txtSecAdj) > 0 And Val(txtMinAdj) < 0 Then 'this ElseIf is not likely to run since negative minutes clears seconds box if mismatch
        MsgBox "You have entered minutes as a negative number and seconds as a positive number" & _
        vbCrLf & vbCrLf & "The two entries must agree as to negative or positive case.", vbOKOnly, "Case Mismatch"
        txtSecAdj = ""
    End If
    
    If txtMinAdj <> "" Or txtSecAdj <> "" Then
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
    Else
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
    End If
    
     If txtSecAdj <> "" And txtSecAdj <= "9" Then
        txtSecAdj.Text = Format(txtSecAdj, "0")
     End If
 
End Sub

Private Sub txtSecond1_Change()

    Dim sSeconds As String 'to enter 0 seconds if txtSecond1 = ""
    Dim iMinute2 As Integer
    Dim iSecond2 As Integer
    
    If txtMinute2 <> "" Then
       iMinute2 = Val(txtMinute2)
    Else
       iMinute2 = 0
    End If
    
    If txtSecond2 <> "" Then
       iSecond2 = Val(txtSecond2)
    Else
       iSecond2 = 0
    End If
  
  '  lblHour.Visible = False 'manual entry clears setting time with system clock
  '  Shape1.Visible = True
    mCurrentTime = 0
    Frame7.Caption = "Set Current Time"

    cmdSetTime.Visible = False
    If chkStopWatch.Value = 0 Then
    
        If txtMinAdj = "" And txtSecAdj = "" Then
            cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock) F5"
        Else
             cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Computer Clock Adjusted)  F5"
        End If
        
    Else
        cmdSystemTime.Caption = "Click to &Set Current Time Past the Hour (Time Source Stopwatch) F5"
    End If
    
    If frmPlanner!mnuLinkTimeRemain.Checked = True And (txtMinute1 <> "" Or txtSecond1 <> "") Then
    
        If txtSecond1 <> "" Then
            sSeconds = txtSecond1
        Else
            sSeconds = "00"
        End If
        
        If lblCurrentTime.ForeColor = vbRed Then
            frmPlanner!staStatus.Panels(2) = "Timing begin in previous hour at time " & txtMinute1 & " min " & sSeconds & " sec"
        Else
           frmPlanner!staStatus.Panels(2) = "* Run time began at " & txtMinute1 & " min " & sSeconds & " sec past the hour"
        End If
    
    Else
        frmPlanner!staStatus.Panels(2) = ""
        frmPlanner!staStatus.Panels(3) = ""
    End If
    
    pAnnounce
End Sub

Private Sub txtSecond1_GotFocus()
    txtSecond1.SelStart = 0 'begin selection at start
    txtSecond1.SelLength = Len(txtSecond1)
End Sub

Private Sub txtSecond1_LostFocus()

    If (txtMinute1.Text = "" Or txtMinute1 = "00") And (txtSecond1 = "" Or txtSecond1 = "00") Then
        txtMinute1 = ""
        txtSecond1 = ""
    ElseIf (txtMinute1.Text <> "" And txtMinute1.Text <> "00") And txtSecond1 = "" Then
        txtSecond1 = "00"
    ElseIf txtMinute1.Text = "" And txtSecond1 <> "" Then
        txtMinute1.Text = "00"
    End If
    
    If txtMinute1 = "" And txtSecond1 = "" Then
        Label12.Visible = False
        imgOnAirSign.Visible = False
        imgDisc.Visible = False
        Label13.Visible = False
    End If
    
    txtSecond1.Text = Format$(txtSecond1, "00")
    
     'begin timing the next hour prior to end of current hour?
    If txtMinute1 >= "57" And txtMinute1 <= "59" And mCurrentTime = 0 Then
        Dim iResponse As Integer
        iResponse = MsgBox("You have entered the current time as " & txtMinute1 & " minutes and " & txtSecond1 & " seconds past the hour." _
        & vbCrLf & vbCrLf & "If this is an early beginning of the next hour's program click 'Yes' to begin timing the next hour from time " _
        & txtMinute1 & ":" & txtSecond1 & "." & vbCrLf & vbCrLf & _
        "If this is a continuation of the current hour's program click 'No' to remain within the current hour." & vbCrLf & vbCrLf & _
        "Begin a new hour of programming from time " & txtMinute1 & ":" & txtSecond1 & "?", vbYesNo + vbQuestion, _
        "Begin a new hour or remain within the current one?")

        If iResponse = vbYes Then
            lblCurrentTime_DblClick
        Else
            lblCurrentTime.ForeColor = vbBlack
            lblCurrentTime.Caption = " 1. Click the 'Set Current Time' button to enter the current time:"
        End If
    End If
    
    If chkBackAnnc.Value = 1 Then
        chkBackAnnc.Value = 0
    End If
    txtSecond1.TabStop = False

End Sub

Private Sub txtSecond10_GotFocus()
    txtSecond10.SelStart = 0 'begin selection at start
    txtSecond10.SelLength = Len(txtSecond10)
End Sub

Private Sub txtSecond10_LostFocus()
    If txtMinute10.Text <> "" And txtSecond10 = "" Then
        txtSecond10 = "00"
    End If

    txtSecond10.Text = Format$(txtSecond10, "00")
    If frmPlanner!mnuLinkTimeRemain.Checked = False And Len(txtComposer10) > 2 Then
        mnuFileSave_Click
    End If
    
End Sub

Private Sub txtSecond11_Change()

    If (Val(txtSecond11) > 59) Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond11 = ""
        Exit Sub
    End If

    If ck11Played.Value = 0 Then
        iSecond11 = txtSecond11
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl8.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtSecond11 <> frmPlanner!txtSecond8 Then
        txtSecond11.ForeColor = vbRed
        frmPlanner!txtSecond8.ForeColor = vbRed
    Else
        txtSecond11.ForeColor = &H80000012 'black
        frmPlanner!txtSecond8.ForeColor = &H80000012 'black
    End If
    
'If txtMinute11 <> "" Or txtSecond11 <> "" Then
'    Check1(7).Enabled = True
'Else
'    Check1(7).Value = 0
'    Check1(7).Enabled = False
'End If
'
    If ck11Played.Value = 0 And txtMinute11 = "" And txtSecond11 = "" Then
        txtComposer11 = ""
        txtComposer11.BackColor = &H80000005 ' white
    End If
    
If txtMinute3 = "" And txtSecond3 = "" Then
    If txtMinute11 <> "" Or txtSecond11 <> "" Then
        Check1(7).Enabled = True
    Else
        Check1(7).Value = 0
        Check1(7).Enabled = False
    End If
End If
    
    If txtMinute11 = "" And txtSecond11 = "" Then
        txtAnnc11.Text = ""
        txtAnnc11.Enabled = False
    Else
        txtAnnc11.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtSecond11_GotFocus()
    txtSecond11.SelStart = 0 'begin selection at start
    txtSecond11.SelLength = Len(txtSecond11)
End Sub

Private Sub txtSecond11_LostFocus()

    If txtMinute11.Text <> "" And txtSecond11 = "" Then
        txtSecond11 = "00"
    End If

    txtSecond11.Text = Format$(txtSecond11, "00")
    If frmPlanner!mnuLinkTimeRemain.Checked = False And Len(txtComposer11) > 2 Then
        mnuFileSave_Click
    End If

End Sub

Private Sub txtSecond2_Change()

    Dim iMinute2 As Integer
    Dim iSecond2 As Integer

    If txtMinute2 <> "" Then
       iMinute2 = Val(txtMinute2)
    Else
       iMinute2 = 0
    End If

    If txtSecond2 <> "" Then
       iSecond2 = Val(txtSecond2)
    Else
       iSecond2 = 0
    End If

 '-----------

    If iSecond2 > 0 Or iMinute2 > 0 Then

        If txtMinute1 <> "" Or txtSecond1 <> "" Then
            Label12.Visible = True
            imgOnAirSign.Visible = True
            imgDisc.Visible = True
            Label13.Visible = True
        Else
            Label12.Visible = False
            imgOnAirSign.Visible = False
            imgDisc.Visible = False
            Label13.Visible = False
        End If

    Else
        txtBackAnnc.Visible = False
        chkBackAnnc.Visible = False

        If chkBackAnnc.Value = 1 Then
            chkBackAnnc.Value = 0
        End If
        
        txtBackAnnc = "0"
        Label21.Visible = False
        Label12.Visible = False
        Label13.Visible = False
        lblEndTime.Visible = False
    End If

    If txtMinute2.Text = "" And txtSecond2 <> "" Then
        txtMinute2.Text = "00"
    ElseIf txtMinute2 = "00" And txtSecond2 = "" Then
        txtMinute2 = ""
    End If

    pAnnounce
End Sub

Private Sub txtSecond2_GotFocus()
    txtSecond2.SelStart = 0 'begin selection at start
    txtSecond2.SelLength = Len(txtSecond2)
End Sub

Private Sub txtSecond2_LostFocus() 'txtSecond2 is CD time remaining

    If txtMinute2.Text <> "" And txtSecond2 = "" Then
        txtSecond2 = "00"
    End If
        
    If txtMinute2.Text = "" And txtSecond2 = "" Then
        imgDisc.Visible = False
    End If
        
    If chkBackAnnc.Value = 1 Then
        chkBackAnnc.Value = 0
    End If
    
    txtSecond2.Text = Format$(txtSecond2, "00")
    
On Error GoTo HandleErrors
    If txtMinute1 <> "" And txtMinute2 <> "" Then 'saves for restore times command
        Open "RunTime.dat" For Output As #502
        Write #502, txtMinute1, txtMinute2, txtSecond1, txtSecond2, lblHour
        Close #502
    End If
    '----------------
    Dim SpotTime1 As Integer
    Dim SpotNumber1 As Currency
    Dim SpotTimeTotal As Integer
    SpotTime1 = Val(txtSpotLength)
    SpotNumber1 = Val(txtSpotsS)
    SpotTimeTotal = SpotTime1 * SpotNumber1
  
    If txtSpotsS <> "" And txtSpotsS <> "0" And cMusicMin < 30 And cMusicMin > 5 Then
    
        If Val(txtSpotsS) = 1 Then
        
            MsgBox txtSpotsS & "  (" & Val(txtSpotLength) & "-seconds average time) spot insert remains scheduled in the time period.", _
            vbOKOnly, SpotTimeTotal & " seconds spot-insert remains unplayed"
                      
        ElseIf Val(txtSpotsS) <> 1 Then
     
            MsgBox txtSpotsS & "  (" & Val(txtSpotLength) & "-seconds each average time) spot inserts remain scheduled in the time period.", _
            vbOKOnly, SpotTimeTotal & " seconds spot-inserts remain unplayed"
        End If
    End If
    
    If txtAnnc4 <> "" Then
        txtBackAnnc.Visible = True
        chkBackAnnc.Visible = True
        Label21.Visible = True
        txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
        txtAnnc4 = Val(txtIntro) - Val(txtBackAnnc)
    End If
     
    If txtComposer4 = "" Then
        txtComposer4.SetFocus
    ElseIf txtComposer5 = "" Then
        txtComposer5.SetFocus
    ElseIf txtComposer6 = "" Then
        txtComposer6.SetFocus
    ElseIf txtComposer7 = "" Then
        txtComposer7.SetFocus
    ElseIf txtComposer8 = "" Then
        txtComposer8.SetFocus
    ElseIf txtComposer9 = "" Then
        txtComposer9.SetFocus
    ElseIf txtComposer10 = "" Then
        txtComposer10.SetFocus
    ElseIf txtComposer11 = "" Then
        txtComposer11.SetFocus
    Else
        txtComposer4.SetFocus
    End If
       
    Exit Sub
'--------------
HandleErrors:
    Close #502
End Sub

Private Sub txtSecond3_Change()

    If mnuOptionsTime.Checked = True Then
       mnuOptionsTime.Checked = False
       fraIntro.Visible = False
    End If

    If lblLinked.Visible = True Then 'break link to frmPlanner
        lblLinked_DblClick
    End If
    
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
    Check1(6).Value = 0
    Check1(7).Value = 0
   
    If txtSecond3 <> "" Then
    
        Check1(0).Enabled = False
        Check1(1).Enabled = False
        Check1(2).Enabled = False
        Check1(3).Enabled = False
        Check1(4).Enabled = False
        Check1(5).Enabled = False
        Check1(6).Enabled = False
        Check1(7).Enabled = False
            
        txtSpotsS.Visible = True
        lblSpots.Visible = True
            
        txtSpotLength.Visible = True
        Label3.Visible = True
        
        miAnncTime = 0
        giSpots = 0
        Label17.Visible = False
       
        Label24.Visible = False
        Label26.Visible = False
        txtCloseOut.Visible = False
        txtBackAnnc.Visible = False
        chkBackAnnc.Value = 0
        chkBackAnnc.Visible = False
        Label21.Visible = False
       
        'Frame8.Caption = ""
        Frame8.Enabled = False
        chkBackAnnc.Value = 0
        cmdClearAnncTimes.Enabled = False
       
        chkAnnounce.Enabled = False
        chkAnnounce.Visible = False
        fraAnnounce.Caption = ""
        
        If txtMinute3 <> "" Then
           ' 'Label9.Alignment = 2
            Label9.Caption = "Estimated Announce Time (Double-click here to clear all or minutes to clear minutes, seconds to clear seconds)"
        Else
           ' 'Label9.Alignment = 2
            Label9.Caption = "Estimated Announce Time (double-click to clear)"
        End If
        
         If frmPlanner!mnuLinkTimeRemain.Checked = True Then
            frmPlanner!txtSpots.Visible = False
            frmPlanner!Shape1.Visible = False
            frmPlanner!lblSpotLength.Visible = False
            frmPlanner!lblAnnc.Visible = False
            frmPlanner!lblDate2.Visible = True
        End If
        
    ElseIf txtMinute3 = "" And txtSecond3 = "" Then
    
        Check1(0).Enabled = True
        Check1(1).Enabled = True
        Check1(2).Enabled = True
        Check1(3).Enabled = True
        Check1(4).Enabled = True
        Check1(5).Enabled = True
        Check1(6).Enabled = True
        Check1(7).Enabled = True
     
        If chkAnnounce.Value = 1 Then
            txtSpotsS = ""
            
            txtSpotsS.Visible = False
            lblSpots.Visible = False
            txtSpotLength.Visible = False
            Label3.Visible = False
        Else
       
            lblAverageTime.Visible = True
            txtIntro.Visible = True
            Frame12.Height = 1200
            txtAnnc4.Visible = True
            txtAnnc5.Visible = True
            txtAnnc6.Visible = True
            txtAnnc7.Visible = True
            txtAnnc8.Visible = True
            txtAnnc9.Visible = True
            txtAnnc10.Visible = True
            txtAnnc11.Visible = True
            txtBackAnnc.Enabled = True
            
            If txtMinute2 <> "" Then
            
                If Val(txtIntro) > 29 Then
                    txtBackAnnc = Format((Val((txtIntro) / 2) - 10), "##")
                Else
                    txtBackAnnc = "0"
                End If
                
            End If
                
            If txtCloseOut = "" Then
                txtCloseOut = "0"
            
            End If
            txtCloseOut.Visible = True
            Label24.Visible = True
            Label26.Visible = True
    
            Label22.Visible = True
            Label17.Visible = True
            Label21.Enabled = True
            lblAnncTime.Visible = True
            
            'Label9.Alignment = 0
            Label9.Caption = "You can replace the program's estimated announce time with your estimate of the announce time you will need:"
          
            Frame8.Enabled = True
            miCloseOut = Val(txtCloseOut)
            
           ' 'Frame8.Caption = Format(miCloseOut, " ##") & " sec allocated for Closeout && ID"
            'Frame8.Caption = "Seconds Allocated for Closeout && ID"

            cmdClearAnncTimes.Enabled = True
    
            If lblEndTime.Visible = True Then
                txtBackAnnc.Visible = True
                chkBackAnnc.Value = 0
                chkBackAnnc.Visible = True
                Label21.Visible = True
            End If
            
            If frmPlanner!mnuLinkTimeRemain.Checked = True Then
                frmPlanner!txtSpots.Visible = True
                frmPlanner!Shape1.Visible = True
                frmPlanner!lblSpotLength.Visible = True
                frmPlanner!lblAnnc.Visible = True
                frmPlanner!lblDate2.Visible = False
            End If
            chkBackAnnc.Value = 0
        End If
        
            chkAnnounce.Enabled = True
            chkAnnounce.Value = 0
            chkAnnounce.Visible = True
            fraAnnounce.ForeColor = &H80& 'rust
            fraAnnounce.Caption = "Estimated Announce Time"
    End If
    
    '--------
    If txtSpotsS <> "" And txtSpotsS <> "0" And txtSecond3.Visible = True Then
        lblS.Visible = True
        lblSspot.Visible = True
        Label31.Visible = True
       
        If txtSpotsS = "1" Then
            lblS.Caption = txtSpotsS & " spot"
            Label31.Caption = txtSpotsS & " spot"
        Else
            lblS.Caption = txtSpotsS & " spots"
            Label31.Caption = txtSpotsS & " spots"
        End If
    Else
       lblS.Visible = False
       lblSspot.Visible = False
       Label31.Visible = False
       lblS.Caption = ""
       Label31.Caption = ""
    End If
    
    pAnnounce
End Sub
Private Sub txtSecond3_DblClick()
    txtSecond3 = ""
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    Check1(3).Value = 0
    Check1(4).Value = 0
    Check1(5).Value = 0
    Check1(6).Value = 0
    Check1(7).Value = 0
    
    If txtSecond3 = "" And txtMinute3 = "" Then
        Check1(0).Enabled = True
        Check1(1).Enabled = True
        Check1(2).Enabled = True
        Check1(3).Enabled = True
        Check1(4).Enabled = True
        Check1(5).Enabled = True
        Check1(6).Enabled = True
        Check1(7).Enabled = True
    End If
    
    End Sub
Private Sub txtSecond3_GotFocus()
    txtSecond3.SelStart = 0 'begin selection at start
    txtSecond3.SelLength = Len(txtSecond3) 'selects # of characters
    If txtMinute3 <> "" Then
    'Label9.Alignment = 2
    Label9.Caption = "Estimated Announce Time (Double-click here to clear all or minutes to clear minutes, seconds to clear seconds)"
    End If
End Sub

Private Sub txtSecond3_LostFocus()
    If txtSecond3 <> "" Then
        txtSecond3 = Format$(txtSecond3, "00")
        chkBackAnnc.Value = 0
        txtBackAnnc.Text = ""
    End If
    
'    If txtMinute3 <> "" And txtSecond3 <> "" And txtSpotsS.Visible = True And Val(txtSpotsS) = 1 And cMusicMin < 10 Then
'        MsgBox "Reminder, " & Val(txtSpotsS) & " spot insert is listed as unplayed", vbOKOnly, "Unplayed Spot"
'    ElseIf txtSpotsS.Visible = True And Val(txtSpotsS) > 1 And cMusicMin < 10 Then
'        MsgBox "Reminder, " & Val(txtSpotsS) & " spot inserts are listed as unplayed", vbOKOnly, "Unplayed Spots"
'    End If
    
End Sub

Private Sub txtSecond4_Change()

    If ck4Played.Value = 0 Then
        iSecond4 = txtSecond4
    End If
    If mnuToolsExportLineCopy.Checked = True Then
        lbl1.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtSecond4 <> frmPlanner!txtSecond1 Then 'And ck4Played.Value = 0 Then
        txtSecond4.ForeColor = vbRed
        frmPlanner!txtSecond1.ForeColor = vbRed
    Else
        txtSecond4.ForeColor = &H80000012 'black
        frmPlanner!txtSecond1.ForeColor = &H80000012 'black
    End If
    
    If txtMinute3 = "" And txtSecond3 = "" Then
         If txtMinute4 <> "" Or txtSecond4 <> "" Then
            Check1(0).Enabled = True
        Else
            Check1(0).Value = 0
            Check1(0).Enabled = False
        End If
    End If
    
    If ck4Played.Value = 0 And txtMinute4 = "" And txtSecond4 = "" Then
        txtComposer4 = ""
        txtComposer4.BackColor = &H80000005 ' white
    End If
    
    If txtMinute4 = "" And txtSecond4 = "" Then
        txtAnnc4.Text = ""
        txtAnnc4.Enabled = False
    Else
        txtAnnc4.Enabled = True
    End If

    pAnnounce
End Sub

Private Sub txtSecond4_GotFocus()
    txtSecond4.SelStart = 0 'begin selection at start
    txtSecond4.SelLength = Len(txtSecond4)
End Sub

Private Sub txtSecond4_LostFocus()
    If txtMinute4.Text <> "" And txtSecond4 = "" Then
        txtSecond4 = "00"
    End If

    txtSecond4.Text = Format$(txtSecond4, "00")
    If frmPlanner!mnuLinkTimeRemain.Checked = False And Len(txtComposer4) > 2 Then 'saves entry if composer name > 2
        mnuFileSave_Click
    End If
        
End Sub
Private Sub txtSecond5_Change()
   
    If ck5Played.Value = 0 Then
        iSecond5 = txtSecond5
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl2.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtSecond5 <> frmPlanner!txtSecond2 Then
        txtSecond5.ForeColor = vbRed
        frmPlanner!txtSecond2.ForeColor = vbRed
    Else
        txtSecond5.ForeColor = &H80000012 'black
        frmPlanner!txtSecond2.ForeColor = &H80000012 'black
    End If
        
    If txtMinute3 = "" And txtSecond3 = "" Then
        If txtMinute5 <> "" Or txtSecond5 <> "" Then
            Check1(1).Enabled = True
        Else
            Check1(1).Value = 0
            Check1(1).Enabled = False
        End If
    End If
    
    If ck5Played.Value = 0 And txtMinute5 = "" And txtSecond5 = "" Then
        txtComposer5 = ""
        txtComposer5.BackColor = &H80000005 ' white
    End If
    
    If txtMinute5 = "" And txtSecond5 = "" Then
        txtAnnc5.Text = ""
        txtAnnc5.Enabled = False
    Else
        txtAnnc5.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtSecond5_GotFocus()
    txtSecond5.SelStart = 0 'begin selection at start
    txtSecond5.SelLength = Len(txtSecond5)
End Sub

Private Sub txtSecond5_LostFocus()
    If txtMinute5.Text <> "" And txtSecond5 = "" Then
        txtSecond5 = "00"
    End If

    txtSecond5.Text = Format$(txtSecond5, "00")
    If frmPlanner!mnuLinkTimeRemain.Checked = False And Len(txtComposer5) > 2 Then
        mnuFileSave_Click
    End If
 
End Sub

Private Sub txtSecond6_Change()

    If ck6Played.Value = 0 Then
        iSecond6 = txtSecond6
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl3.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtSecond6 <> frmPlanner!txtSecond3 Then
        txtSecond6.ForeColor = vbRed
        frmPlanner!txtSecond3.ForeColor = vbRed
    Else
        txtSecond6.ForeColor = &H80000012 'black
        frmPlanner!txtSecond3.ForeColor = &H80000012 'black
    End If
    
    If txtMinute3 = "" And txtSecond3 = "" Then
        If txtMinute6 <> "" Or txtSecond6 <> "" Then
            Check1(2).Enabled = True
        Else
            Check1(2).Value = 0
            Check1(2).Enabled = False
        End If
    End If
    
    If ck6Played.Value = 0 And txtMinute6 = "" And txtSecond6 = "" Then
        txtComposer6 = ""
        txtComposer6.BackColor = &H80000005 ' white
    End If
    
    If txtMinute6 = "" And txtSecond6 = "" Then
        txtAnnc6.Text = ""
        txtAnnc6.Enabled = False
    Else
        txtAnnc6.Enabled = True
    End If

    pAnnounce
End Sub

Private Sub txtSecond6_GotFocus()
    txtSecond6.SelStart = 0 'begin selection at start
    txtSecond6.SelLength = Len(txtSecond6)
End Sub

Private Sub txtSecond6_LostFocus()
    If txtMinute6.Text <> "" And txtSecond6 = "" Then
        txtSecond6 = "00"
    End If
    
    txtSecond6.Text = Format$(txtSecond6, "00")
    If frmPlanner!mnuLinkTimeRemain.Checked = False And Len(txtComposer6) > 2 Then
        mnuFileSave_Click
    End If

End Sub

Private Sub txtSecond7_Change()
    If ck7Played.Value = 0 Then
        iSecond7 = txtSecond7
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl4.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtSecond7 <> frmPlanner!txtSecond4 Then
        txtSecond7.ForeColor = vbRed
        frmPlanner!txtSecond4.ForeColor = vbRed
    Else
        txtSecond7.ForeColor = &H80000012 'black
        frmPlanner!txtSecond4.ForeColor = &H80000012 'black
    End If
    
    If txtMinute3 = "" And txtSecond3 = "" Then
        If txtMinute7 <> "" Or txtSecond7 <> "" Then
            Check1(3).Enabled = True
        Else
            Check1(3).Value = 0
            Check1(3).Enabled = False
        End If
    End If
    
    If ck7Played.Value = 0 And txtMinute7 = "" And txtSecond7 = "" Then
        txtComposer7 = ""
        txtComposer7.BackColor = &H80000005 ' white
    End If
    
    If txtMinute7 = "" And txtSecond7 = "" Then
        txtAnnc7.Text = ""
        txtAnnc7.Enabled = False
    Else
        txtAnnc7.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtSecond7_GotFocus()
    txtSecond7.SelStart = 0 'begin selection at start
    txtSecond7.SelLength = Len(txtSecond7)
End Sub

Private Sub txtSecond7_LostFocus()
    If txtMinute7.Text <> "" And txtSecond7 = "" Then
        txtSecond7 = "00"
    End If
    
    txtSecond7.Text = Format$(txtSecond7, "00")
    If frmPlanner!mnuLinkTimeRemain.Checked = False And Len(txtComposer7) > 2 Then
        mnuFileSave_Click
    End If

End Sub

Private Sub txtSecond8_Change()
    If ck8Played.Value = 0 Then
        iSecond8 = txtSecond8
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl5.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtSecond8 <> frmPlanner!txtSecond5 Then
        txtSecond8.ForeColor = vbRed
        frmPlanner!txtSecond5.ForeColor = vbRed
    Else
        txtSecond8.ForeColor = &H80000012 'black
        frmPlanner!txtSecond5.ForeColor = &H80000012 'black
    End If
    
    If txtMinute3 = "" And txtSecond3 = "" Then
        If txtMinute8 <> "" Or txtSecond8 <> "" Then
            Check1(4).Enabled = True
        Else
            Check1(4).Value = 0
            Check1(4).Enabled = False
        End If
    End If
    
    If ck8Played.Value = 0 And txtMinute8 = "" And txtSecond8 = "" Then
        txtComposer8 = ""
        txtComposer8.BackColor = &H80000005 ' white
    End If
    
    If txtMinute8 = "" And txtSecond8 = "" Then
        txtAnnc8.Text = ""
        txtAnnc8.Enabled = False
    Else
        txtAnnc8.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtSecond8_GotFocus()
    txtSecond8.SelStart = 0 'begin selection at start
    txtSecond8.SelLength = Len(txtSecond8)
End Sub

Private Sub txtSecond8_LostFocus()
    If txtMinute8.Text <> "" And txtSecond8 = "" Then
        txtSecond8 = "00"
    End If

    txtSecond8.Text = Format$(txtSecond8, "00")
    If frmPlanner!mnuLinkTimeRemain.Checked = False And Len(txtComposer8) > 2 Then
        mnuFileSave_Click
    End If

End Sub

Private Sub txtSecond9_Change()
    If ck9Played.Value = 0 Then
        iSecond9 = txtSecond9
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl6.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtSecond9 <> frmPlanner!txtSecond6 Then
        txtSecond9.ForeColor = vbRed
        frmPlanner!txtSecond6.ForeColor = vbRed
    Else
        txtSecond9.ForeColor = &H80000012 'black
        frmPlanner!txtSecond6.ForeColor = &H80000012 'black
    End If
    
    If txtMinute3 = "" And txtSecond3 = "" Then
        If txtMinute9 <> "" Or txtSecond9 <> "" Then
            Check1(5).Enabled = True
        Else
            Check1(5).Value = 0
            Check1(5).Enabled = False
        End If
    End If
    
    If ck9Played.Value = 0 And txtMinute9 = "" And txtSecond9 = "" Then
        txtComposer9 = ""
        txtComposer9.BackColor = &H80000005 ' white
    End If
    
    If txtMinute9 = "" And txtSecond9 = "" Then
        txtAnnc9.Text = ""
        txtAnnc9.Enabled = False
    Else
        txtAnnc9.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtSecond10_Change()
    If ck10Played.Value = 0 Then
        iSecond10 = txtSecond10
    End If
    
    If mnuToolsExportLineCopy.Checked = True Then
        lbl7.BorderStyle = 0
    End If
    
    If lblLinked.Visible = True And txtSecond10 <> frmPlanner!txtSecond7 Then
        txtSecond10.ForeColor = vbRed
        frmPlanner!txtSecond7.ForeColor = vbRed
    Else
        txtSecond10.ForeColor = &H80000012 'black
        frmPlanner!txtSecond7.ForeColor = &H80000012 'black
    End If
    
'If txtMinute10 <> "" Or txtSecond10 <> "" Then
'    Check1(6).Enabled = True
'Else
'    Check1(6).Value = 0
'    Check1(6).Enabled = False
'End If

If txtMinute3 = "" And txtSecond3 = "" Then
    If txtMinute10 <> "" Or txtSecond10 <> "" Then
        Check1(6).Enabled = True
    Else
        Check1(6).Value = 0
        Check1(6).Enabled = False
    End If
End If

    If ck10Played.Value = 0 And txtMinute10 = "" And txtSecond10 = "" Then
        txtComposer10 = ""
        txtComposer10.BackColor = &H80000005 ' white
    End If
    
    If txtMinute10 = "" And txtSecond10 = "" Then
        txtAnnc10.Text = ""
        txtAnnc10.Enabled = False
    Else
        txtAnnc10.Enabled = True
    End If
    
    pAnnounce
End Sub

Private Sub txtSecond9_GotFocus()
    txtSecond9.SelStart = 0 'begin selection at start
    txtSecond9.SelLength = Len(txtSecond9)
End Sub

Private Sub txtSecond9_LostFocus()
    If txtMinute9.Text <> "" And txtSecond9 = "" Then
        txtSecond9 = "00"
    End If
    
    txtSecond9.Text = Format$(txtSecond9, "00")
    If frmPlanner!mnuLinkTimeRemain.Checked = False And Len(txtComposer9) > 2 Then
        mnuFileSave_Click
    End If
 
End Sub

Private Sub txtSpotLength_Change()

    If Not IsNumeric(txtSpotLength) And txtSpotLength <> "" Then
         MsgBox "You have entered the non-numeric character:  " & txtSpotLength & vbCrLf & vbCrLf & _
         "Enter in seconds the average length of promos," & vbCrLf & "spot announcements, weather inserts, etc.", vbOKOnly, "Entry Error"
         txtSpotLength = "30"
         txtSpotLength.SetFocus
         Exit Sub
    End If
    
    If Val(txtSpotLength) > 180 Then
        MsgBox "Enter in seconds the time allotted for each spot announcement." _
        & vbCrLf & vbCrLf & "Entry can range from 0 to a maximum of 180 seconds (which is 3 minutes).", _
        vbOKOnly, "Entry Greater than 180 Seconds"
        txtSpotLength = "30"
        txtSpotLength.SetFocus
        Exit Sub
    End If

    If lblRemain30.Visible = True Then
        lblSpots.Caption = "Enter the number of (" & Val(txtSpotLength) & _
        "-second average time) spot, promo, PSA, weather, etc. inserts REMAINING in the hour (or half-hour) time period"
    Else
        lblSpots.Caption = "Enter the number of (" & Val(txtSpotLength) & _
        "-second average time) spot, promo, PSA, weather, etc. inserts REMAINING in the hour"
    End If
 
      If txtSpotLength <> "" Then
        txtSpotLengthSetting = txtSpotLength
    End If
    
    If txtSpotsS = "" Then
        lblSpotSecs = ""
    Else
        lblSpotSecs = Val(txtSpotsS) * Val(txtSpotLength) & " secs"
    End If
    
    If Val(txtSpotLength) >= 20 Then
        txtSpotLength.ToolTipText = "Double-click to reduce spot length by 5 seconds"
    Else
        txtSpotLength.ToolTipText = ""
    End If
    
    pAnnounce
End Sub

Private Sub txtSpotLength_DblClick()
    If Val(txtSpotLength) > 15 Then
        txtSpotLength = Format((Val(txtSpotLength) - 5), "##")
        txtSpotLength.SelStart = 0 'begin selection at start
        txtSpotLength.SelLength = Len(txtSpotLength)
        txtSpotLength.SetFocus
     ElseIf Val(txtSpotLength) <= 15 Then
        txtSpotLength = frmDefaults!txtSpot
     End If
End Sub

Private Sub txtSpotLength_GotFocus()
    txtSpotLength.SelStart = 0 'begin selection at start
    txtSpotLength.SelLength = Len(txtSpotLength)
    Label3.ForeColor = &H80&
End Sub

Private Sub txtSpotLength_LostFocus()
    If Not IsNumeric(txtSpotLength) Or txtSpotLength = "" Then
        txtSpotLength = "30"
    Else
        txtSpotLength.Text = Format$(txtSpotLength, "##")
    End If
    Label3.ForeColor = &H404040
End Sub

Private Sub txtSpotLengthSetting_Change()

    If Not IsNumeric(txtSpotLengthSetting) And txtSpotLengthSetting <> "" Then
    
        MsgBox "You have entered the non-numeric character:  " & txtSpotLengthSetting & vbCrLf & vbCrLf & _
        "Enter in seconds the average length of promos," & vbCrLf & _
        "spot announcements, weather inserts, etc.", vbOKOnly, "Entry Error"
        
        txtSpotLengthSetting = "30"
        txtSpotLengthSetting.SetFocus
        Exit Sub
    End If

    If txtSpotLengthSetting <> "" Then
        txtSpotLength = txtSpotLengthSetting
    End If
End Sub

Private Sub txtSpotLengthSetting_GotFocus()
    txtSpotLengthSetting.SelStart = 0 'begin selection at start
    txtSpotLengthSetting.SelLength = Len(txtSpotLengthSetting)
End Sub

Private Sub txtSpotLengthSetting_LostFocus()
    If txtSpotLengthSetting = "" Then
        txtSpotLengthSetting = "30"
    Else
        txtSpotLengthSetting.Text = Format$(txtSpotLengthSetting, "##")
    End If
End Sub

Private Sub txtSpotsS_Change()

    If Not IsNumeric(txtSpotsS) And txtSpotsS <> "" Then
         MsgBox "You have entered a non-numeric value", vbOKOnly, "Entry Error"
         txtSpotsS = ""
         txtSpotsS.SetFocus
         Exit Sub
    End If
'-------------
    If Val(txtSpotLength) <= 35 Then
    
        If Val(txtSpotsS) > 10 Then
             MsgBox "You have entered " & txtSpotsS & " spot/inserts. Ten is the maximum for an hour.", vbOKOnly, "Excessive Number of Spot / PSA / Wx Inserts"
             txtSpotsS = ""
             If frmTimeRemain.Visible = True And txtSpotsS.Visible = True Then
                txtSpotsS.SetFocus
             End If
             Exit Sub
        End If
    
    ElseIf Val(txtSpotLength) > 35 And Val(txtSpotLength) <= 65 Then
    
        If Val(txtSpotsS) > 8 Then
        
            MsgBox "You have entered " & txtSpotsS & " spot/inserts. Eight is the maximum for an hour.", vbOKOnly, "Excessive Number of Spot / PSA / Wx Inserts"
            txtSpotsS = ""
            If frmTimeRemain.Visible = True And txtSpotsS.Visible = True Then
               txtSpotsS.SetFocus
            End If
            Exit Sub
        End If
    
    ElseIf Val(txtSpotLength) > 65 Then
    
        If Val(txtSpotsS) > 6 Then
        
            MsgBox "You have entered " & txtSpotsS & " spot/inserts. Six is the maximum for an hour.", vbOKOnly, "Excessive Number of Spot / PSA / Wx Inserts"
            txtSpotsS = ""
            If frmTimeRemain.Visible = True And txtSpotsS.Visible = True Then
               txtSpotsS.SetFocus
            End If
            Exit Sub
        End If
        
    End If
'--------------
    If txtSpotsS <> "" And txtSpotsS <> "0" Then
        lblSpots.ForeColor = vbBlue
        txtSpotsS.ToolTipText = " Double-click to reduce the number of spots, PSA's, etc. by one "
    Else
        lblSpots.ForeColor = &H80&       'rust
        txtSpotsS.ToolTipText = " Enter the number of PSA's and spots inserted in the time period "
    End If
    
    If txtSpotsS <> "" And txtSpotsS <> "0" And txtSecond3.Visible = True Then
        lblS.Visible = True
        lblSspot.Visible = True
        Label31.Visible = True
        If txtSpotsS = "1" Then
            lblS.Caption = txtSpotsS & " spot"
            Label31.Caption = txtSpotsS & " spot"
        Else
            lblS.Caption = txtSpotsS & " spots"
            Label31.Caption = txtSpotsS & " spots"
        End If
    Else
       lblS.Visible = False
       lblSspot.Visible = False
       Label31.Visible = False
       lblS.Caption = ""
       Label31.Caption = ""
    End If

    If frmPlanner!mnuLinkTimeRemain.Checked = True Then
       frmPlanner!txtSpots = txtSpotsS
    End If
       
    If txtSpotsS = "" Or Val(txtSpotsS) = 0 Then
        lblSpotSecs = ""
    Else
        lblSpotSecs = Val(txtSpotsS) * Val(txtSpotLength) & " secs"
    End If
  
   pAnnounce
End Sub

Private Sub txtSpotsS_DblClick()
    If txtSpotsS > "1" Then
        txtSpotsS = Format((Val(txtSpotsS) - 1), "##")
    Else
        txtSpotsS = ""
    End If
    
    txtSpotsS.SelStart = 0 'begin selection at start
    txtSpotsS.SelLength = Len(txtSpotsS)
    txtSpotsS.SetFocus
End Sub

Private Sub txtSpotsS_GotFocus()
    txtSpotsS.SelStart = 0 'begin selection at start
    txtSpotsS.SelLength = Len(txtSpotsS)
End Sub

Private Sub txtSpotsS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then 'comma
        KeyAscii = 46 'period
    End If
End Sub

Private Sub txtSpotsS_LostFocus()
    If txtSpotsS = "0" Or txtSpotsS = "0." Or txtSpotsS = "0.0" Then
        txtSpotsS = ""
    End If
End Sub
