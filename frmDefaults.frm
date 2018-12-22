VERSION 5.00
Begin VB.Form frmDefaults 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7830
   ClientLeft      =   345
   ClientTop       =   945
   ClientWidth     =   12660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDefaults.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Exit"
      ForeColor       =   &H00000080&
      Height          =   1350
      Left            =   5040
      TabIndex        =   66
      Top             =   4035
      Width           =   3135
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000018&
         Cancel          =   -1  'True
         Caption         =   "E&xit Page"
         Default         =   -1  'True
         Height          =   405
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   255
         Width           =   2745
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Changes and Exit Page"
         Enabled         =   0   'False
         Height          =   405
         Left            =   210
         TabIndex        =   67
         Top             =   840
         Width           =   2745
      End
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Save Current Transmitter && Time Settings as &Defaults (access code required)"
      Height          =   570
      Left            =   4605
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5580
      Width           =   4005
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Disc&ontinue Setup"
      Height          =   315
      Left            =   5820
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   6285
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Station Call Letters && Transmitter Power Limits"
      ForeColor       =   &H00000080&
      Height          =   3135
      Left            =   4410
      TabIndex        =   47
      Top             =   750
      Width           =   4380
      Begin VB.TextBox txtStation4 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1380
         MaxLength       =   4
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   2385
         Width           =   855
      End
      Begin VB.TextBox txtStation3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1380
         MaxLength       =   4
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   1830
         Width           =   855
      End
      Begin VB.TextBox txtStation2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1380
         MaxLength       =   4
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   1275
         Width           =   855
      End
      Begin VB.TextBox txtStation1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1380
         MaxLength       =   4
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   705
         Width           =   855
      End
      Begin VB.TextBox txtStation1Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3450
         MaxLength       =   6
         TabIndex        =   40
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtStation2Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   41
         Top             =   1290
         Width           =   735
      End
      Begin VB.TextBox txtStation2Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3450
         MaxLength       =   6
         TabIndex        =   42
         Top             =   1290
         Width           =   735
      End
      Begin VB.TextBox txtStation3Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   43
         Top             =   1845
         Width           =   735
      End
      Begin VB.TextBox txtStation3Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3450
         MaxLength       =   6
         TabIndex        =   44
         Top             =   1845
         Width           =   735
      End
      Begin VB.TextBox txtStation4Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   45
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtStation4Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3450
         MaxLength       =   6
         TabIndex        =   46
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtStation1Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   39
         Top             =   705
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000C0&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   120
         Left            =   1215
         Shape           =   1  'Square
         Top             =   465
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label8 
         Caption         =   "Station 4 Optional"
         Height          =   390
         Left            =   570
         TabIndex        =   56
         Top             =   2430
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Station 3 Optional"
         Height          =   405
         Left            =   540
         TabIndex        =   55
         Top             =   1890
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Station 2 Optional"
         Height          =   405
         Left            =   540
         TabIndex        =   54
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "Station 1 Flagship Station (Required Entry)"
         ForeColor       =   &H00C00000&
         Height          =   600
         Left            =   135
         TabIndex        =   53
         Top             =   585
         Width           =   1170
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Maximum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3540
         TabIndex        =   52
         Top             =   465
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         X1              =   2535
         X2              =   2535
         Y1              =   495
         Y2              =   2865
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Watts"
         Height          =   240
         Left            =   3135
         TabIndex        =   51
         Top             =   2745
         Width           =   555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderStyle     =   3  'Dot
         X1              =   1485
         X2              =   2355
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Call Letters"
         Height          =   210
         Left            =   1410
         TabIndex        =   50
         Top             =   420
         Width           =   825
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Power Legal Limits"
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
         Left            =   2640
         TabIndex        =   49
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Minimum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2745
         TabIndex        =   48
         Top             =   465
         Width           =   705
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000C0&
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   1
         Left            =   2475
         Shape           =   1  'Square
         Top             =   300
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Average Times (Access Code Required)"
      ForeColor       =   &H00000080&
      Height          =   4170
      Left            =   435
      TabIndex        =   21
      Top             =   750
      Width           =   3690
      Begin VB.TextBox txtSpot 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2685
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   2460
         Width           =   765
      End
      Begin VB.TextBox txtClose 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2685
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   3255
         Width           =   765
      End
      Begin VB.TextBox txtPlanTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2685
         MaxLength       =   4
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   750
         Width           =   765
      End
      Begin VB.TextBox txtIntroOut 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2685
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Average length in seconds of spot, wx, etc. announcements"
         Height          =   480
         Left            =   210
         TabIndex        =   34
         Top             =   2460
         Width           =   2280
      End
      Begin VB.Label Label3 
         Caption         =   "Average seconds required for Show Closeout and Station ID"
         Height          =   420
         Left            =   210
         TabIndex        =   33
         Top             =   3255
         Width           =   2265
      End
      Begin VB.Label Label2 
         Caption         =   "Average total seconds required to Introduce and Back-Announce a music selection"
         Height          =   600
         Left            =   210
         TabIndex        =   32
         Top             =   1680
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   "Average minutes of music played in an hour program (Required Entry)"
         ForeColor       =   &H00C00000&
         Height          =   630
         Left            =   210
         TabIndex        =   31
         Top             =   735
         Width           =   2190
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   " Minutes"
         Height          =   165
         Left            =   2752
         TabIndex        =   30
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Seconds"
         Height          =   165
         Left            =   2745
         TabIndex        =   29
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Seconds"
         Height          =   165
         Left            =   2745
         TabIndex        =   28
         Top             =   3630
         Width           =   630
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Seconds"
         Height          =   165
         Left            =   2745
         TabIndex        =   27
         Top             =   2835
         Width           =   630
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         X1              =   345
         X2              =   3435
         Y1              =   1425
         Y2              =   1425
      End
      Begin VB.Label Label18 
         Caption         =   "For suggested average times, click ""Help"" menu"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         TabIndex        =   26
         Top             =   270
         Width           =   3495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Limits as  Percent of Ideal Power"
      ForeColor       =   &H00000080&
      Height          =   5835
      Left            =   9075
      TabIndex        =   2
      Top             =   750
      Width           =   3120
      Begin VB.TextBox txtIdeal1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   825
         TabIndex        =   5
         Top             =   3630
         Width           =   735
      End
      Begin VB.TextBox txtIdeal2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   825
         TabIndex        =   6
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtIdeal3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   825
         TabIndex        =   7
         Top             =   4530
         Width           =   735
      End
      Begin VB.TextBox txtIdeal4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   825
         TabIndex        =   8
         Top             =   4980
         Width           =   735
      End
      Begin VB.TextBox txtMin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   225
         MaxLength       =   2
         TabIndex        =   3
         Top             =   2445
         Width           =   450
      End
      Begin VB.CheckBox ckIdeal4 
         Alignment       =   1  'Right Justify
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   5010
         Width           =   345
      End
      Begin VB.CheckBox ckIdeal3 
         Alignment       =   1  'Right Justify
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   4560
         Width           =   345
      End
      Begin VB.CheckBox ckIdeal2 
         Alignment       =   1  'Right Justify
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   4110
         Width           =   345
      End
      Begin VB.CheckBox ckIdeal1 
         Alignment       =   1  'Right Justify
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   3660
         Width           =   345
      End
      Begin VB.TextBox txtMax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   225
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2775
         Width           =   450
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Ideal Power Watts"
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
         Height          =   405
         Left            =   735
         TabIndex        =   63
         Top             =   5310
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000C0&
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   120
         Index           =   0
         Left            =   150
         Shape           =   1  'Square
         Top             =   375
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblIdeal1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1830
         TabIndex        =   20
         Top             =   3615
         Width           =   765
      End
      Begin VB.Label lblIdeal2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1830
         TabIndex        =   19
         Top             =   4065
         Width           =   765
      End
      Begin VB.Label lblIdeal3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1830
         TabIndex        =   18
         Top             =   4515
         Width           =   765
      End
      Begin VB.Label lblIdeal4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1830
         TabIndex        =   17
         Top             =   4965
         Width           =   765
      End
      Begin VB.Label Label20 
         Caption         =   $"frmDefaults.frx":0442
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
         Height          =   1950
         Left            =   255
         TabIndex        =   16
         Top             =   330
         Width           =   2640
      End
      Begin VB.Label Label21 
         Caption         =   "%  Minimum.  (less than 100%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   690
         TabIndex        =   15
         Top             =   2490
         Width           =   2160
      End
      Begin VB.Label Label23 
         Caption         =   "Check  to set limits as % of ideal power"
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
         Left            =   75
         TabIndex        =   14
         Top             =   3225
         Width           =   2925
      End
      Begin VB.Label Label22 
         Caption         =   "%  Maximum.  (100% or more)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   690
         TabIndex        =   13
         Top             =   2820
         Width           =   2100
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Set Music Logbook Printout Signature Line  (access code required)"
      ForeColor       =   &H00000080&
      Height          =   795
      Left            =   690
      TabIndex        =   0
      Top             =   6735
      Width           =   7710
      Begin VB.TextBox txtSignatureLine 
         Height          =   285
         Left            =   3960
         MaxLength       =   55
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "Fine Arts Radio"
         Top             =   338
         Width           =   3570
      End
      Begin VB.CommandButton cmdSetSignature 
         Caption         =   "Set as Music Logbook Page Printout Default Signature"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   300
         Width           =   3600
      End
   End
   Begin VB.Image imgOnAirSign 
      Height          =   255
      Left            =   10185
      Picture         =   "frmDefaults.frx":057D
      Stretch         =   -1  'True
      Top             =   7200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   10073
      Picture         =   "frmDefaults.frx":0C5B
      Top             =   6765
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   120
      Picture         =   "frmDefaults.frx":23C5
      Top             =   6990
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "Saved"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   8550
      TabIndex        =   62
      ToolTipText     =   "Double-click to cancel signature change"
      Top             =   7320
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgSmile 
      Height          =   300
      Index           =   0
      Left            =   8715
      Picture         =   "frmDefaults.frx":2807
      Stretch         =   -1  'True
      ToolTipText     =   "Double-click to cancel signature change"
      Top             =   6975
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   2
      Left            =   195
      Shape           =   3  'Circle
      Top             =   270
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblHelp 
      Caption         =   $"frmDefaults.frx":2C49
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
      Height          =   495
      Left            =   540
      TabIndex        =   59
      Top             =   135
      Visible         =   0   'False
      Width           =   11715
   End
   Begin VB.Label Label9 
      Caption         =   "Set:  Average Times, Station Call Letters, Transmitter Power Legal Limits, Printout Signature Line &&  Default Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      TabIndex        =   58
      Top             =   210
      Width           =   10230
   End
   Begin VB.Image imgSmile 
      Height          =   300
      Index           =   1
      Left            =   4050
      Picture         =   "frmDefaults.frx":2D8D
      Stretch         =   -1  'True
      ToolTipText     =   "Double-Click to Reset"
      Top             =   5160
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblDefaults 
      Caption         =   $"frmDefaults.frx":31CF
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   585
      TabIndex        =   57
      Top             =   5055
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Menu mnuPage 
      Caption         =   "&Page"
      Begin VB.Menu mnuPageMusicLog 
         Caption         =   "&MusicLog page..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuPageTransmitter 
         Caption         =   "&Transmitter page..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrevious 
         Caption         =   "&Return to  Previous Page..."
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuHelpSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuRestoreDefaults 
      Caption         =   "&Restore Defaults"
      Begin VB.Menu mnuFileTimeDefaults 
         Caption         =   "Restore Intro/Back-Annc, Spot && Closeout Time Defaults"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileXmitterDefaults 
         Caption         =   "Restore Station Call Letters && Transmitter Power Limits Defaults"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTimes 
         Caption         =   "Suggested Average &Times && Power Limit Note"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpPrint 
         Caption         =   "&Print a Copy of this Page"
      End
   End
End
Attribute VB_Name = "frmDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      
    Dim mCount As Integer
    Dim iCount As Integer
    Dim iCount2 As Integer
    Dim iStation1min As Integer
    Dim iStation2min As Integer
    Dim iStation3min As Integer
    Dim iStation4min As Integer
    Dim iStation1max As Integer
    Dim iStation2max As Integer
    Dim iStation3max As Integer
    Dim iStation4max As Integer
    Dim iAccessOpen As Integer
    Dim iChange As Integer
    Dim iDefaultsSaved As Integer
    
    Dim iCodeRequired As Integer
 
Option Explicit

Private Sub ckIdeal1_Click()

    If ckIdeal1.Value = 1 And txtIdeal1 <> "" And txtMin <> "" And txtMax <> "" Then
        iStation1min = Int(txtIdeal1) * Int(txtMin) / 100
        txtStation1Min = iStation1min
        iStation1max = Int(txtIdeal1) * Int(txtMax) / 100
        txtStation1Max = iStation1max
    
    ElseIf ckIdeal1.Value = 1 Then
        MsgBox "Minimum & Maximum percentage power limits and Ideal Power (Watts) text entry boxes must have data", _
        vbOKOnly, "Missing Data"
        ckIdeal1.Value = 0
    End If
    
    If iChange = 0 Then
        iChange = 5
    End If
End Sub

Private Sub ckIdeal1_GotFocus()
    ckIdeal1.Caption = "*"
End Sub

Private Sub ckIdeal1_LostFocus()
    ckIdeal1.Caption = ""
End Sub

Private Sub ckIdeal2_Click()

    If ckIdeal2.Value = 1 And txtIdeal2 <> "" And txtMin <> "" And txtMax <> "" Then
        iStation2min = Int(txtIdeal2) * Int(txtMin) / 100
        txtStation2Min = iStation2min
        iStation2max = Int(txtIdeal2) * Int(txtMax) / 100
        txtStation2Max = iStation2max
        
    ElseIf ckIdeal2.Value = 1 Then
        MsgBox "Minimum & Maximum percentage power limits and Ideal Power (Watts) text entry boxes must have data", _
        vbOKOnly, "Missing Data"
        ckIdeal2.Value = 0
    End If
    
   If iChange = 0 Then
        iChange = 5
    End If
End Sub

Private Sub ckIdeal2_GotFocus()
    ckIdeal2.Caption = "*"
End Sub

Private Sub ckIdeal2_LostFocus()
    ckIdeal2.Caption = ""
End Sub

Private Sub ckIdeal3_Click()

    If ckIdeal3.Value = 1 Then

        If txtIdeal3 <> "" And txtMin <> "" And txtMax <> "" Then
            iStation3min = Int(txtIdeal3) * Int(txtMin) / 100
            txtStation3Min = iStation3min
            iStation3max = Int(txtIdeal3) * Int(txtMax) / 100
            txtStation3Max = iStation3max
    
        Else
            MsgBox "Minimum & Maximum percentage power limits and Ideal Power (Watts) text entry boxes must have data", _
            vbOKOnly, "Missing Data"
            ckIdeal3.Value = 0
        End If
    End If
    
    If iChange = 0 Then
        iChange = 5
    End If
End Sub

Private Sub ckIdeal3_GotFocus()
    ckIdeal3.Caption = "*"
End Sub

Private Sub ckIdeal3_LostFocus()
   ckIdeal3.Caption = ""
End Sub

Private Sub ckIdeal4_Click()

    If ckIdeal4.Value = 1 And txtIdeal4 <> "" And txtMin <> "" And txtMax <> "" Then
        iStation4min = Int(txtIdeal4) * Int(txtMin) / 100
        txtStation4Min = iStation4min
        iStation4max = Int(txtIdeal4) * Int(txtMax) / 100
        txtStation4Max = iStation4max
        
    ElseIf ckIdeal4.Value = 1 Then
        MsgBox "Minimum & Maximum percentage power limits and Ideal Power text entry boxes must have data", _
        vbOKOnly, "Missing Data"
        ckIdeal4.Value = 0
    End If
    
    If iChange = 0 Then
        iChange = 5
    End If
End Sub

Private Sub ckIdeal4_GotFocus()
    ckIdeal4.Caption = "*"
End Sub

Private Sub ckIdeal4_LostFocus()
    ckIdeal4.Caption = ""
End Sub

Private Sub cmdCancel_Click()
    imgSmile(0).Visible = False
    Label25.Visible = False

    Dim iResponse As Integer

    If iChange <> 0 And iDefaultsSaved = 0 Then
       ' Else
    
        Select Case iChange
        
            Case 1
                iResponse = MsgBox("Do want to save any changes made to transmitter ideal power settings?", vbYesNoCancel, "Confirm Exit")
            Case 2
                iResponse = MsgBox("Do want to save any changes made to transmitter minimum power settings?", vbYesNoCancel, "Confirm Exit")
            Case 3
                iResponse = MsgBox("Do want to save any changes made to transmitter maximum power settings?", vbYesNoCancel, "Confirm Exit")
            Case 4
                iResponse = MsgBox("Do want to save any changes made to minimum or maximum percentage of ideal power?", vbYesNoCancel, "Confirm Exit")
            Case 5
                iResponse = MsgBox("Do want to save any changes made to the minimum/maximum percentage power check boxes?", vbYesNoCancel, "Confirm Exit")
                  
        End Select
        
        iChange = 0
        If iResponse = vbNo Then
        ElseIf iResponse = vbYes Then
            cmdSave_Click
        ElseIf iResponse = vbCancel Then
            iCodeRequired = 8 'cancel
            Exit Sub
        End If
    End If
'===================================
    
    Select Case giClockShow
    
        Case 3
            frmTransmitter.Show
            Unload Me
            giActivity = 1
            frmTransmitter!cmdPrevious.Caption = "&Return to Previous Page F6"
            giClockShow = 6
            Exit Sub
        Case 4
            frmPlanner.Show
            Unload Me
            giClockShow = 6
        Case 7
            frmTransmitter.Show
            Unload Me
            giActivity = 1
            frmTransmitter!cmdRestoreDefaults.SetFocus 'Take Entries Snapshot
            frmTransmitter!cmdPrevious.Caption = "&Return to Previous Page F6"
            giClockShow = 6
            Exit Sub
        Case 0
            frmPlanner.Show
            Unload Me
        Case Else
            frmPlanner.Show
            Unload Me
        End Select
    '=================================
    If txtStation1 = "" Then
        mCount = mCount + 1
        If mCount < 2 Then
            MsgBox "Station Call Letters are required for the 'Flagship Station'.", 0, "Flagship Station Call Letters Missing"
            txtStation1.SetFocus
            Exit Sub
        ElseIf mCount >= 2 Then
            mCount = 0
            GoTo SkipTo 'name of GoTo line
            Exit Sub
        End If
    End If
    
    Frame1.Caption = "Set Transmitter Power Limits"
    Label14.Enabled = True
    Label15.Enabled = True
    Label16.Enabled = True
    Label17.Enabled = True
    Label19.Enabled = True
    
    Frame2.Enabled = True
    Label1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label18.Enabled = True
    
    mnuHelp.Enabled = True
    frmDefaults!mnuPage.Enabled = True
    frmDefaults!mnuFile.Enabled = True
    frmDefaults!mnuRestoreDefaults.Enabled = True
    
    Label5.Enabled = True
    Label6.Enabled = True
    Label7.Enabled = True
    Label8.Enabled = True
    Label10.Enabled = True
    Label11.Enabled = True
    Label12.Enabled = True
    Label13.Enabled = True
    
    '-------------reactivate from set power limits
     frmDefaults!txtStation1.Enabled = True
    frmDefaults!txtStation2.Enabled = True
    frmDefaults!txtStation3.Enabled = True
    frmDefaults!txtStation4.Enabled = True

    frmDefaults!txtPlanTime.Enabled = True
    frmDefaults!txtIntroOut.Enabled = True
    frmDefaults!txtClose.Enabled = True
    frmDefaults!txtSpot.Enabled = True
    frmDefaults!Frame2.Enabled = True

    frmDefaults!mnuHelp.Enabled = True
    frmDefaults!mnuPage.Enabled = True
    frmDefaults!mnuFile.Enabled = True
    frmDefaults!mnuRestoreDefaults.Enabled = True

    '----------reactivate from set call letters

    frmDefaults!txtStation1Min.Enabled = True
    frmDefaults!txtStation2Min.Enabled = True
    frmDefaults!txtStation3Min.Enabled = True
    frmDefaults!txtStation4Min.Enabled = True

    frmDefaults!txtStation1Max.Enabled = True
    frmDefaults!txtStation2Max.Enabled = True
    frmDefaults!txtStation3Max.Enabled = True
    frmDefaults!txtStation4Max.Enabled = True

    frmDefaults!txtPlanTime.Enabled = True
    frmDefaults!txtIntroOut.Enabled = True
    frmDefaults!txtClose.Enabled = True
    frmDefaults!txtSpot.Enabled = True
    frmDefaults!Frame2.Enabled = True

   ' frmDefaults!Shape1.Visible = True
    frmDefaults!Label14.Enabled = True
    frmDefaults!Label15.Enabled = True
    frmDefaults!Label16.Enabled = True
    frmDefaults!Label17.Enabled = True
    frmDefaults!Label10.Enabled = True
    frmDefaults!Label11.Enabled = True
    frmDefaults!Label12.Enabled = True
    frmDefaults!Label13.Enabled = True
    frmDefaults!mnuHelp.Enabled = True
    frmDefaults!mnuPage.Enabled = True
    frmDefaults!mnuFile.Enabled = True
    frmDefaults!mnuRestoreDefaults.Enabled = True

    frmDefaults!Frame2.Enabled = True

    '----------
     
SkipTo:

    If giClockShow = 3 Then
        frmTransmitter.Show
    ElseIf giClockShow = 5 Then
       frmTimeRemain.Show
    ElseIf giClockShow = 4 Then
        frmPlanner.Show
    End If
  
    cmdSetSignature.Enabled = True
    txtSignatureLine.Enabled = True
    cmdExit.Visible = False
    iCodeRequired = 0
    cmdSave.Caption = "&Save Changes and Exit Page"
    iCount2 = 0
    Unload frmMemos
    Unload Me
End Sub

Private Sub cmdDefaults_Click()

'================ DUALV 1 ====STATION Version. Comment out HOME version below.
'
    If txtStation1 = "" Then
        MsgBox "Flag Ship station is a required entry", vbOKOnly, "Missing Required Data"
        Exit Sub
    ElseIf txtPlanTime = "" Then
        MsgBox "Planned Time (average minutes of music per hour) is a required entry", vbOKOnly, "Missing Required Data"
        Exit Sub
    End If

    cmdDefaults.Caption = "Save Time && Transmitter Settings as &Defaults"
    '----
On Error GoTo HandleErrors
    Dim prompt, AccessCode
    prompt = "Access Code is Required to Set Current Data as Default Values." & vbCrLf & vbCrLf & "Enter Access Code."
    AccessCode = InputBox$(prompt, "Access Code Required")

    If AccessCode = giAccess Then

        giClockShow = 4
        imgSmile(1).Visible = True
        lblDefaults.Visible = True
        cmdDefaults.Caption = "&Defaults Are Set"
        cmdDefaults.Enabled = False
        cmdSetSignature.Enabled = False
        cmdCancel.Caption = "Close"
        cmdCancel.SetFocus
        cmdSave.Enabled = True

    ElseIf AccessCode <> giAccess And AccessCode <> "" And iCount <= 1 Then
        MsgBox AccessCode & " is incorrect Access Code", vbOKOnly, "Incorrect Code"
        cmdDefaults.Caption = "Try &Again"
        cmdDefaults.SetFocus
        iCount = iCount + 1
        Exit Sub
     ElseIf iCount = 2 Then
        MsgBox "Incorrect Access Codes. Exiting Set Default Procedure", vbOKOnly, "Incorrect Codes"
        cmdDefaults.Caption = "Save Time && Transmitter Settings as &Defaults"
        iCount = 0
        cmdCancel_Click
        Exit Sub
      Else
        Exit Sub
    End If
'-------
    Dim rDate As Date
    rDate = Now

    If txtPlanTime <> "" Then
        Open "DefaultTimes.dat" For Output As #22
            Write #22, txtIntroOut, txtClose, txtSpot
        Close #22
    End If

    If txtStation1 <> "" Then
        Open "DefaultStation.dat" For Output As #20
            Write #20, txtStation1, txtStation2, txtStation3, txtStation4
        Close #20
    End If

    If txtStation1Min <> "" Then
        Open "DefaultMinMax.dat" For Output As #24
            Write #24, txtStation1Min, txtStation1Max, txtStation2Min, txtStation2Max, _
                txtStation3Min, txtStation3Max, txtStation4Min, txtStation4Max
            Close #24
    End If

    Open "DefaultDate.dat" For Output As #17
        Write #17, rDate
    Close #17

     If txtMin <> "" And txtMax <> "" And (txtIdeal1 <> "" Or txtIdeal2 <> "" _
    Or txtIdeal3 <> "" Or txtIdeal4 <> "") Then

        Open "DefaultXmitterIdeal.dat" For Output As #26
                Write #26, txtMin, txtMax, txtIdeal1, txtIdeal2, txtIdeal3, txtIdeal4, _
                ckIdeal1.Value, ckIdeal2.Value, ckIdeal3.Value, ckIdeal4.Value,
        Close #26
    End If

'    If frmTransmitter!txtVmnr <> "" And frmTransmitter!txtAmnr <> "" Then
'        Open "DefaultReadings.dat" For Output As #16
'        Write #16, frmTransmitter!txtVmnr, frmTransmitter!txtVrxc, frmTransmitter!txtVgrs, frmTransmitter!txtVgsk, frmTransmitter!txtAmnr, frmTransmitter!txtArxc, _
'        frmTransmitter!txtAgrs, frmTransmitter!txtAgsk, frmTransmitter!txtEmnr, frmTransmitter!txtErxc, frmTransmitter!txtEgrs, frmTransmitter!txtEgsk, _
'        frmTransmitter!txtVolt1, frmTransmitter!txtVolt2, frmTransmitter!txtVolt3, frmTransmitter!txtVolt4, _
'        frmTransmitter!txtAmp1, frmTransmitter!txtAmp2, frmTransmitter!txtAmp3, frmTransmitter!txtAmp4, _
'        frmTransmitter!txtEmnr, frmTransmitter!txtErxc, frmTransmitter!txtEgrs, frmTransmitter!txtEgsk
'        Close #16
'    End If

    cmdCancel.SetFocus
    Exit Sub
HandleErrors:
    MsgBox "An error saving current data as default has occurred", vbOKOnly, "Data Not Saved as Default"
    Close #17
    Close #26
    Close #16
    Close #24
    Close #20
    Close #22

'=========== DUALV 2 ----HOME Version  Comment out entire STATION Version above

'    If cmdDefaults.Caption <> "&Defaults Are Set" Then
'    Dim iResponse As Integer
'    iResponse = MsgBox("Current Default settings will be replaced", vbYesNo, "Save Settings as Defaults")
'
'    If iResponse = vbNo Then
'        Exit Sub
'    Else
'        End If
'    End If
'
'    If cmdDefaults.Caption = "&Defaults Are Set" Then
'        imgSmile(1).Visible = False
'        lblDefaults.Visible = False
'        cmdDefaults.Caption = "Save Time && Transmitter Settings as &Defaults"
'        cmdCancel.Caption = "Close"
'        cmdSetSignature.Enabled = True
'        txtSignatureLine.Enabled = True
'        cmdCancel.SetFocus
'        Exit Sub
'    End If
'
'        If txtStation1 = "" Then
'            MsgBox "Flag Ship station is a required entry", vbOKOnly, "Missing Required Data"
'            Exit Sub
'        ElseIf txtPlanTime = "" Then
'            MsgBox "Planned Time (average minutes of music per hour) is a required entry", vbOKOnly, "Missing Required Data"
'            Exit Sub
'     End If
'
'    cmdDefaults.Caption = "Save Time && Transmitter Settings as &Defaults"
'    '----
'On Error GoTo HandleErrors
'
'    giClockShow = 4
'    imgSmile(1).Visible = True
'    lblDefaults.Visible = True
'    cmdDefaults.Caption = "&Defaults Are Set"
'    cmdSetSignature.Enabled = False
'    txtSignatureLine.Enabled = False
'    cmdExit.Enabled = True
''-------
'    Dim rDate As Date
'    rDate = Now
'
'    If txtPlanTime <> "" Then
'        Open "DefaultTimes.dat" For Output As #22
'            Write #22, txtIntroOut, txtClose, txtSpot
'        Close #22
'    End If
'
'    If txtStation1 <> "" Then
'        Open "DefaultStation.dat" For Output As #20
'            Write #20, txtStation1, txtStation2, txtStation3, txtStation4
'        Close #20
'    End If
'
'    If txtStation1Min <> "" Then
'        Open "DefaultMinMax.dat" For Output As #24
'            Write #24, txtStation1Min, txtStation1Max, txtStation2Min, txtStation2Max, _
'                txtStation3Min, txtStation3Max, txtStation4Min, txtStation4Max
'            Close #24
'    End If
'
'    Open "DefaultDate.dat" For Output As #17
'        Write #17, rDate
'    Close #17
'
'     If txtMin <> "" And txtMax <> "" And (txtIdeal1 <> "" Or txtIdeal2 <> "" _
'    Or txtIdeal3 <> "" Or txtIdeal4 <> "") Then
'
'        Open "DefaultXmitterIdeal.dat" For Output As #26
'                Write #26, txtMin, txtMax, txtIdeal1, txtIdeal2, txtIdeal3, txtIdeal4, _
'                ckIdeal1.Value, ckIdeal2.Value, ckIdeal3.Value, ckIdeal4.Value,
'        Close #26
'    End If
'
'    If frmTransmitter!txtVmnr <> "" And frmTransmitter!txtAmnr <> "" Then
'        Open "DefaultReadings.dat" For Output As #16
'        Write #16, frmTransmitter!txtVmnr, frmTransmitter!txtVrxc, frmTransmitter!txtVgrs, frmTransmitter!txtVgsk, frmTransmitter!txtAmnr, frmTransmitter!txtArxc, _
'        frmTransmitter!txtAgrs, frmTransmitter!txtAgsk, frmTransmitter!txtEmnr, frmTransmitter!txtErxc, frmTransmitter!txtEgrs, frmTransmitter!txtEgsk, _
'        frmTransmitter!txtVolt1, frmTransmitter!txtVolt2, frmTransmitter!txtVolt3, frmTransmitter!txtVolt4, _
'        frmTransmitter!txtAmp1, frmTransmitter!txtAmp2, frmTransmitter!txtAmp3, frmTransmitter!txtAmp4, _
'        frmTransmitter!txtEmnr, frmTransmitter!txtErxc, frmTransmitter!txtEgrs, frmTransmitter!txtEgsk
'        Close #16
'    End If
'    cmdDefaults.SetFocus
'    Exit Sub
'HandleErrors:
'    MsgBox "An error saving current data as default has occurred", vbOKOnly, "Data Not Saved as Default"
'    Close #17
'    Close #26
'    Close #16
'    Close #24
'    Close #20
'    Close #22
'''-----------end Home version
End Sub

Private Sub cmdExit_Click()
   
    giExit = 1
    Unload frmAbout
    Unload frmAddHelp
    Unload frmAddTime
    Unload frmEditHelp
    Unload frmF4Help
    Unload frmLogActivity
    Unload frmMemos
    Unload frmNote
    Unload frmPlanHelp
    Unload frmPrintHelp
    Unload frmSplash
    Unload frmStaff
    Unload frmStopWatch
    Unload frmTimeRemain
    Unload frmTransmitterHints
    Unload frmTransmitter
    Unload Me
    Unload frmPlanner
End Sub

Private Sub cmdSave_Click()

    If txtPlanTime < 54 Or txtPlanTime > 58 Then
        MsgBox "The average minutes of music planned for the hour temporarily can be any amount you select." & vbCrLf & vbCrLf & _
        "However, to save the setting it must be between 54 and 58 minutes of music in the hour", vbOKOnly, _
        "To Save, Music Time Must be Between 54 and 58 Minutes"
        Exit Sub
    End If

    iChange = 0
    Unload frmMemos
On Error GoTo HandleErrors
    If txtStation1 = "" Then
        mCount = mCount + 1
        If mCount < 3 Then
            MsgBox "Station Call Letters are required for the 'Flagship Station'.", 0, "Flagship Station Call Letters Missing"
            txtStation1.SetFocus
            Exit Sub
        ElseIf mCount >= 3 Then
            giClockShow = 4
            mCount = 0
            Unload Me
            Unload frmTransmitter
            frmPlanner.Show
        Exit Sub
        End If
    End If
    '-------
 
    If txtStation2 <> "" Then
        frmTransmitter!txtVrxc.Enabled = True
        frmTransmitter!txtArxc.Enabled = True
        frmTransmitter!lblStation2.Enabled = True
        frmTransmitter!lblrxc.Visible = True
        frmTransmitter!lblPwrLimit2.Visible = True
        frmTransmitter!txtVrxc.BackColor = &H80000005 'white
        frmTransmitter!txtArxc.BackColor = &H80000005
    Else
        frmTransmitter!txtVrxc = ""
        frmTransmitter!txtArxc = ""
        frmTransmitter!txtErxc = ""
        frmTransmitter!lblStation2.Enabled = False
        
        frmTransmitter!lblrxc.Visible = False
        frmTransmitter!lblPwrLimit2.Visible = False
        frmDefaults!txtStation2Min = ""
        frmDefaults!txtStation2Max = ""
      
        frmTransmitter!txtVrxc.Enabled = False
        frmTransmitter!txtArxc.Enabled = False
        frmTransmitter!txtVrxc.BackColor = &HE0E0E0 'gray
        frmTransmitter!txtArxc.BackColor = &HE0E0E0
    End If
    '--------
    If txtStation3 <> "" Then
        frmTransmitter!txtVgrs.Enabled = True
        frmTransmitter!txtAgrs.Enabled = True
        frmTransmitter!lblStation3.Enabled = True
        frmTransmitter!lblgrs.Visible = True
        frmTransmitter!lblPwrLimit3.Visible = True
        frmTransmitter!txtVgrs.BackColor = &H80000005 'white
        frmTransmitter!txtAgrs.BackColor = &H80000005
    Else
        frmTransmitter!txtVgrs = ""
        frmTransmitter!txtAgrs = ""
        frmTransmitter!txtEgrs = ""
        frmTransmitter!lblgrs.Visible = False
        frmTransmitter!lblStation3.Enabled = False
        
        frmTransmitter!lblPwrLimit3.Visible = False
        frmDefaults!txtStation3Min = ""
        frmDefaults!txtStation3Max = ""
        
        frmTransmitter!txtVgrs.Enabled = False
        frmTransmitter!txtAgrs.Enabled = False
        frmTransmitter!txtVgrs.BackColor = &HE0E0E0 'gray
        frmTransmitter!txtAgrs.BackColor = &HE0E0E0
    End If
    '-------
    If txtStation4 <> "" Then
        frmTransmitter!txtVgsk.Enabled = True
        frmTransmitter!txtAgsk.Enabled = True
        frmTransmitter!lblStation4.Enabled = True
        frmTransmitter!lblgsk.Visible = True
        frmTransmitter!lblPwrLimit4.Visible = True
        frmTransmitter!txtVgsk.BackColor = &H80000005 'white
        frmTransmitter!txtAgsk.BackColor = &H80000005
    Else
        frmTransmitter!txtVgsk = ""
        frmTransmitter!txtAgsk = ""
        frmTransmitter!txtEgsk = ""
        frmTransmitter!lblStation4.Enabled = False
        
        frmTransmitter!lblgsk.Visible = False
        frmTransmitter!lblPwrLimit4.Visible = False
        frmDefaults!txtStation4Min = ""
        frmDefaults!txtStation4Max = ""
     
        frmTransmitter!txtVgsk.Enabled = False
        frmTransmitter!txtAgsk.Enabled = False
        frmTransmitter!txtVgsk.BackColor = &HE0E0E0 'gray
        frmTransmitter!txtAgsk.BackColor = &HE0E0E0
    End If
    '----------
'=========================

    If giCodeRequired = 0 Then 'Normal; giCodeRequired will = 1 only after initial setup
        
        gStation1Min = txtStation1Min
        gStation1Max = txtStation1Max
        gStation2Min = txtStation2Min
        gStation2Max = txtStation2Max
        gStation3Min = txtStation3Min
        gStation3Max = txtStation3Max
        gStation4Min = txtStation4Min
        gStation4Max = txtStation4Max
                
        Open "MinMax.dat" For Output As #19
        Write #19, txtStation1Min, txtStation1Max, txtStation2Min, txtStation2Max, _
                txtStation3Min, txtStation3Max, txtStation4Min, txtStation4Max
        Close #19
        
        If txtStation1 <> "" Then
           giStation1 = txtStation1 'frmOption station1 is loaded for distribution
        End If
        frmTransmitter!lblStation1 = giStation1 'and as frmTransmitter flagship station1.
        frmTransmitter!lblStation2 = txtStation2 'defaults loaded into remaining frmTransmitter stations
        frmTransmitter!lblStation3 = txtStation3
        frmTransmitter!lblStation4 = txtStation4
        
        cmdCancel.Enabled = True
        
        If txtStation1 <> "" Then
            cmdCancel.SetFocus
        Else
            Beep
            txtStation1.SetFocus
        End If
        
        If txtMin <> "" And txtMax <> "" And (txtIdeal1 <> "" Or txtIdeal2 <> "" _
        Or txtIdeal3 <> "" Or txtIdeal4 <> "") Then
            Open "XmitterIdeal.dat" For Output As #25
                    Write #25, txtMin, txtMax, txtIdeal1, txtIdeal2, txtIdeal3, txtIdeal4, _
                    ckIdeal1.Value, ckIdeal2.Value, ckIdeal3.Value, ckIdeal4.Value,
            Close #25
        End If
     
       '-----------
       
        If txtPlanTime = "" Then
            txtPlanTime = "56"
        End If
        
        If txtIntroOut = "" Then
            txtIntroOut = "50"
        End If
        
        If txtSpot = "" Then
            txtSpot = "30"
        End If
        
        If txtClose = "" Then
            txtClose = "30"
        End If
       
        If iCodeRequired = 0 Then
            MsgBox "Changes have been saved", vbOKOnly, "Changes Saved"
            cmdCancel_Click
            Exit Sub
            
        ElseIf iCount2 = 0 Then
            
                Dim iResponse As Integer
                
                Select Case iCodeRequired
                Case 1
                    iResponse = MsgBox("Access Code required to save Planned Time change", vbOKCancel, "Planned Time Change, Access Code Required")
                Case 2
                    iResponse = MsgBox("Access Code required to save Intro-Back Announce time change", vbOKCancel, "Intro-Back Announce Time Change, Access Code Required")
                Case 3
                    iResponse = MsgBox("Access Code required to save Spot Average Time change", vbOKCancel, "Spot Time Change, Access Code Required")
                Case 4
                    iResponse = MsgBox("Access Code required to save Close Out-ID Time change", vbOKCancel, "Close Out-ID Time Change, Access Code Required")
                Case 5
                    iResponse = MsgBox("Access Code Required to Change Station Call Letters", vbOKCancel, "Access Code Required")
                End Select
                    
            If iResponse = vbCancel Then
                iCodeRequired = 8 'cancel
                Exit Sub
            End If
        End If
    '-----
    On Error GoTo HandleErrors
        Dim prompt, AccessCode
        prompt = "Access Code is Required to Set Current Data as Default Values." & vbCrLf & vbCrLf & "Enter Access Code."
        AccessCode = InputBox$(prompt, "Access Code Required")
    
        If AccessCode = "" Then
            cmdCancel_Click
            Exit Sub
        
        ElseIf AccessCode = giAccess Then
           
        ElseIf AccessCode <> giAccess And AccessCode <> "" Then
        
            If iCount2 = 0 Then
                MsgBox AccessCode & " is an incorrect access code" & vbCrLf & vbCrLf & "Try again", vbOKOnly, "Incorrect Code"
                cmdSave.Caption = "Try &Again"
                cmdSave.SetFocus
                iCount2 = iCount2 + 1
                Exit Sub
                
            ElseIf iCount2 = 1 Then
                MsgBox AccessCode & " is an incorrect access code" & vbCrLf & vbCrLf & "Try one more time.", vbOKOnly, "Incorrect Code"
                cmdSave.Caption = "Try &Again"
                cmdSave.SetFocus
                iCount2 = iCount2 + 1
                Exit Sub
                
             ElseIf iCount2 = 2 Then
                MsgBox "Incorrect Access Code." & vbCrLf & vbCrLf & "Exiting Access Code Procedure", vbOKOnly, "Incorrect Code"
                cmdSave.Caption = "&Save Changes and Exit Page"
                iCount2 = 0
                cmdCancel_Click
                Exit Sub
            Else
                iCount2 = 0
                iCodeRequired = 8 'cancel
                Exit Sub
            End If
    
        End If
    '------
        If iCodeRequired = 1 Or iCodeRequired = 2 Or iCodeRequired = 3 Or iCodeRequired = 4 Then
            Open "Times.dat" For Output As #23
            Write #23, txtPlanTime, txtIntroOut, txtClose, txtSpot
            Close #23
           
            '-------
            If txtIntroOut <> "" Then
                frmTimeRemain!txtIntroSetting = txtIntroOut
            End If
        
             If txtClose <> "" Then
                frmTimeRemain!txtCloseOut = txtClose
            End If
        
            If txtSpot <> "" Then
                frmTimeRemain!txtSpotLength = txtSpot
            End If
        ElseIf iCodeRequired = 5 Then
            Open "Stations.dat" For Output As #18
            Write #18, txtStation1, txtStation2, txtStation3, txtStation4
            Close #18
        End If
        If iCodeRequired <> 8 Then
            MsgBox "Changes have been saved", vbOKOnly, "Changes Saved"
        End If
        cmdCancel_Click 'save function completed. passes action to cmdCancel to close & unload form
        Exit Sub
        
    '=============

    ElseIf giCodeRequired = 1 Then 'Not Normal; giCodeRequired will = 1 only after initial setup

        If txtPlanTime <> "" And IsNumeric(txtPlanTime) Then
            giPlannedTime = txtPlanTime
        Else
            txtPlanTime = "56.0"
            giPlannedTime = "56.0"
        End If
        
        frmPlanner!txtBlock = txtPlanTime  'giPlannedTime 'Planner time block loads default
    
        If txtPlanTime > "56.0" Or txtPlanTime < "1" Then
            txtPlanTime = ""
        End If
        
        gStation1Min = txtStation1Min
        gStation1Max = txtStation1Max
        gStation2Min = txtStation2Min
        gStation2Max = txtStation2Max
        gStation3Min = txtStation3Min
        gStation3Max = txtStation3Max
        gStation4Min = txtStation4Min
        gStation4Max = txtStation4Max
        
        Open "Times.dat" For Output As #23
        Write #23, txtPlanTime, txtIntroOut, txtClose, txtSpot
        Close #23
        
        Open "Stations.dat" For Output As #18
        Write #18, txtStation1, txtStation2, txtStation3, txtStation4
        Close #18
                
        Open "MinMax.dat" For Output As #19
        Write #19, txtStation1Min, txtStation1Max, txtStation2Min, txtStation2Max, _
                txtStation3Min, txtStation3Max, txtStation4Min, txtStation4Max
        Close #19
        
        If txtStation1 <> "" Then
           giStation1 = txtStation1 'frmOption station1 is loaded for distribution
        End If
        frmTransmitter!lblStation1 = giStation1 'and as frmTransmitter flagship station1.
        frmTransmitter!lblStation2 = txtStation2 'defaults loaded into remaining frmTransmitter stations
        frmTransmitter!lblStation3 = txtStation3
        frmTransmitter!lblStation4 = txtStation4
           
        If txtIntroOut <> "" Then
            frmTimeRemain!txtIntroSetting = txtIntroOut
        End If
        
         If txtClose <> "" Then
            frmTimeRemain!txtCloseOut = txtClose
        End If
        
        If txtSpot <> "" Then
            frmTimeRemain!txtSpotLength = txtSpot
        End If
        
        cmdCancel.Enabled = True
        
        If txtStation1 <> "" Then
            cmdCancel.SetFocus
        Else
            Beep
            txtStation1.SetFocus
        End If
        
        If txtMin <> "" And txtMax <> "" And (txtIdeal1 <> "" Or txtIdeal2 <> "" _
        Or txtIdeal3 <> "" Or txtIdeal4 <> "") Then
            Open "XmitterIdeal.dat" For Output As #25
                    Write #25, txtMin, txtMax, txtIdeal1, txtIdeal2, txtIdeal3, txtIdeal4, _
                    ckIdeal1.Value, ckIdeal2.Value, ckIdeal3.Value, ckIdeal4.Value,
            Close #25
        End If
        
        iChange = 0
        cmdCancel_Click 'save function completed. passes action to cmdCancel to close & unload form
        
        Exit Sub
    End If
'===================
 
HandleErrors:
    MsgBox "Error, some or all data not saved.", vbOKOnly, "Error"
    Close #25
    Close #23
    Close #18
    Close #19

End Sub

Private Sub cmdSetSignature_Click()

'================ DUALV 1 ====STATION VERSION. Comment out HOME VERSION below

''''access code procedure

    Dim prompt, AccessCode
On Error GoTo HandleErrors

    If iAccessOpen = 0 Then

        prompt = "Access Code is Required to Set or Change MusicLog Printout Signature Line." & vbCrLf & vbCrLf & "Enter Access Code."
        AccessCode = InputBox$(prompt, "Access Code Required")

        If AccessCode = giAccess Then
            iAccessOpen = 1
            Frame5.Caption = "Set Music Logbook Printout Signature Line"

        ElseIf AccessCode <> giAccess And AccessCode <> "" Then
            MsgBox "Incorrect Access Code", vbOKOnly, "Incorrect Code"
            cmdCancel.Caption = "Close"
            Exit Sub
        ElseIf AccessCode = "" Then

            Dim Signature As String

            Open "Signature.dat" For Input As #27
            Input #27, Signature
            Close #27

            cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
            txtSignatureLine.Text = Signature

            cmdCancel.Caption = "Close"
            cmdCancel.SetFocus
            Exit Sub
        End If
    End If
'''''------------------end of access code procedure, beginning of save signature procedure

    imgSmile(0).Visible = False
    Label25.Visible = False

    If cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature" Then

            cmdSetSignature.Caption = "Music Logbook Page Default Signature"
            txtSignatureLine.SetFocus
            giClockShow = 4
            cmdDefaults.Enabled = False

    ElseIf cmdSetSignature.Caption = "Music Logbook Page Printout Default Signature" Then

        MsgBox "No change was made to the signature line", vbOKOnly, "No Change"
        cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
         cmdCancel.SetFocus
        Exit Sub

        ''---------
    ElseIf cmdSetSignature.Caption = "Click to Save or Cancel Change" Then

        Dim iResponse As Integer

        iResponse = MsgBox("Do you want to save the Signature Line change?", vbYesNo, "Save Change")
        If iResponse = vbNo Then

            Open "Signature.dat" For Input As #27
            Input #27, Signature
            Close #27

            txtSignatureLine.Text = Signature
            cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
            cmdCancel.Caption = "Close"
            cmdCancel.SetFocus
            cmdDefaults.Enabled = True
            cmdSave.Enabled = True
            imgSmile(0).Visible = False
            Label25.Visible = False
            imgHand.Visible = False
            Exit Sub

        ElseIf iResponse = vbYes Then

            Open "Signature.dat" For Output As #27
            Write #27, txtSignatureLine
            Close #27

            cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
            cmdCancel.Caption = "Close"
            cmdCancel.SetFocus
            cmdDefaults.Enabled = True
            cmdSave.Enabled = True
            imgSmile(0).Visible = True
            imgSmile(0).ToolTipText = ""
            Label25.Caption = "Saved"
            Label25.ToolTipText = ""
            Label25.Visible = True
            imgHand.Visible = False
        End If
    End If
HandleErrors:
    Close #27
''===================End STATION Version

'================ DUALV 2 ====HOME VERSION. Comment out STATION VERSION above

'imgSmile(0).Visible = False
'Label25.Visible = False
'
'On Error GoTo HandleErrors
'  If cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature" Then
'
'            cmdSetSignature.Caption = "Music Logbook Page Default Signature"
'            txtSignatureLine.SetFocus
'            giClockShow = 4
'            cmdDefaults.Enabled = False
'
'    ElseIf cmdSetSignature.Caption = "Music Logbook Page Default Signature" Then
'
'        MsgBox "No change was made to the signature line", vbOKOnly, "No Change"
'        cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
'         cmdCancel.SetFocus
'        Exit Sub
'
'        '---------
'    ElseIf cmdSetSignature.Caption = "Click to Save or Cancel Change" Then
'        imgHand.Visible = True
'        Dim iResponse As Integer
'        Dim Signature As String
'
'        iResponse = MsgBox("Do you want to save the Signature Line change?", vbYesNo, "Save Change")
'        If iResponse = vbNo Then
'            Open "Signature.dat" For Input As #27
'            Input #27, Signature
'            Close #27
'            txtSignatureLine.Text = Signature
'            cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
'            cmdCancel.Caption = "Close"
'            cmdCancel.SetFocus
'            cmdDefaults.Enabled = True
'            cmdSave.Enabled = True
'            imgSmile(0).Visible = False
'            Label25.Visible = False
'            imgHand.Visible = False
'            Exit Sub
'
'        ElseIf iResponse = vbYes Then
'
'            Open "Signature.dat" For Output As #27
'            Write #27, txtSignatureLine
'            Close #27
'
'            cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
'            cmdCancel.Caption = "Close"
'            cmdCancel.SetFocus
'            cmdDefaults.Enabled = True
'            cmdSave.Enabled = True
'            imgSmile(0).Visible = True
'            Label25.Caption = "Saved"
'            Label25.Visible = True
'            imgHand.Visible = False
'        End If
'    End If
'HandleErrors:
'    Close #27
''===============End HOME Version
End Sub

Private Sub Form_Activate()
    iChange = 0
    
    If iAccessOpen = 1 Then
        Frame5.Caption = "Set Music Logbook Printout Signature Line"
    End If
   ' cmdCancel.SetFocus
End Sub

Private Sub Form_DblClick()
    lblHelp.Visible = False
    Shape2(2).Visible = False
    Label9.Visible = True
    mnuHelpTimes.Checked = False
    cmdCancel.SetFocus
End Sub

Private Sub Form_Load()
   
On Error GoTo HandleErrors

    Dim Station1, Station2, Station3, Station4 As String
    
    Dim Station1Min, Station1Max, Station2Min, Station2Max, Station3Min, _
        Station3Max, Station4Min, Station4Max As String

    Open "Stations.dat" For Input As #18
    Input #18, Station1, Station2, Station3, Station4
    Close #18
            
    Open "MinMax.dat" For Input As #19
    Input #19, Station1Min, Station1Max, Station2Min, Station2Max, _
            Station3Min, Station3Max, Station4Min, Station4Max
    Close #19

    txtStation1 = Station1
    txtStation2 = Station2
    txtStation3 = Station3
    txtStation4 = Station4
    
    txtStation1Min = Station1Min
    txtStation1Max = Station1Max
    txtStation2Min = Station2Min
    txtStation2Max = Station2Max
    txtStation3Min = Station3Min
    txtStation3Max = Station3Max
    txtStation4Min = Station4Min
    txtStation4Max = Station4Max
    
    gStation1Min = Station1Min 'loads station min & max values globally
    gStation1Max = Station1Max
    gStation2Min = Station2Min
    gStation2Max = Station2Max
    gStation3Min = Station3Min
    gStation3Max = Station3Max
    gStation4Min = Station4Min
    gStation4Max = Station4Max
    
    If txtStation1 = "" Then
        cmdExit.Visible = True
    Else
        cmdExit.Visible = False
    End If
    
    Dim PlanTime, IntroOut, sClose, Spot As String
    Dim sSignatureline As String
    
    Open "Times.dat" For Input As #23
    Input #23, PlanTime, IntroOut, sClose, Spot
    Close #23
    
    Open "Signature.dat" For Input As #27
    Input #27, sSignatureline
    Close #27
    
    If IsNumeric(PlanTime) Then
        txtPlanTime = Format$(PlanTime, "00.0")
    End If
    
    If IsNumeric(IntroOut) Then
        txtIntroOut = IntroOut
    End If
    
     If IsNumeric(sClose) Then
        txtClose = sClose
     End If
     
    If IsNumeric(Spot) Then
       txtSpot = Spot
    End If
          
    txtSignatureLine = sSignatureline
    frmPlanner!txtSignature = txtSignatureLine
    
    Dim iMin, iMax, iIdeal1, iIdeal2, iIdeal3, iIdeal4, iCk1, iCk2, iCk3, iCk4, iHourNow As Integer
    
    Open "XmitterIdeal.dat" For Input As #25
        Input #25, iMin, iMax, iIdeal1, iIdeal2, iIdeal3, iIdeal4, iCk1, iCk2, iCk3, iCk4
    Close #25
    
    txtMin = iMin
    txtMax = iMax
    txtIdeal1 = iIdeal1
    txtIdeal2 = iIdeal2
    txtIdeal3 = iIdeal3
    txtIdeal4 = iIdeal4
    ckIdeal1.Value = iCk1
    ckIdeal2.Value = iCk2
    ckIdeal3.Value = iCk3
    ckIdeal4.Value = iCk4
    
    If txtMin = "" Or txtMin = "0" Then
        txtMin = "90"
    End If
    
    If txtMax = "" Or txtMax = "0" Then
        txtMax = "105"
    End If

    Exit Sub
HandleErrors:

    If txtStation1 = "" Then
        cmdExit.Visible = True
    End If
    
    If txtPlanTime = "" Then
        txtPlanTime = "56.0"
    End If
    
    If txtIntroOut = "" Then
        txtIntroOut = "50"
    End If
    
    If txtClose = "" Then
        txtClose = "30"
    End If
    
    If txtSpot = "" Then
        txtSpot = "30"
    End If
    
    txtMin = "90"
    txtMax = "105"
'----------
Open "DefaultXmitterIdeal.dat" For Output As #26
Write #26, 90, 105, 0, 0, 0, 0, 0, 0, 0, 0
Close #26

    '-------
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
  '-----------
    Close #25
    Close #23
    Close #18
    Close #19
End Sub

Private Sub Frame1_DblClick()
    lblHelp.Visible = False
    Shape2(2).Visible = False
    Label9.Visible = True
    mnuHelpTimes.Checked = False
    cmdCancel.SetFocus
End Sub

Private Sub Frame2_DblClick()
    lblHelp.Visible = False
    Shape2(2).Visible = False
    Label9.Visible = True
    mnuHelpTimes.Checked = False
    cmdCancel.SetFocus
End Sub

Private Sub Frame3_DblClick()
    lblHelp.Visible = False
    Shape2(2).Visible = False
    Label9.Visible = True
    mnuHelpTimes.Checked = False
    cmdCancel.SetFocus
End Sub

Private Sub imgSmile_Click(Index As Integer)

    If Label25.Caption = "Saved" Then
        imgSmile(0).ToolTipText = ""
        imgSmile(0).Visible = False
        Label25.Visible = False
        Exit Sub
    End If

On Error GoTo HandleErrors
    Dim Signature As String

    Open "Signature.dat" For Input As #27
    Input #27, Signature
    Close #27
    
    txtSignatureLine.Text = Signature
    imgSmile(0).Visible = False
    Label25.Visible = False
    cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
    Exit Sub
    
HandleErrors:
End Sub

Private Sub Label25_Click()

    If Label25.Caption = "Saved" Then
        Label25.Visible = False
        imgSmile(0).Visible = False
        Exit Sub
    End If
   
On Error GoTo HandleErrors

    Dim Signature As String
    
    Open "Signature.dat" For Input As #27
    Input #27, Signature
    Close #27
    
    txtSignatureLine.Text = Signature
    imgSmile(0).Visible = False
    Label25.Visible = False
    cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
    Exit Sub
HandleErrors:
End Sub

Private Sub lblHelp_DblClick()
    lblHelp.Visible = False
    Shape2(2).Visible = False
    Label9.Visible = True
    mnuHelpTimes.Checked = False
    Shape2(2).Visible = False
End Sub

Private Sub mnuFileTimeDefaults_Click()
    
On Error GoTo HandleErrors

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

    txtIntroOut = sIntroOut
    txtClose = sClose
    txtSpot = sSpot
    
    MsgBox "Default average Announce, Spot and Closeout times have been restored." & vbCrLf & "Click 'Save Changes and Exit Page' button to save these changes.", _
    vbOKOnly, "Save Changes"
    
    cmdSave.Enabled = True
    cmdSave.SetFocus
    Exit Sub
HandleErrors:

    MsgBox "Default times for Intro/Back-Annc, Spot & Closeout have not been saved", vbOKOnly + vbExclamation, "Missing Data"
    Close #22
    
End Sub

Private Sub mnuFilePrevious_Click()
    
    If iChange <> 0 Then
        Dim iResponse As Integer
        iResponse = MsgBox("Do you want to save any changes that may have been made?", vbYesNoCancel, "Save Changes?")
        
        If iResponse = vbCancel Then
            Exit Sub
        ElseIf iResponse = vbYes Then
            cmdSave_Click
        ElseIf iResponse = vbNo Then
        End If
        
        iChange = 0
    End If

    If giClockShow <> 0 Then
        Select Case giClockShow
            Case 4
                frmPlanner.Show
            Case 3
                frmTransmitter.Show
                frmTransmitter!cmdPrevious.Caption = "&Return to Previous Page F6"
            Case 5
                frmTimeRemain.Show
            Case Else
                frmPlanner.Show
        End Select
    Else
        frmPlanner.Show
    End If
    giClockShow = 6
    frmDefaults.Hide
End Sub

Private Sub mnuFileXmitterDefaults_Click()

    If giCodeRequired = 1 Then
        txtStation1 = "WMNR"
        txtStation2 = "WRXC"
        txtStation3 = "WGRS"
        txtStation4 = "WGSK"
    End If

On Error GoTo HandleErrors

    Dim Station1, Station2, Station3, Station4 As String

    Open "DefaultStation.dat" For Input As #20 'default data
    Input #20, Station1, Station2, Station3, Station4
    Close #20

    frmDefaults!txtStation1 = Station1 'UCase(Station1)
    frmDefaults!txtStation2 = Station2
    frmDefaults!txtStation3 = Station3
    frmDefaults!txtStation4 = Station4
'-----------

    Dim Station1Min, Station1Max, Station2Min, Station2Max, Station3Min, _
    Station3Max, Station4Min, Station4Max As String

    Open "DefaultMinMax.dat" For Input As #24
    Input #24, Station1Min, Station1Max, Station2Min, Station2Max, _
    Station3Min, Station3Max, Station4Min, Station4Max
    Close #24

    frmDefaults!txtStation1Min = Station1Min
    frmDefaults!txtStation1Max = Station1Max
    frmDefaults!txtStation2Min = Station2Min
    frmDefaults!txtStation2Max = Station2Max
    frmDefaults!txtStation3Min = Station3Min
    frmDefaults!txtStation3Max = Station3Max
    frmDefaults!txtStation4Min = Station4Min
    frmDefaults!txtStation4Max = Station4Max

'-------------
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
    
    MsgBox "Default Station Call Letters and Power Limits have been restored.", vbOKOnly, "Default Values"
    
    Exit Sub
HandleErrors:

    MsgBox "Station call letters and power limits default data has not been saved", vbOKOnly + vbExclamation, "Data Missing"
    Close #20
    Close #24
    Close #26
End Sub

Private Sub mnuHelpPrint_Click()

On Error GoTo HandleErrors
    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "Defaults")
    If iResponse = vbNo Then
        cmdCancel.SetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then
        PrintForm
    End If
    cmdCancel.SetFocus
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

Private Sub mnuHelpSave_Click()

On Error GoTo HandleErrors

    If txtIntroOut <> "" Or txtSpot <> "" Or txtClose <> "" Then
       Open "Times.dat" For Output As #23
       Write #23, txtPlanTime, txtIntroOut, txtClose, txtSpot
       Close #23
    End If
    
    If txtStation1 <> "" Then
        Open "Stations.dat" For Output As #18
        Write #18, txtStation1, txtStation2, txtStation3, txtStation4
        Close #18
    End If
            
    If txtStation1Min <> "" And txtStation1Max <> "" Then
        Open "MinMax.dat" For Output As #19
        Write #19, txtStation1Min, txtStation1Max, txtStation2Min, txtStation2Max, _
                txtStation3Min, txtStation3Max, txtStation4Min, txtStation4Max
        Close #19
    End If
    
    If txtMin <> "" And txtMax <> "" And (txtIdeal1 <> "" Or txtIdeal2 <> "" _
    Or txtIdeal3 <> "" Or txtIdeal4 <> "") Then
        Open "XmitterIdeal.dat" For Output As #25
        Write #25, txtMin, txtMax, txtIdeal1, txtIdeal2, txtIdeal3, txtIdeal4, _
        ckIdeal1.Value, ckIdeal2.Value, ckIdeal3.Value, ckIdeal4.Value,
        Close #25
    End If
    
        Exit Sub
HandleErrors:
    MsgBox "Error saving.", vbOKOnly, "Error"
    Close #25
    Close #23
    Close #18
    Close #19
End Sub

Private Sub mnuHelpTimes_Click()
    If mnuHelpTimes.Checked = True Then
        mnuHelpTimes.Checked = False
        lblHelp.Visible = False
        Shape2(2).Visible = False
        Label9.Visible = True
    Else
        mnuHelpTimes.Checked = True
        Label9.Visible = False
        lblHelp.Visible = True
        Shape2(2).Visible = True
    End If
End Sub

Private Sub mnuPageMusicLog_Click()
    giClockShow = 6
    frmPlanner.Show
    frmDefaults.Hide
End Sub

Private Sub mnuPageTransmitter_Click()
    giClockShow = 6
    frmTransmitter.Show
    frmDefaults.Hide
End Sub


Private Sub txtClose_Change()
    If Not IsNumeric(txtClose) And txtClose <> "" Then
        MsgBox "You have entered :     " & txtClose & "      This is an incorrect entry. Entry should be the estimated" & vbCrLf & _
        "average NUMBER of seconds required for a show's closeout and station ID.", 0, "Entry Error"
        
        txtClose = ""
        txtClose.SetFocus
    End If
End Sub

Private Sub txtClose_GotFocus()
    txtClose.SelStart = 0
    txtClose.SelLength = Len(txtClose)
    cmdSave.Enabled = True
End Sub

Private Sub txtClose_LostFocus()
    
    If Not IsNumeric(txtClose) Or txtClose = "" Then
        txtClose = "30"
    Else
        txtClose.Text = Format$(txtClose, "###")
    End If
    iCodeRequired = 4
End Sub

Private Sub txtIdeal1_Change()

    If txtIdeal1 <> "" And txtMin <> "" And txtMax <> "" Then
        ckIdeal1.Enabled = True
    ElseIf txtIdeal1 = "" Then
        ckIdeal1.Enabled = False
    End If

    If ckIdeal1.Value = 1 And txtIdeal1 <> "" And txtMin <> "" And txtMax <> "" Then
        iStation1min = Int(txtIdeal1) * Int(txtMin) / 100
        txtStation1Min = iStation1min
        iStation1max = Int(txtIdeal1) * Int(txtMax) / 100
        txtStation1Max = iStation1max
    
    ElseIf ckIdeal1.Value = 1 Then
        MsgBox "Minimum & Maximum percentage power limits and Ideal Power (Watts) text entry boxes must have data", _
        vbOKOnly, "Missing Data"
        ckIdeal1.Value = 0
        ckIdeal1.Enabled = False
    End If
    lblIdeal1 = txtStation1

End Sub

Private Sub txtIdeal1_GotFocus()
    txtIdeal1.SelStart = 0
    txtIdeal1.SelLength = Len(txtIdeal1)
    cmdSave.Enabled = True
End Sub

Private Sub txtIdeal1_LostFocus()
    If iChange = 0 Then
        iChange = 1
    End If
End Sub

Private Sub txtIdeal2_Change()

    If txtIdeal2 <> "" And txtMin <> "" And txtMax <> "" Then
        ckIdeal2.Enabled = True
    ElseIf txtIdeal2 = "" Then
        ckIdeal2.Enabled = False
    End If

    If ckIdeal2.Value = 1 And txtIdeal2 <> "" And txtMin <> "" And txtMax <> "" Then
        iStation2min = Int(txtIdeal2) * Int(txtMin) / 100
        txtStation2Min = iStation2min
        iStation2max = Int(txtIdeal2) * Int(txtMax) / 100
        txtStation2Max = iStation2max

    ElseIf ckIdeal2.Value = 1 Then
        MsgBox "Minimum & Maximum percentage power limits and Ideal Power (Watts) text entry boxes must have data", _
        vbOKOnly, "Missing Data"
        ckIdeal2.Value = 0
    End If
    
    lblIdeal2 = txtStation2
    iCodeRequired = 0
End Sub

Private Sub txtIdeal2_GotFocus()
    txtIdeal2.SelStart = 0
    txtIdeal2.SelLength = Len(txtIdeal2)
    cmdSave.Enabled = True
End Sub

Private Sub txtIdeal2_LostFocus()
    If iChange = 0 Then
        iChange = 1
    End If
End Sub

Private Sub txtIdeal3_Change()

    If txtIdeal3 <> "" And txtMin <> "" And txtMax <> "" Then
        ckIdeal3.Enabled = True
    ElseIf txtIdeal3 = "" Then
        ckIdeal3.Enabled = False
    End If

    If ckIdeal3.Value = 1 And txtIdeal3 <> "" And txtMin <> "" And txtMax <> "" Then
        iStation3min = Int(txtIdeal3) * Int(txtMin) / 100
        txtStation3Min = iStation3min
        iStation3max = Int(txtIdeal3) * Int(txtMax) / 100
        txtStation3Max = iStation3max
    
    ElseIf ckIdeal3.Value = 1 Then
        MsgBox "Minimum & Maximum percentage power limits and Ideal Power (Watts) text entry boxes must have data", _
        vbOKOnly, "Missing Data"
        ckIdeal3.Value = 0
    End If
    
    lblIdeal3 = txtStation3
End Sub

Private Sub txtIdeal3_GotFocus()
    txtIdeal3.SelStart = 0
    txtIdeal3.SelLength = Len(txtIdeal3)
    cmdSave.Enabled = True
End Sub

Private Sub txtIdeal3_LostFocus()
    If iChange = 0 Then
        iChange = 1
    End If
End Sub

Private Sub txtIdeal4_Change()

    If txtIdeal4 <> "" And txtMin <> "" And txtMax <> "" Then
            ckIdeal4.Enabled = True
        ElseIf txtIdeal4 = "" Then
            ckIdeal4.Enabled = False
        End If
        
    If ckIdeal4.Value = 1 And txtIdeal4 <> "" And txtMin <> "" And txtMax <> "" Then
        iStation4min = Int(txtIdeal4) * Int(txtMin) / 100
        txtStation4Min = iStation4min
        iStation4max = Int(txtIdeal4) * Int(txtMax) / 100
        txtStation4Max = iStation4max
        
    ElseIf ckIdeal4.Value = 1 Then
        MsgBox "Minimum & Maximum percentage power limits and Ideal Power (Watts) text entry boxes must have data", _
        vbOKOnly, "Missing Data"
        ckIdeal4.Value = 0
    End If
    
    lblIdeal4 = txtStation4
End Sub

Private Sub txtIdeal4_GotFocus()
    txtIdeal4.SelStart = 0
    txtIdeal4.SelLength = Len(txtIdeal4)
    cmdSave.Enabled = True
End Sub

Private Sub txtIdeal4_LostFocus()
    If iChange = 0 Then
        iChange = 1
    End If
End Sub

Private Sub txtIntroOut_Change()

    If Not IsNumeric(txtIntroOut) And txtIntroOut <> "" Then
        MsgBox "You have entered :     " & txtIntroOut & "      This is an incorrect entry. Entry should be an estimate of the" & vbCrLf & _
        "average total number of seconds required to announce and back announce a music selection.", 0, "Entry Error"

        txtIntroOut = ""
        txtIntroOut.SetFocus
    End If
End Sub

Private Sub txtIntroOut_GotFocus()
    txtIntroOut.SelStart = 0
    txtIntroOut.SelLength = Len(txtIntroOut)
    cmdSave.Enabled = True
End Sub

Private Sub txtIntroOut_LostFocus()
    
   If Not IsNumeric(txtIntroOut) Or txtIntroOut = "" Then
        txtIntroOut = "50"
    Else
        txtIntroOut.Text = Format$(txtIntroOut, "###")
    End If
    iCodeRequired = 2
End Sub

Private Sub txtMax_Change()

    If txtMax = "" Then
        ckIdeal1.Enabled = False
        ckIdeal2.Enabled = False
        ckIdeal3.Enabled = False
        ckIdeal4.Enabled = False
    ElseIf txtMax <> "" And txtMin <> "" Then
        If txtIdeal1 <> "" Then
            ckIdeal1.Enabled = True
        End If
        If txtIdeal2 <> "" Then
            ckIdeal2.Enabled = True
        End If
        If txtIdeal3 <> "" Then
            ckIdeal3.Enabled = True
        End If
        If txtIdeal4 <> "" Then
            ckIdeal4.Enabled = True
        End If
    End If

    If txtMax = "." Or Not IsNumeric(txtMax) And txtMax <> "" Then
         MsgBox "Entry represents an approved maximum percent of ideal power and must be between 100 and 150%." _
         & vbCrLf & "105 percent is the default legal power maximum.", vbOKOnly, "Transmitter Power Upper Limit.  Entry Out of Range or Non-Numeric"
         txtMax = ""
    End If
    
    ckIdeal1.Value = 0
    ckIdeal2.Value = 0
    ckIdeal3.Value = 0
    ckIdeal4.Value = 0
    
End Sub

Private Sub txtMax_GotFocus()
    txtMax.SelStart = 0
    txtMax.SelLength = Len(txtMax)
    cmdSave.Enabled = True
End Sub

Private Sub txtMax_LostFocus()

    If iChange = 0 Then
        iChange = 4
    End If
    
    If txtMax = "" Then
         MsgBox "The FCC assigns desired (ideal) transmitter power, and Maximum-Minimum power deviation limits," _
         & vbCrLf & "normally as percents of the desired power. 105% of ideal is the default legal maximum power limit.", vbOKOnly, "Transmitter Power Upper Limit"
        Exit Sub
    End If

    If txtMax <> "" And txtMax < "100" Or txtMax > "150" Then
        MsgBox "Entry represents an approved maximum percent of ideal power and must be between 100 and 150%", vbOKOnly, "Transmitter Power Upper Limit. Entry Out of Range"
        txtMax = ""
        txtMax.SetFocus
    End If
    
End Sub

Private Sub txtMin_Change()

    If txtMin = "" Then
        ckIdeal1.Enabled = False
        ckIdeal2.Enabled = False
        ckIdeal3.Enabled = False
        ckIdeal4.Enabled = False
    ElseIf txtMin <> "" And txtMax <> "" Then
        If txtIdeal1 <> "" Then
            ckIdeal1.Enabled = True
        End If
        If txtIdeal2 <> "" Then
            ckIdeal2.Enabled = True
        End If
        If txtIdeal3 <> "" Then
            ckIdeal3.Enabled = True
        End If
        If txtIdeal4 <> "" Then
            ckIdeal4.Enabled = True
        End If
    End If
    
    If txtMin = "." Or txtMin = "0" Or Not IsNumeric(txtMin) And txtMin <> "" Then
         MsgBox "Entry represents the approved minimum percent of ideal power and must be between 1 and 99." _
         & vbCrLf & "90 percent of ideal is the default legal minimum power limit.", vbOKOnly, "Transmitter Power Lower Limit.  Out of Range or Non-Numeric Entry "
         txtMin = ""
    End If
    ckIdeal1.Value = 0
    ckIdeal2.Value = 0
    ckIdeal3.Value = 0
    ckIdeal4.Value = 0
        
End Sub

Private Sub txtMin_GotFocus()
    txtMin.SelStart = 0
    txtMin.SelLength = Len(txtMin)
    cmdSave.Enabled = True
End Sub

Private Sub txtMin_LostFocus()

'    If iChange = 0 Then
'        iChange = 4
'    End If
    
    If txtMin = "" Then
         MsgBox "The FCC assigns desired (ideal) transmitter power, and Minimum-Maximum power deviation limits," _
         & vbCrLf & "normally as percents of the desired power. 90% of ideal is the default minimum legal power.", vbOKOnly, "Transmitter Power Lower Limit"
    End If
End Sub

Private Sub txtMin_MouseUp(Button As Integer, Shift As Integer, iHourNow As Single, Y As Single)
 If iChange = 0 Then
        iChange = 4
    End If
End Sub

Private Sub txtPlanTime_Change()
    If txtPlanTime <> "" Then
        If Not IsNumeric(txtPlanTime) Then
            MsgBox "You have entered :    " & txtPlanTime & "   This is an incorrect entry. Entry should" & vbCrLf & _
            "be the NUMBER of minutes of music planned for the hour." _
            & vbCrLf & vbCrLf & "Entry cannot exceed 60 minutes or be less than 1 minute.", 0, "Planned Time Entry Error"
            
            txtPlanTime = ""
            txtPlanTime.SetFocus
            Exit Sub
        ElseIf Val(txtPlanTime) > 60.1 Or txtPlanTime < 1 Then
            MsgBox "Entry cannot exceed 60 minutes or be less than 1 minute.", vbExclamation, "Entry Error"
            txtPlanTime = ""
        End If
    End If
    
    If txtPlanTime.Text = "56.0" Then
        txtPlanTime.ForeColor = &H80000008  'black
         txtPlanTime.Appearance = 1
        txtPlanTime.BorderStyle = 1
    Else
        txtPlanTime.ForeColor = &H80& ' rust '&HC00000    'blue
        txtPlanTime.Appearance = 0
        txtPlanTime.BorderStyle = 1
    End If
   
End Sub

Private Sub txtPlanTime_GotFocus()
    txtPlanTime.SelStart = 0
    txtPlanTime.SelLength = Len(txtPlanTime)
    cmdSave.Enabled = True
End Sub

Private Sub txtPlanTime_LostFocus()

    If txtPlanTime <> "" And IsNumeric(txtPlanTime) Then
        giPlannedTime = Format$(txtPlanTime, "#0.0")
    Else
        txtPlanTime = "56.0"
        giPlannedTime = "56.0"
    End If
   
    frmPlanner!txtBlock = Format$(txtPlanTime, "#0.0")

    If Not IsNumeric(txtPlanTime) Or txtPlanTime = "" Then
        txtPlanTime = "56.0"
   Else
        txtPlanTime.Text = Format$(txtPlanTime, "#0.0")
    End If
    iCodeRequired = 1
End Sub

Private Sub txtSignatureLine_Change()

imgSmile(0).Visible = False
Label25.Visible = False

On Error GoTo HandleErrors

    Dim Signature As String

    Open "Signature.dat" For Input As #27
    Input #27, Signature
    Close #27

    If txtSignatureLine.Text = Signature Then 'compare the two. No change to signature line
        imgSmile(0).Visible = False
        Label25.Visible = False
        cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
        Exit Sub
    End If

    cmdSetSignature.Caption = "Click to Save or Cancel Change" 'signature line changed
    imgHand.Visible = True
'    cmdSave.Enabled = False
'    imgSmile(0).Visible = True
'    Label25.Caption = "Cancel"
'    Label25.Visible = True
HandleErrors:
End Sub

Private Sub txtSignatureLine_GotFocus()
    txtSignatureLine.SelStart = 0
    txtSignatureLine.SelLength = Len(txtSignatureLine)
End Sub

Private Sub txtSignatureLine_LostFocus()


On Error GoTo HandleErrors

    Dim Signature As String

    Open "Signature.dat" For Input As #27
    Input #27, Signature
    Close #27

    If txtSignatureLine.Text = Signature Then 'compare the two. No change to signature line
        imgSmile(0).Visible = False
        Label25.Visible = False
        cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
MsgBox "No change was made to the existing default signature line.", vbOKOnly, "No Change"
        Exit Sub
    End If

    cmdSetSignature.Caption = "Click to Save or Cancel Change" 'signature line changed
    cmdSave.Enabled = False
'    imgSmile(0).Visible = True
'    Label25.Caption = "Cancel"
'    Label25.Visible = True
HandleErrors:


' Dim Signature As String
'
'    Open "Signature.dat" For Input As #27
'    Input #27, Signature
'    Close #27
'
'    If txtSignatureLine.Text = Signature Then 'compare the two. No change to signature line
'        imgSmile(0).Visible = False
'        Label25.Visible = False
'        cmdSetSignature.Caption = "Set as Music Logbook Page Printout Default Signature"
'        MsgBox "No change was made to the existing default signature line.", vbOKOnly, "No Change"
'        Exit Sub
'    End If

    If Len(txtSignatureLine) <= 2 Then
        MsgBox "Signature must be at least 3 characters in length. Default signature is 'Fine Arts Radio'.", vbOKOnly, "Entry Too Short"
        txtSignatureLine = "Fine Arts Radio"
    End If
    
    frmPlanner!txtSignature.ToolTipText = "Signature Line. Default, " & txtSignatureLine
End Sub

Private Sub txtSpot_Change()
   If Not IsNumeric(txtSpot) And txtSpot <> "" Then
        MsgBox "You have entered :     " & txtSpot & "      This is an incorrect entry. Entry should be the estimated" & vbCrLf & _
        "NUMBER of seconds required for an average spot announcement or PSA .", 0, "Entry Error"

        txtSpot = ""
        txtSpot.SetFocus
    End If
End Sub

Private Sub txtSpot_GotFocus()
    txtSpot.SelStart = 0
    txtSpot.SelLength = Len(txtSpot)
    cmdSave.Enabled = True
End Sub

Private Sub txtSpot_LostFocus()
   
    If Not IsNumeric(txtSpot) Or txtSpot = "" Then
        txtSpot = "30"
    Else
        txtSpot.Text = Format$(txtSpot, "###")
    End If
    iCodeRequired = 3
End Sub

Private Sub txtStation1_Change()

    If txtStation1 = "" Then
        Label5.ForeColor = vbRed
    Else
        Label5.ForeColor = &HC00000    'blue
    End If
    
    lblIdeal1 = UCase(txtStation1)
    
    If txtStation1 = "" Then
        txtStation1Min = ""
        txtStation1Max = ""
        txtIdeal1 = ""
        ckIdeal1.Value = 0
    End If
End Sub

Private Sub txtStation1_GotFocus()
    txtStation1.SelStart = 0
    txtStation1.SelLength = Len(txtStation1)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtStation1_LostFocus()
    
    txtStation1 = UCase(txtStation1)
    iCodeRequired = 5
End Sub

Private Sub txtStation1Max_Change()
    If IsNumeric(txtStation1Max) Or txtStation1Max = "" Then
    Else
         MsgBox "Non-Numerical Entry", vbOKOnly, "Entry Error"
         txtStation1Max = ""
    End If
    
    If Val(txtStation1Max) <> iStation1max Then
      ckIdeal1.Value = 0
    End If
End Sub

Private Sub txtStation1Max_GotFocus()
    txtStation1Max.SelStart = 0
    txtStation1Max.SelLength = Len(txtStation1Max)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation1Max_LostFocus()
    txtStation1Max.Text = Format$(txtStation1Max, "######")
    
    If iChange = 0 Then
        iChange = 3
    End If
    
    If Val(txtStation1Max) <> iStation1max Then
        ckIdeal1.Value = 0
    End If
End Sub

Private Sub txtStation1Min_Change()
    If IsNumeric(txtStation1Min) Or txtStation1Min = "" Then
    Else
         MsgBox "Non-Numerical Entry", vbOKOnly, "Entry Error"
         txtStation1Min = ""
    End If

    If Val(txtStation1Min) <> iStation1min Then '<> Val(txtIdeal1) * Val(txtMin) / 100 Then
      ckIdeal1.Value = 0
    End If
 
End Sub

Private Sub txtStation1Min_GotFocus()
    txtStation1Min.SelStart = 0
    txtStation1Min.SelLength = Len(txtStation1Min)
    
    cmdSave.Enabled = True
End Sub

Private Sub txtStation1Min_LostFocus()
    txtStation1Min.Text = Format$(txtStation1Min, "######")
    
    If iChange = 0 Then
        iChange = 2
    End If
    
    If Val(txtStation1Min) <> iStation1min Then
        ckIdeal1.Value = 0
    End If
End Sub

Private Sub txtStation2_Change()
    lblIdeal2 = UCase(txtStation2)
    
    If txtStation2 = "" Then
        txtStation2Min = ""
        txtStation2Max = ""
        txtIdeal2 = ""
        ckIdeal2.Value = 0
    End If
    
End Sub

Private Sub txtStation2_GotFocus()
    txtStation2.SelStart = 0
    txtStation2.SelLength = Len(txtStation2)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtStation2_LostFocus()
    iCodeRequired = 5
    txtStation2 = UCase(txtStation2)
End Sub

Private Sub txtStation2Max_Change()
    If IsNumeric(txtStation2Max) Or txtStation2Max = "" Then
    Else
         MsgBox "Non-Numerical Entry", vbOKOnly, "Entry Error"
         txtStation2Max = ""
    End If
    
    If Val(txtStation2Max) <> iStation2max Then
      ckIdeal2.Value = 0
    End If
End Sub

Private Sub txtStation2Max_GotFocus()
    txtStation2Max.SelStart = 0
    txtStation2Max.SelLength = Len(txtStation2Max)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation2Max_LostFocus()
    txtStation2Max.Text = Format$(txtStation2Max, "######")
    
    If iChange = 0 Then
        iChange = 3
    End If
     
    If Val(txtStation2Max) <> iStation2max Then
        ckIdeal2.Value = 0
    End If
End Sub

Private Sub txtStation2Min_Change()
    If IsNumeric(txtStation2Min) Or txtStation2Min = "" Then
    Else
         MsgBox "Non-Numerical Entry", vbOKOnly, "Entry Error"
         txtStation2Min = ""
    End If
    
    If Val(txtStation2Min) <> iStation2min Then
      ckIdeal2.Value = 0
    End If
End Sub

Private Sub txtStation2Min_GotFocus()
    txtStation2Min.SelStart = 0
    txtStation2Min.SelLength = Len(txtStation2Min)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation2Min_LostFocus()
    txtStation2Min.Text = Format$(txtStation2Min, "######")
    
    If iChange = 0 Then
        iChange = 2
    End If
    
    If Val(txtStation2Min) <> iStation2min Then
    ckIdeal2.Value = 0
    End If
End Sub

Private Sub txtStation3_Change()
     lblIdeal3 = UCase(txtStation3)
     
     If txtStation3 = "" Then
        txtStation3Min = ""
        txtStation3Max = ""
        txtIdeal3 = ""
        ckIdeal3.Value = 0
    End If
End Sub

Private Sub txtStation3_GotFocus()
    txtStation3.SelStart = 0
    txtStation3.SelLength = Len(txtStation3)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtStation3_LostFocus()
    iCodeRequired = 5
    txtStation3 = UCase(txtStation3)
End Sub

Private Sub txtStation3Max_Change()
    If IsNumeric(txtStation3Max) Or txtStation3Max = "" Then
    Else
         MsgBox "Non-Numerical Entry", vbOKOnly, "Entry Error"
         txtStation3Max = ""
    End If
    
    If Val(txtStation3Max) <> iStation3max Then
      ckIdeal3.Value = 0
    End If
End Sub

Private Sub txtStation3Max_GotFocus()
    txtStation3Max.SelStart = 0
    txtStation3Max.SelLength = Len(txtStation3Max)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation3Max_LostFocus()
    txtStation3Max.Text = Format$(txtStation3Max, "######")
    
    If iChange = 0 Then
        iChange = 3
    End If
   
    If Val(txtStation3Max) <> iStation3max Then
        ckIdeal3.Value = 0
    End If
End Sub

Private Sub txtStation3Min_Change()
    If IsNumeric(txtStation3Min) Or txtStation3Min = "" Then
    Else
         MsgBox "Non-Numerical Entry", vbOKOnly, "Entry Error"
         txtStation3Min = ""
    End If
    
    If Val(txtStation3Min) <> iStation3min Then
      ckIdeal3.Value = 0
    End If
End Sub

Private Sub txtStation3Min_GotFocus()
    txtStation3Min.SelStart = 0
    txtStation3Min.SelLength = Len(txtStation3Min)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation3Min_LostFocus()
    txtStation3Min.Text = Format$(txtStation3Min, "######")
    
    If iChange = 0 Then
        iChange = 2
    End If
    
    If Val(txtStation3Min) <> iStation3min Then
        ckIdeal3.Value = 0
    End If
End Sub

Private Sub txtStation4_Change()
     lblIdeal4 = UCase(txtStation4)
     
     If txtStation4 = "" Then
        txtStation4Min = ""
        txtStation4Max = ""
        txtIdeal4 = ""
        ckIdeal4.Value = 0
    End If
End Sub

Private Sub txtStation4_GotFocus()
    txtStation4.SelStart = 0
    txtStation4.SelLength = Len(txtStation4)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        KeyAscii = 39
    End If
End Sub

Private Sub txtStation4_LostFocus()

     txtStation4 = UCase(txtStation4)
    iCodeRequired = 5
    If txtStation1Min = "" Or txtStation1Max = "" Then
    
    MsgBox "If the Transmitter Power computation page is to be used:" & vbCrLf & vbCrLf & _
    " (1) On this page, using either percentage of ideal power (recommended), or direct entry, set and save" _
    & vbCrLf & "the minimum and maximum Power Limits for each transmitter." _
        & vbCrLf & vbCrLf & " (2) Next, go to the 'Transmitter' page. From the 'Settings' menu, select 'Change Transmitter Efficiencies'." _
        & vbCrLf & vbCrLf & "Enter the efficiencies for each transmitter. Efficiencies normally will range from a low of about" _
        & vbCrLf & ".50 (50%) to a high of about .90 (90%). In all cases, efficiencies will be less than .99 (99%)." & vbCrLf & vbCrLf & "Note: These instructions are repeated in the 'Transmitter' page 'User Hints'." _
        , vbOKOnly, "Setting up transmitter power computation page"
        
'        txtIdeal1.SetFocus
    End If
End Sub

Private Sub txtStation4Max_Change()
    If IsNumeric(txtStation4Max) Or txtStation4Max = "" Then
    Else
         MsgBox "Non-Numerical Entry", vbOKOnly, "Entry Error"
         txtStation4Max = ""
    End If
    
    If Val(txtStation4Max) <> iStation4max Then
      ckIdeal4.Value = 0
    End If
    
End Sub

Private Sub txtStation4Max_GotFocus()
    txtStation4Max.SelStart = 0
    txtStation4Max.SelLength = Len(txtStation4Max)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation4Max_LostFocus()
    txtStation4Max.Text = Format$(txtStation4Max, "######")
    
    If iChange = 0 Then
        iChange = 3
    End If
    
    If Val(txtStation4Max) <> iStation4max Then
        ckIdeal4.Value = 0
    End If
End Sub

Private Sub txtStation4Min_Change()
    If IsNumeric(txtStation4Min) Or txtStation4Min = "" Then
    Else
         MsgBox "Non-Numerical Entry", vbOKOnly, "Entry Error"
         txtStation4Min = ""
    End If
    
    If Val(txtStation4Min) <> iStation4min Then
      ckIdeal4.Value = 0
    End If
End Sub

Private Sub txtStation4Min_GotFocus()
    txtStation4Min.SelStart = 0
    txtStation4Min.SelLength = Len(txtStation4Min)
    cmdSave.Enabled = True
End Sub

Private Sub txtStation4Min_LostFocus()
    txtStation4Min.Text = Format$(txtStation4Min, "######")
    
    If iChange = 0 Then
        iChange = 2
    End If
    
    If Val(txtStation4Min) <> iStation4min Then
        ckIdeal4.Value = 0
    End If
End Sub
