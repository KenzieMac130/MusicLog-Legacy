VERSION 5.00
Begin VB.Form frmAddTime 
   BorderStyle     =   0  'None
   Caption         =   "Add Time  -  F9"
   ClientHeight    =   6525
   ClientLeft      =   1365
   ClientTop       =   2220
   ClientWidth     =   4305
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   Icon            =   "frmAddKeypad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraAddTime 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   225
      TabIndex        =   27
      Top             =   90
      Width           =   3855
      Begin VB.CommandButton cmdUndoClear 
         Caption         =   "&Previous Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2115
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4890
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000080&
         Height          =   1020
         Left            =   210
         TabIndex        =   37
         Top             =   4980
         Width           =   1815
         Begin VB.Label lblTotal2 
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
            Height          =   690
            Left            =   90
            OLEDropMode     =   1  'Manual
            TabIndex        =   38
            Top             =   240
            Width           =   1650
         End
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "E&xit Page  F9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2265
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   5460
         Width           =   1245
      End
      Begin VB.Frame fraFrame1 
         Caption         =   "Enter Times"
         ForeColor       =   &H00000080&
         Height          =   4275
         Left            =   240
         TabIndex        =   34
         Top             =   225
         Width           =   1545
         Begin VB.TextBox txtSecond5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   1725
            Width           =   465
         End
         Begin VB.TextBox txtMinute5 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   1725
            Width           =   465
         End
         Begin VB.TextBox txtMinute1 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   375
            Width           =   465
         End
         Begin VB.TextBox txtMinute2 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   690
            Width           =   465
         End
         Begin VB.TextBox txtMinute3 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   1020
            Width           =   465
         End
         Begin VB.TextBox txtMinute4 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   1410
            Width           =   465
         End
         Begin VB.TextBox txtSecond1 
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   1
            Top             =   375
            Width           =   465
         End
         Begin VB.TextBox txtSecond2 
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   690
            Width           =   465
         End
         Begin VB.TextBox txtSecond3 
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1020
            Width           =   465
         End
         Begin VB.TextBox txtSecond4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   1410
            Width           =   465
         End
         Begin VB.TextBox txtSecond7 
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   2430
            Width           =   465
         End
         Begin VB.TextBox txtMinute7 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   2430
            Width           =   465
         End
         Begin VB.TextBox txtSecond10 
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   3465
            Width           =   465
         End
         Begin VB.TextBox txtMinute10 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   3465
            Width           =   465
         End
         Begin VB.TextBox txtMinute11 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   3795
            Width           =   465
         End
         Begin VB.TextBox txtSecond11 
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   3795
            Width           =   465
         End
         Begin VB.TextBox txtMinute8 
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
            Left            =   360
            MaxLength       =   3
            TabIndex        =   14
            Top             =   2745
            Width           =   465
         End
         Begin VB.TextBox txtSecond8 
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
            Left            =   855
            MaxLength       =   2
            TabIndex        =   15
            Top             =   2745
            Width           =   465
         End
         Begin VB.TextBox txtMinute9 
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
            Left            =   360
            MaxLength       =   3
            TabIndex        =   16
            Top             =   3060
            Width           =   465
         End
         Begin VB.TextBox txtSecond9 
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
            Left            =   855
            MaxLength       =   2
            TabIndex        =   17
            Top             =   3060
            Width           =   465
         End
         Begin VB.TextBox txtSecond6 
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
            Left            =   855
            MaxLength       =   2
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   2040
            Width           =   465
         End
         Begin VB.TextBox txtMinute6 
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
            Left            =   360
            MaxLength       =   3
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "1"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   54
            Top             =   390
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "2"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   53
            Top             =   735
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "3"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   52
            Top             =   1080
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "4"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   51
            Top             =   1455
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "5"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   50
            Top             =   1770
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "6"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   5
            Left            =   90
            TabIndex        =   49
            Top             =   2100
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "7"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   6
            Left            =   90
            TabIndex        =   48
            Top             =   2460
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "8"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   7
            Left            =   90
            TabIndex        =   47
            Top             =   2790
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "9"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   8
            Left            =   90
            TabIndex        =   46
            Top             =   3120
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "10"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   9
            Left            =   90
            TabIndex        =   45
            Top             =   3495
            Width           =   195
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "11"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   10
            Left            =   90
            TabIndex        =   44
            Top             =   3840
            Width           =   195
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
            Height          =   255
            Left            =   390
            TabIndex        =   36
            Top             =   150
            Width           =   375
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
            Height          =   255
            Left            =   900
            TabIndex        =   35
            Top             =   150
            Width           =   375
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000002&
            DrawMode        =   6  'Mask Pen Not
            Index           =   0
            X1              =   375
            X2              =   1300
            Y1              =   3405
            Y2              =   3405
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000002&
            DrawMode        =   6  'Mask Pen Not
            Index           =   2
            X1              =   360
            X2              =   1305
            Y1              =   2370
            Y2              =   2370
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000002&
            DrawMode        =   6  'Mask Pen Not
            Index           =   3
            X1              =   360
            X2              =   1305
            Y1              =   1350
            Y2              =   1350
         End
      End
      Begin VB.CommandButton cmdClearEntries 
         Caption         =   "&Clear Entries"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2115
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   " Clears all 'Min' and 'Sec' entries "
         Top             =   4440
         Width           =   1545
      End
      Begin VB.Frame fraFrame3 
         Caption         =   "Block/Remain"
         ForeColor       =   &H00000080&
         Height          =   1485
         Left            =   2190
         TabIndex        =   30
         Top             =   1470
         Width           =   1395
         Begin VB.CommandButton cmdClearBlock 
            Caption         =   "Set &Block Time"
            Height          =   285
            Left            =   128
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   " Clears 'Block/Remain' enty only "
            Top             =   1110
            Width           =   1125
         End
         Begin VB.TextBox txtBlock 
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
            Left            =   210
            MaxLength       =   5
            MultiLine       =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Enter the number of minutes from which the time entered in the KeyPad will be subtracted."
            Top             =   210
            Width           =   585
         End
         Begin VB.Label Label5 
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
            Height          =   210
            Left            =   870
            TabIndex        =   33
            Top             =   240
            Width           =   300
         End
         Begin VB.Label lblLabel6 
            Alignment       =   2  'Center
            Height          =   165
            Left            =   75
            TabIndex        =   32
            Top             =   540
            Width           =   1245
         End
         Begin VB.Label lblRemain 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   75
            TabIndex        =   31
            ToolTipText     =   """Block"" Time remaining after KeyPad Min-Sec entries have been subtracted."
            Top             =   750
            Width           =   1230
         End
      End
      Begin VB.CheckBox chkSubtract 
         Caption         =   "&Subtract Time"
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
         Left            =   2250
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Time entered on line 9 (the minus line) is subtracted from time entered on line 8 (the plus line)"
         Top             =   3855
         Width           =   1350
      End
      Begin VB.CommandButton cmdAdditional 
         Caption         =   "&Add Additional Times to Total"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2152
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   " Totals current times and clears entry boxes to allow additional times to be added to current total "
         Top             =   3090
         Width           =   1470
      End
      Begin VB.Label lblTotal1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   210
         TabIndex        =   43
         Top             =   4575
         Width           =   1605
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "© James Wright"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2257
         TabIndex        =   42
         Top             =   5865
         Width           =   1260
      End
      Begin VB.Label lblLabelPlus 
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   41
         ToolTipText     =   "Click ' + ' to clear line 10"
         Top             =   3675
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblLabelMinus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1927
         TabIndex        =   40
         ToolTipText     =   "Click ' - ' to clear line 11"
         Top             =   3975
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Use the TAB Key to advance the cursor in sequence through the Min and Sec time entry boxes."
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
         Height          =   1095
         Left            =   2145
         TabIndex        =   39
         Top             =   285
         Width           =   1485
      End
   End
   Begin VB.Menu mnuPage 
      Caption         =   "Pa&ge"
      Begin VB.Menu mnuPagePlanner 
         Caption         =   "&Music Planner..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuPageTimeRemain 
         Caption         =   "&Time Remain Page..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPageXmitter 
         Caption         =   "Transmitter &Log..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu sepPage1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageCloseKeypad 
         Caption         =   "&Close AddTime KeyPad"
         Shortcut        =   {F9}
      End
      Begin VB.Menu sepPage2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageStopWatch 
         Caption         =   "&Stopwatch..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPageSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPagePrintPage 
         Caption         =   "&Print a Copy of this Page"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&How to use this page"
   End
End
Attribute VB_Name = "frmAddTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim iSecond10 As Integer
    Dim cMCal3 As Currency
    Dim cMCal4 As Currency
    Dim iSubOption As Integer
    Dim cSubtract3 As Currency
    Dim cSubtract4 As Currency
    Dim iMin10 As String
    Dim iSec10 As String
    Dim iMin11 As String
    Dim iSec11 As String
    Dim iFocus8 As Integer 'prevents subtract query message while txtMinute10 has focus
    Dim iUndoClear As Integer 'sets whether clears txtMinute11/txtSeconds9 or txtMinute10/txtSeconds8 & 9
'    Dim giEntrySave As Integer 'alternates saves for previous save command
Option Explicit
Private Sub pAdd()
    'this is the "motor" that drives the add functions
    Dim cBlock As Currency
    Dim iMinute1 As Integer
    Dim iMinute2 As Integer
    Dim iMinute3 As Integer
    Dim iMinute4 As Integer
    Dim iMinute5 As Integer
    Dim iMinute6 As Integer
    Dim iMinute7 As Integer
    Dim iMinute8 As Integer
    Dim iMinute9 As Integer
    Dim iMinute10 As Integer
    Dim iMinute11 As Integer
    
    Dim cTotalMin As Currency
    Dim cTotalSec As Currency
    Dim iSecond1 As Integer
    Dim iSecond2 As Integer
    Dim iSecond3 As Integer
    Dim iSecond4 As Integer
    Dim iSecond5 As Integer
    Dim iSecond6 As Integer
    Dim iSecond7 As Integer
    Dim iSecond8 As Integer
    Dim iSecond9 As Integer
    Dim iSecond10 As Integer
    Dim iSecond11 As Integer
    Dim cCombined As Currency
    Dim cHours As Currency
    Dim cSecAdd As Integer
    Dim cMinAdd As Integer
    Dim cMCal1 As Currency
    Dim cMCal2 As Currency
    Dim cHCal1 As Currency
    Dim cHCal2 As Currency
    Dim cHCal3 As Currency
    Dim cHCal4 As Currency
    Dim cSec As Long
     
    'set "Min" values & formulate as numeric
    If IsNumeric(txtMinute1) Or txtMinute1 = "" Or txtMinute1 = "-" Then
         iMinute1 = Val(txtMinute1)
    Else
        MsgBox "Min 1, Enter Integer", vbOKOnly, "Entry Error"
        txtMinute1 = ""
    End If

    If IsNumeric(txtMinute2) Or txtMinute2 = "" Or txtMinute2 = "-" Then
        iMinute2 = Val(txtMinute2)
    Else
         MsgBox "Min 2, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute2 = ""
    End If

    If IsNumeric(txtMinute3) Or txtMinute3 = "" Or txtMinute3 = "-" Then
        iMinute3 = Val(txtMinute3)
    Else
         MsgBox "Min 3, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute3 = ""
    End If

    If IsNumeric(txtMinute4) Or txtMinute4 = "" Or txtMinute4 = "-" Then
         iMinute4 = Val(txtMinute4)
    Else
         MsgBox "Min 4, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute4 = ""
    End If

    If IsNumeric(txtMinute5) Or txtMinute5 = "" Or txtMinute5 = "-" Then
        iMinute5 = Val(txtMinute5)
    Else
         MsgBox "Min 5, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute5 = ""
    End If

    If IsNumeric(txtMinute6) Or txtMinute6 = "" Or txtMinute6 = "-" Then
        iMinute6 = Val(txtMinute6)
    Else
         MsgBox "Min 6, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute6 = ""
    End If

    If IsNumeric(txtMinute7) Or txtMinute7 = "" Or txtMinute7 = "-" Then
        iMinute7 = Val(txtMinute7)
    Else
         MsgBox "Min 7, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute7 = ""
    End If
    
    If IsNumeric(txtMinute8) Or txtMinute8 = "" Or txtMinute8 = "-" Then
        iMinute8 = Val(txtMinute8)
    Else
         MsgBox "Min 8, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute8 = ""
    End If
    
    If IsNumeric(txtMinute9) Or txtMinute9 = "" Or txtMinute9 = "-" Then
        iMinute9 = Val(txtMinute9)
    Else
         MsgBox "Min 9, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute9 = ""
    End If

    If IsNumeric(txtMinute10) Or txtMinute10 = "" Or txtMinute10 = "-" Then
        iMinute10 = Val(txtMinute10)
    Else
         MsgBox "Min 8, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute10 = ""
    End If

    If IsNumeric(txtMinute11) Or txtMinute11 = "" Or txtMinute11 = "-" Then
        iMinute11 = Val(txtMinute11)
    Else
         MsgBox "Min 9, Enter Integer", vbOKOnly, "Entry Error"
         txtMinute11 = ""
    End If

'------set "Sec" values & formulate as numeric
    If IsNumeric(txtSecond1) Or txtSecond1 = "" Or txtSecond1 = "-" Then
        iSecond1 = Val(txtSecond1)
    Else
         MsgBox "Second 1, Enter Integer", vbOKOnly, "Entry Error"
         txtSecond1 = ""
    End If

    If Val(txtSecond1) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond1 = ""
        Exit Sub
    End If

    If IsNumeric(txtSecond2) Or txtSecond2 = "" Or txtSecond2 = "-" Then
         iSecond2 = Val(txtSecond2)
    Else
         MsgBox "Second 2, Enter Integer", vbOKOnly, "Entry Error"
         txtSecond2 = ""
    End If

    If Val(txtSecond2) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond2 = ""
        Exit Sub
    End If

    If IsNumeric(txtSecond3) Or txtSecond3 = "" Or txtSecond3 = "-" Then
         iSecond3 = Val(txtSecond3)
    Else
         MsgBox "Second 3, Enter Integer", vbOKOnly, "Entry Error"
         txtSecond3 = ""
    End If

    If Val(txtSecond3) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond3 = ""
        Exit Sub
    End If

    If IsNumeric(txtSecond4) Or txtSecond4 = "" Or txtSecond4 = "-" Then
        iSecond4 = Val(txtSecond4)
    Else
        MsgBox "Second 4, Enter Integer", vbOKOnly, "Entry Error"
        txtSecond4 = ""
    End If

    If Val(txtSecond4) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond4 = ""
        Exit Sub
    End If

    If IsNumeric(txtSecond5) Or txtSecond5 = "" Or txtSecond5 = "-" Then
        iSecond5 = Val(txtSecond5)
    Else
        MsgBox "Second 5, Enter Integer", vbOKOnly, "Entry Error"
        txtSecond5 = ""
    End If

    If Val(txtSecond5) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond5 = ""
        Exit Sub
    End If

    If IsNumeric(txtSecond6) Or txtSecond6 = "" Or txtSecond6 = "-" Then
        iSecond6 = Val(txtSecond6)
    Else
        MsgBox "Second 6, Enter Integer", vbOKOnly, "Entry Error"
        txtSecond6 = ""
    End If

    If Val(txtSecond6) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond6 = ""
        Exit Sub
    End If

    If IsNumeric(txtSecond7) Or txtSecond7 = "" Or txtSecond7 = "-" Then
        iSecond7 = Val(txtSecond7)
    Else
        MsgBox "Second 7, Enter Integer", vbOKOnly, "Entry Error"
        txtSecond7 = ""
    End If

    If Val(txtSecond7) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond7 = ""
        Exit Sub
    End If
    
    If IsNumeric(txtSecond8) Or txtSecond8 = "" Or txtSecond8 = "-" Then
        iSecond8 = Val(txtSecond8)
    Else
        MsgBox "Second 8, Enter Integer", vbOKOnly, "Entry Error"
        txtSecond8 = ""
    End If
    
    If Val(txtSecond8) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond8 = ""
        Exit Sub
    End If
    
    If IsNumeric(txtSecond9) Or txtSecond9 = "" Or txtSecond9 = "-" Then
        iSecond9 = Val(txtSecond9)
    Else
        MsgBox "Second 9, Enter Integer", vbOKOnly, "Entry Error"
        txtSecond9 = ""
    End If
    
    If Val(txtSecond9) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond9 = ""
        Exit Sub
    End If

    If IsNumeric(txtSecond10) Or txtSecond10 = "" Or txtSecond10 = "-" Then
        iSecond10 = Val(txtSecond10)
    Else
        MsgBox "Second 10, Enter Integer", vbOKOnly, "Entry Error"
        txtSecond10 = ""
    End If

    If Val(txtSecond10) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond10 = ""
        Exit Sub
    End If

     If IsNumeric(txtSecond11) Or txtSecond11 = "" Or txtSecond11 = "-" Then
        iSecond11 = Val(txtSecond11)
    Else
        MsgBox "Second 11, Enter Integer", vbOKOnly, "Entry Error"
        txtSecond11 = ""
    End If

    If Val(txtSecond11) > 59 Then
        MsgBox "Entry may not exceed 59 seconds", 0, "Error"
        txtSecond11 = ""
        Exit Sub
    End If
    
'----------Subtract mode -------------------
'----------------------------------------------

    Dim cSubtract As Currency
    Dim cSubtract2 As Currency
    Dim iResponse As Integer

    If chkSubtract.Value = 1 Then

        If (((iMinute10 * 60) + iSecond10) >= ((iMinute11 * 60) + iSecond11)) Then
        
            cSubtract = (((iMinute10 * 60) + iSecond10) - ((iMinute11 * 60) + iSecond11)) / 60
            
        ElseIf (((iMinute10 * 60) + iSecond10) < ((iMinute11 * 60) + iSecond11)) Then
        
            cSubtract = (((iMinute11 * 60) + iSecond11) - ((iMinute10 * 60) + iSecond10)) / 60
            
        End If
         
        cSubtract = (((iMinute10 * 60) + iSecond10) - ((iMinute11 * 60) + iSecond11)) / 60
        
        cSubtract2 = cSubtract - Int(cSubtract) 'extracts the fraction of min
        cSubtract3 = cSubtract - cSubtract2 'removes fraction, leaving whole min
        cSubtract4 = cSubtract2 * 60 'changes fraction of minute to seconds
    
        If (((iMinute10 * 60) + iSecond10) >= ((iMinute11 * 60) + iSecond11)) / 60 Then
        
            lblTotal1.Caption = Format$(cSubtract3, "#0") _
            & " min " & Format$(cSubtract4, "00") & " sec"
            
            
        ElseIf (((iMinute10 * 60) + iSecond10) < ((iMinute11 * 60) + iSecond11)) / 60 Then
             lblTotal1.Caption = Format$(cSubtract3, "-#0") & " min " & Format$(cSubtract4, "00") & " sec"
        End If
        
        Exit Sub
    
    '------------end subtract mode
    ElseIf chkSubtract.Value = 0 Then

        'calculation for minutes/seconds display
        cSecAdd = iSecond1 + iSecond2 + iSecond3 + iSecond4 + iSecond5 + iSecond6 + iSecond7 + iSecond8 + iSecond9 + iSecond10 + iSecond11
        cMinAdd = iMinute1 + iMinute2 + iMinute3 + iMinute4 + iMinute5 + iMinute6 + iMinute7 + iMinute8 + iMinute9 + iMinute10 + iMinute11
        cMCal1 = cMinAdd + cSecAdd / 60 'adds min & sec/60
        
        cMCal2 = cMCal1 - Int(cMCal1) 'extracts the fraction of min
        cMCal3 = cMCal1 - cMCal2 'removes fraction, leaving whole min
    
        cMCal4 = cMCal2 * 60 'multiplies fraction by 60 producing seconds
        'display cMCal3 for minutes; display cMCal4 for seconds
        
        'calculation for hours/minutes display
        cTotalMin = (iMinute1 + iMinute2 + iMinute3 + iMinute4 + iMinute5 + iMinute6 + iMinute7 + iMinute8 + iMinute9 + iMinute10 + iMinute11)
        cTotalSec = (iSecond1 + iSecond2 + iSecond3 + iSecond4 + iSecond5 + iSecond6 + iSecond7 + iSecond8 + iSecond9 + iSecond10 + iSecond11) / 60
       
        'converts total minutes to seconds & adds to total seconds
        cSec = (cTotalMin * 60) + (cTotalSec * 60)
       
       'combines minutes & seconds into decimal hours
        cCombined = cTotalMin + cTotalSec
        cHours = cCombined / 60
         
        cHCal2 = cHours - Int(cHours) 'extracts fraction of hour
        cHCal3 = cHours - cHCal2 'removes fraction leaving whole hour
        cHCal4 = cHCal2 * 60 'multiplies fraction by 60 producing minutes
        'display cHCal3 for hours; display cHCal4 for minutes

        Dim cHCal5 As String
        If cHCal3 <= 0 Then
            cHCal5 = ""
        Else: cHCal5 = " hr "
        End If
    End If
'1 --------------calculate block-remain & display values in minutes & seconds--------------

    Dim cRemain As Currency
    Dim cRemain1 As Currency
    Dim cRemain2 As Currency
    Dim cRemain3 As Currency
    Dim cRemain4 As Currency
    Dim cRemain5 As Currency
    Dim cRemain6 As Currency
    Dim cRemain7 As Currency
    
    If txtBlock <> "" Then
    
        cBlock = Val(txtBlock)
        cRemain = cBlock - cCombined
        
        cRemain1 = cRemain - Int(cRemain) 'removes integer
    
        If cRemain >= 0 Then
            cRemain2 = cRemain - cRemain1
            cRemain3 = cRemain1 * 60
            lblRemain.Caption = Format$(cRemain2, "0") & " min " & Format$(cRemain3, "#00") & " sec"
        Else
            cRemain7 = cRemain * (-1)
            cRemain4 = cRemain7 - Int(cRemain7)
            cRemain5 = cRemain7 - cRemain4
            cRemain6 = cRemain4 * 60
            lblRemain.Caption = Format$(cRemain5, "0") & " min " & Format$(cRemain6, "#00") & " sec"
        End If

  'sets Block Time label readings
        If cRemain >= 0 Then
            lblLabel6.ForeColor = vbBlack
            lblLabel6.Caption = "Time Remaining"
        Else
            lblLabel6.ForeColor = vbRed
            lblLabel6.Caption = "Time Exceeded by"
        End If
    
        If cRemain > -2 Then
            lblRemain.ForeColor = vbBlack
        Else
            lblRemain.ForeColor = vbRed
        End If
    End If

    '-----------------
    If cMCal3 > 999 Then 'txtMinute1 max length is 3 digits, or 999. prevents error created by 4 digit number.
        cmdAdditional.Enabled = False
'    ElseIf cMCal3 > 1 And cMCal3 <= 999 Then
'        cmdAdditional.Enabled = True
    End If

    '------------------------------------------------------------------
    lblTotal1.Caption = Format$(cMCal3, "#0") _
    & " min " & Format$(cMCal4, "00") & " sec"
    
    lblTotal2.Caption = "• " & Format(cHCal3, "#") & _
    cHCal5 & Format$(cHCal4, "#0.0") & " min, or" & vbCrLf & _
    "• " & Format$(cHours, "Standard") & " hours, or" & vbCrLf & _
    "• " & Format$(cSec, "#,###") & " seconds"

 End Sub

Private Sub chkSubtract_Click()

    cmdUndoClear.Caption = "&Undo Clear"
    
    If chkSubtract.Value = 1 Then 'SUBTRACT mode is selected
     
        txtMinute1.Enabled = False
        txtSecond1.Enabled = False
        txtMinute2.Enabled = False
        txtSecond2.Enabled = False
        txtMinute3.Enabled = False
        txtSecond3.Enabled = False
        txtMinute4.Enabled = False
        txtSecond4.Enabled = False
        txtMinute5.Enabled = False
        txtSecond5.Enabled = False
        txtMinute6.Enabled = False
        txtSecond6.Enabled = False
        txtMinute7.Enabled = False
        txtSecond7.Enabled = False
        txtMinute8.Enabled = False
        txtSecond8.Enabled = False
        txtMinute9.Enabled = False
        txtSecond9.Enabled = False

        lblRemain.Visible = False
        txtMinute1.BackColor = &HE0E0E0 'gray
        txtSecond1.BackColor = &HE0E0E0
        txtMinute2.BackColor = &HE0E0E0
        txtSecond2.BackColor = &HE0E0E0
        txtMinute3.BackColor = &HE0E0E0
        txtSecond3.BackColor = &HE0E0E0
        txtMinute4.BackColor = &HE0E0E0
        txtSecond4.BackColor = &HE0E0E0
        txtMinute5.BackColor = &HE0E0E0
        txtSecond5.BackColor = &HE0E0E0
        txtMinute6.BackColor = &HE0E0E0
        txtSecond6.BackColor = &HE0E0E0
        txtMinute7.BackColor = &HE0E0E0
        txtSecond7.BackColor = &HE0E0E0
        txtMinute8.BackColor = &HE0E0E0
        txtSecond8.BackColor = &HE0E0E0
        txtMinute9.BackColor = &HE0E0E0
        txtSecond9.BackColor = &HE0E0E0

        Line1(0).BorderWidth = 2
        Line1(3).Visible = False
        Line1(2).Visible = False
        
        lblLabelPlus.Visible = True
        lblLabelMinus.Visible = True
        lblLabel6.Visible = False
        cmdAdditional.Enabled = False
        
        lblTotal1 = ""
        lblTotal2.Caption = ""

        If cMCal3 <> 0 Or cMCal4 <> 0 Then
            txtMinute10 = Format$(cMCal3, "#0")
            txtSecond10 = Format$(cMCal4, "00")
            txtMinute11.SetFocus
        Else
            txtMinute10.SetFocus
        End If
        
        '----------
        cmdUndoClear.FontSize = 8
        cmdClearEntries.FontSize = 8
        cmdClearEntries.Caption = "&Clear Subtraction"
        cmdClearEntries.ToolTipText = "Click to clear denominator. Repeat to clear numerator."
        
 '-------------------------------------------------------
    ElseIf chkSubtract.Value = 0 Then 'subtract mode NOT selected
        
        Line1(0).BorderWidth = 1
        Line1(3).Visible = True
        Line1(2).Visible = True
        
        txtMinute1.Enabled = True
        txtSecond1.Enabled = True
        txtMinute2.Enabled = True
        txtSecond2.Enabled = True
        txtMinute3.Enabled = True
        txtSecond3.Enabled = True
        txtMinute4.Enabled = True
        txtSecond4.Enabled = True
        txtMinute5.Enabled = True
        txtSecond5.Enabled = True
        txtMinute6.Enabled = True
        txtSecond6.Enabled = True
        txtMinute7.Enabled = True
        txtSecond7.Enabled = True
        txtMinute8.Enabled = True
        txtSecond8.Enabled = True
        txtMinute9.Enabled = True
        txtSecond9.Enabled = True
        
        lblRemain.Visible = True
        
        txtMinute10 = ""
        txtSecond10 = ""
        txtMinute11 = ""
        txtSecond11 = ""
        
        lblLabelPlus.Visible = False
        lblLabelMinus.Visible = False
        lblLabel6.Visible = True
        cmdAdditional.Enabled = True
        
        txtMinute1.BackColor = &H80000005 'white
        txtSecond1.BackColor = &H80000005
        txtMinute2.BackColor = &H80000005
        txtSecond2.BackColor = &H80000005
        txtMinute3.BackColor = &H80000005
        txtSecond3.BackColor = &H80000005
        txtMinute4.BackColor = &H80000005
        txtSecond4.BackColor = &H80000005
        txtMinute5.BackColor = &H80000005
        txtSecond5.BackColor = &H80000005
        txtMinute6.BackColor = &H80000005
        txtSecond6.BackColor = &H80000005
        txtMinute7.BackColor = &H80000005
        txtSecond7.BackColor = &H80000005
        txtMinute8.BackColor = &H80000005
        txtSecond8.BackColor = &H80000005
        txtMinute9.BackColor = &H80000005
        txtSecond9.BackColor = &H80000005
          
        lblTotal1 = ""
        
        cmdUndoClear.FontSize = 9
        cmdClearEntries.FontSize = 9
        cmdClearEntries.Caption = "&Clear Entries"
        cmdClearEntries.ToolTipText = "Clears All 'Min' & 'Sec' Entries"
        If txtMinute1 <> "" Or txtSecond1 <> "" Or txtMinute2 <> "" Then
            pAdd
        End If
        pSetFocus
    End If
    iSubOption = 0
End Sub

Private Sub cmdAdditional_Click()
    txtMinute1 = Format$(cMCal3, "#0")
    txtSecond1 = Format$(cMCal4, "00")

    txtSecond2 = ""
    txtSecond3 = ""
    txtSecond4 = ""
    txtSecond5 = ""
    txtSecond6 = ""
    txtSecond7 = ""
    txtSecond8 = ""
    txtSecond9 = ""
    txtSecond10 = ""
    txtSecond11 = ""
 
    txtMinute2 = ""
    txtMinute3 = ""
    txtMinute4 = ""
    txtMinute5 = ""
    txtMinute6 = ""
    txtMinute7 = ""
    txtMinute8 = ""
    txtMinute9 = ""
    txtMinute10 = ""
    txtMinute11 = ""
    txtMinute2.SetFocus

End Sub

Private Sub cmdClearBlock_CliCk()
    txtBlock = ""
    lblRemain = ""
    lblLabel6 = ""
    txtBlock.SetFocus
End Sub

Private Sub cmdClearEntries_Click()

On Error GoTo HandleError

 '--------subtract mode---
    If chkSubtract.Value = 1 Then
        
        If iUndoClear = 0 Then
        
            If txtMinute10 <> "" Or txtMinute11 <> "" Then
            
                Open "Subtract.dat" For Output As #5 'for subtract mode
                    Write #5, txtMinute10, txtMinute11, txtSecond10, txtSecond11
                Close #5
            End If
        
            txtMinute11 = ""
            txtSecond11 = ""
            lblTotal1.Caption = ""
            iUndoClear = 1
            txtMinute11.SetFocus
            cmdClearEntries.Caption = "&Clear Numerator"
           
        ElseIf iUndoClear = 1 Then
            
            txtMinute10 = ""
            txtSecond10 = "" '
            lblTotal1.Caption = ""
            iUndoClear = 0
            txtMinute10.SetFocus
            cmdClearEntries.Caption = "&Clear Subtraction"
            Exit Sub
        End If
    
        Exit Sub
    End If
   
'-----------end subtract mode-----

If cmdUndoClear.Caption = "&Undo Clear" And _
    txtMinute1 = "" And txtMinute2 = "" And txtMinute3 = "" And txtMinute4 = "" And txtMinute5 = "" And _
    txtMinute6 = "" And txtMinute7 = "" And txtMinute8 = "" And txtMinute9 = "" And txtMinute10 = "" And txtMinute11 = "" And _
    txtSecond1 = "" And txtSecond2 = "" And txtSecond3 = "" And txtSecond4 = "" And txtSecond5 = "" And _
    txtSecond6 = "" And txtSecond7 = "" And txtSecond8 = "" And txtSecond9 = "" And txtSecond10 = "" And txtSecond11 = "" Then

    cmdUndoClear.Caption = "&Previous Entries"
    txtMinute1.SetFocus
    Exit Sub
End If
'-------
    cmdUndoClear.Caption = "&Undo Clear"
      
    lblRemain = ""
    lblLabel6 = ""
    txtSecond1 = ""
    txtSecond2 = ""
    txtSecond3 = ""
    txtSecond4 = ""
    txtSecond5 = ""
    txtSecond6 = ""
    txtSecond7 = ""
    txtSecond8 = ""
    txtSecond9 = ""
    txtSecond10 = ""
    txtSecond11 = ""
    
    txtMinute1 = ""
    txtMinute2 = ""
    txtMinute3 = ""
    txtMinute4 = ""
    txtMinute5 = ""
    txtMinute6 = ""
    txtMinute7 = ""
    txtMinute8 = ""
    txtMinute9 = ""
    txtMinute10 = ""
    txtMinute11 = ""
    lblTotal2 = ""
    lblTotal1 = ""
    cmdAdditional.Enabled = False
    txtMinute1.SetFocus
    frmTimeRemain!cmdAddTime.Caption = "AddTi&me F9"
    Exit Sub
    
HandleError:
    Close #5
End Sub

Private Sub cmdClose_Click()
    giTimeFocus = 1
    frmPlanner!cmdAddTime.Caption = " AddTime Calculator   F9"
    frmPlanner!cmdAddTime.ToolTipText = " Calculator for adding up times entered as minutes & seconds"
   
    frmTimeRemain!cmdAddTime.ToolTipText = " Calculator for adding up times entered as minutes & seconds"
    
    frmAddTime.Hide
    Unload frmAddHelp
 End Sub

Private Sub cmdUndoClear_Click()

On Error GoTo HandleError

'------subtract mode----
    If chkSubtract.Value = 1 Then
        iUndoClear = 0
    cmdClearEntries.Caption = "&Clear Denominator"
        
        Open "Subtract.dat" For Input As #5 'for subtract mode
        Input #5, iMin10, iMin11, iSec10, iSec11
        Close #5
        
         If cMCal3 <> 0 Or cMCal4 <> 0 Then
            txtMinute10 = Format$(cMCal3, "#0")
            txtSecond10 = Format$(cMCal4, "00")
            txtMinute11 = Format$(iMin11, "#0")
            txtSecond11 = Format$(iSec11, "00")
            txtMinute11.SetFocus
        Else
        
            txtMinute10 = Format$(iMin10, "#0")
            txtSecond10 = Format$(iSec10, "00")
            txtMinute11 = Format$(iMin11, "#0")
            txtSecond11 = Format$(iSec11, "00")
            txtMinute11.SetFocus
        End If
        
        Exit Sub
    Else
        cmdClearEntries.Caption = "&Clear Entries"
    End If
    
'----end subtract mode------

    Dim Minute1 As String
    Dim Minute2 As String
    Dim Minute3 As String
    Dim Minute4 As String
    Dim Minute5 As String
    Dim Minute6 As String
    Dim Minute7 As String
    Dim Minute8 As String
    Dim Minute9 As String
    Dim Minute10 As String
    Dim Minute11 As String
    
    Dim Second1 As String
    Dim Second2 As String
    Dim Second3 As String
    Dim Second4 As String
    Dim Second5 As String
    Dim Second6 As String
    Dim Second7 As String
    Dim Second8 As String
    Dim Second9 As String
    Dim Second10 As String
    Dim Second11 As String

    Dim Block As String
    
    If cmdUndoClear.Caption = "&Previous Entries" Then
        
        If giEntrySave = 0 Then
            giEntrySave = 1
        ElseIf giEntrySave = 1 Then
            giEntrySave = 2
        ElseIf giEntrySave = 2 Then
            giEntrySave = 0
        End If
                
    End If

    If giEntrySave = 0 Then
    
       Open "AddTime.dat" For Input As #3
           Input #3, Minute1, Minute2, Minute3, Minute4, Minute5, Minute6, Minute7, Minute8, Minute9, Minute10, Minute11, _
           Second1, Second2, Second3, Second4, Second5, Second6, Second7, Second8, Second9, Second10, Second11, Block
       Close #3
      
    ElseIf giEntrySave = 1 Then
    
        Open "AddTime2.dat" For Input As #4
           Input #4, Minute1, Minute2, Minute3, Minute4, Minute5, Minute6, Minute7, Minute8, Minute9, Minute10, Minute11, _
           Second1, Second2, Second3, Second4, Second5, Second6, Second7, Second8, Second9, Second10, Second11, Block
       Close #4
 
     ElseIf giEntrySave = 2 Then
     
        Open "AddTime3.dat" For Input As #6
           Input #6, Minute1, Minute2, Minute3, Minute4, Minute5, Minute6, Minute7, Minute8, Minute9, Minute10, Minute11, _
           Second1, Second2, Second3, Second4, Second5, Second6, Second7, Second8, Second9, Second10, Second11, Block
       Close #6
 
    End If
 
    cmdUndoClear.Caption = "&Previous Entries"
    
    txtMinute1 = Minute1
    txtMinute2 = Minute2
    txtMinute3 = Minute3
    txtMinute4 = Minute4
    txtMinute5 = Minute5
    txtMinute6 = Minute6
    txtMinute7 = Minute7
    txtMinute8 = Minute8
    txtMinute9 = Minute9
    txtMinute10 = Minute10
    txtMinute11 = Minute11
    
    txtSecond1 = Second1
    txtSecond2 = Second2
    txtSecond3 = Second3
    txtSecond4 = Second4
    txtSecond5 = Second5
    txtSecond6 = Second6
    txtSecond7 = Second7
    txtSecond8 = Second8
    txtSecond9 = Second9
    txtSecond10 = Second10
    txtSecond11 = Second11
    txtBlock = Block
    pSetFocus
    Exit Sub
    
HandleError:
    Close #3
    Close #4
    Close #6
    
    Open "Subtract.dat" For Output As #5 'for subtract mode
        Write #5, 10, 11, 10, 11
    Close #5
    
pSetFocus
End Sub

Private Sub Form_Activate()
    giTimeFocus = 3
    'To prevent Run-Time Error if Planner control box 'close' (iHourNow)
    'is clicked while AddTime is selected, StopWatch, AddTime, memos & PlanHelp
    'send giTimeFocus = 3 as a control number when any of the forms is activated.
    
'   If frmTimeRemain!fraAdjustTime.Visible = False Then
'        frmAddTime!lblLabelPlus.FontSize = 10
'        frmAddTime!lblLabelPlus.Caption = "+"
'        frmAddTime!lblLabelMinus.FontSize = 16
'        frmAddTime!lblLabelMinus.Caption = "-"
'        frmAddTime!chkSubtract.Caption = "Subtract Time"
'
'     ElseIf frmTimeRemain!fraAdjustTime.Visible = True Then
'        frmAddTime!lblLabelPlus.FontSize = 7
'        frmAddTime!lblLabelPlus.Caption = "actual time (min sec)"
'        frmAddTime!lblLabelMinus.FontSize = 7
'        frmAddTime!lblLabelMinus.Caption = "computer time (min sec)"
'        frmAddTime!chkSubtract.Caption = "clock error"
'    End If
    '--------
  
On Error GoTo HandleErrors
    Call pSetFocus
HandleErrors:
End Sub

Private Sub fraAddTime_Click()
    Call pSetFocus
End Sub

Private Sub lblLabelMinus_Click()
    txtMinute11 = ""
    txtSecond11 = ""
    txtMinute11.SetFocus
End Sub

Private Sub lblLabelPlus_Click()
    txtMinute10 = ""
    txtSecond10 = ""
    txtMinute10.SetFocus
End Sub

Private Sub lblTotal1_Change()

    If Val(lblTotal1) > 0 Then
    
        frmPlanner!fraFrame(0).ForeColor = &HC0&       'darker red '&HC00000   'Blue
        frmPlanner!fraFrame(0).FontSize = 10
        frmPlanner!fraFrame(0).Caption = lblTotal1 & " (AddTime Calculator)" 'displays AddTime total time
        frmTimeRemain!lblAddTime = lblTotal1 & "  AT calculator"  'displays AddTime total time
        frmPlanner!cmdAddTime.Caption = lblTotal1
        frmTimeRemain!cmdAddTime.Caption = lblTotal1
        Frame1.Caption = "Which is the same as"
        
    Else
        frmPlanner!fraFrame(0).ForeColor = &H80&       'rust
        frmPlanner!fraFrame(0).FontSize = 8
        frmPlanner!fraFrame(0).Caption = "Lineup && List Options"
        frmTimeRemain!lblAddTime = ""
        Frame1.Caption = ""

        frmPlanner!cmdAddTime.Caption = " AddTime Calculator   F9"
        frmPlanner!cmdAddTime.ToolTipText = " Calculator for adding up times entered as minutes & seconds"

        frmTimeRemain!cmdAddTime.Caption = "AddTi&me F9"
        frmTimeRemain!cmdAddTime.ToolTipText = " Calculator for adding up times entered as minutes & seconds"
    End If

End Sub

Private Sub mnuHelp_Click()
    frmAddHelp.Show vbModal
End Sub

Private Sub mnuPageCloseKeypad_Click()
    Call cmdClose_Click
End Sub

Private Sub mnuPagePlanner_Click()
    frmTransmitter.Hide
    frmTimeRemain.Hide
    frmPlanner.Show
End Sub

Private Sub mnuPagePrintPage_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "AddTime")
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

Private Sub mnuPageStopWatch_Click()
    frmStopWatch.Show
End Sub

Private Sub mnuPageTimeRemain_Click()
    frmPlanner.Hide
    frmTransmitter.Hide
    frmTimeRemain.Show
End Sub

Private Sub mnuPageXmitter_Click()
    frmPlanner.Hide
    frmTimeRemain.Hide
    Unload frmAddHelp
    frmAddTime.Hide
    
    frmTransmitter.Show
End Sub

Private Sub Form_Click()
    Call pSetFocus
End Sub

Private Sub fraFrame3_Click()
    Call pSetFocus
End Sub

Private Sub Label4_Click()
    Call pSetFocus
End Sub

Private Sub pSetFocus()
    'set focus to first blank minute entry box
    If chkSubtract.Value = 0 Then
        If txtMinute1 = "" Then
            txtMinute1.SetFocus
        ElseIf txtMinute2 = "" Then
            txtMinute2.SetFocus
        ElseIf txtMinute3 = "" Then
            txtMinute3.SetFocus
        ElseIf txtMinute4 = "" Then
            txtMinute4.SetFocus
        ElseIf txtMinute5 = "" Then
            txtMinute5.SetFocus
        ElseIf txtMinute6 = "" Then
            txtMinute6.SetFocus
        ElseIf txtMinute7 = "" Then
            txtMinute7.SetFocus
        ElseIf txtMinute10 = "" Then
            txtMinute10.SetFocus
        ElseIf txtMinute11 = "" Then
            txtMinute11.SetFocus
        End If
    End If
End Sub

Private Sub txtBlock_Change()
    'calls ADD feature after text block entry changes

    If IsNumeric(txtBlock) Or txtBlock = "" Then
       Call pAdd
    Else
       MsgBox "Entry Error" & vbCrLf & vbCrLf & _
       "Enter the number of minutes from which the time entered on the KeyPad will be subtracted." _
       & vbCrLf & vbCrLf & "If seconds are included, they should be entered as the decimal part of a minute. For example," & vbCrLf & _
       "56 minutes and 15 seconds should be entered as 56.25 minutes.", 0, "Display a Running Total of Time Remaining"
       txtBlock = ""
       txtBlock.SetFocus
    End If
  
    If txtBlock = "" Then
        lblRemain = ""
        cmdClearBlock.Caption = "Set &Block Time"
    ElseIf txtBlock <> "" Then
        cmdClearBlock.Caption = "Clear &Block"
    End If
    
    If txtBlock = "" Then
        lblRemain.BackColor = &HE0E0E0
    Else
        lblRemain.BackColor = vbWhite
    End If

 End Sub

Private Sub txtBlock_LostFocus()
   If txtBlock <> "" Then
        txtBlock.Text = Format$(txtBlock, "#0.0#")
    End If
 End Sub

Private Sub txtMinute1_Change()
    If txtMinute1 <> "" Then
        cmdUndoClear.Caption = "&Previous Entries"
    End If
    
    'calls ADD feature after each minute text box change
    Call pAdd
End Sub

Private Sub txtMinute1_GotFocus()
    txtMinute1.SelStart = 0 'begin selection at start
    txtMinute1.SelLength = Len(txtMinute1)
End Sub

Private Sub txtMinute1_LostFocus()
    If txtMinute1.Text = "" And txtSecond1 = "00" Then
        txtSecond1 = ""
    End If
    
    If txtMinute1 <> "" Then
    
        If giEntrySave = 0 Then 'directs new entry to be saved in different file
            giEntrySave = 1
        ElseIf giEntrySave = 1 Then
            giEntrySave = 2
        ElseIf giEntrySave = 2 Then
            giEntrySave = 0
        End If
    
        cmdUndoClear.Caption = "&Previous Entries"
        Call pSave
    End If
    
End Sub

Private Sub txtMinute2_Change()
    Call pAdd
End Sub

Private Sub txtMinute2_GotFocus()
    txtMinute2.SelStart = 0 'begin selection at start
    txtMinute2.SelLength = Len(txtMinute2)
End Sub

Private Sub txtMinute2_LostFocus()
    If txtMinute2.Text = "" And txtSecond2 = "00" Then
        txtSecond2 = ""
    End If
    
    If Val(txtMinute2) < 0 Then
        txtSecond2 = "-"
    Else
        txtSecond2 = ""
    End If
    
    If txtMinute2 <> "" Then
        Call pSave
    End If

End Sub

Private Sub txtMinute3_Change()
    If txtMinute3 <> "" Then
        cmdAdditional.Enabled = True
    End If
    Call pAdd
End Sub

Private Sub txtMinute3_GotFocus()
    txtMinute3.SelStart = 0 'begin selection at start
    txtMinute3.SelLength = Len(txtMinute3)
End Sub

Private Sub txtMinute3_LostFocus()
    If txtMinute3.Text = "" And txtSecond3 = "00" Then
        txtSecond3 = ""
    End If
    
    If txtMinute3 <> "" Then
        Call pSave
    End If

End Sub

Private Sub txtMinute4_Change()
    Call pAdd
End Sub

Private Sub txtMinute4_GotFocus()
    txtMinute4.SelStart = 0 'begin selection at start
    txtMinute4.SelLength = Len(txtMinute4)
End Sub

Private Sub txtMinute4_LostFocus()
    If txtMinute4.Text = "" And txtSecond4 = "00" Then
        txtSecond4 = ""
    End If
    
    If txtMinute4 <> "" Then
        Call pSave
    End If

End Sub

Private Sub txtMinute5_Change()
    Call pAdd
End Sub

Private Sub txtMinute5_GotFocus()
    txtMinute5.SelStart = 0 'begin selection at start
    txtMinute5.SelLength = Len(txtMinute5)
End Sub

Private Sub txtMinute5_LostFocus()
    If txtMinute5.Text = "" And txtSecond5 = "00" Then
        txtSecond5 = ""
    End If
    
    If txtMinute5 <> "" Then
        Call pSave
    End If
    
End Sub

Private Sub txtMinute6_Change()
    If txtMinute6 <> "" Then
        cmdAdditional.Enabled = True
    End If
    Call pAdd
End Sub

Private Sub txtMinute6_GotFocus()
    txtMinute6.SelStart = 0 'begin selection at start
    txtMinute6.SelLength = Len(txtMinute6)
End Sub

Private Sub txtMinute6_LostFocus()
    If txtMinute6.Text = "" And txtSecond6 = "00" Then
        txtSecond6 = ""
    End If
    
    If txtMinute6 <> "" Then
        Call pSave
    End If
    
End Sub

Private Sub txtMinute7_Change()
    Call pAdd
End Sub

Private Sub txtMinute7_GotFocus()
    txtMinute7.SelStart = 0 'begin selection at start
    txtMinute7.SelLength = Len(txtMinute7)
End Sub

Private Sub txtMinute7_LostFocus()
    If txtMinute7.Text = "" And txtSecond7 = "00" Then
        txtSecond7 = ""
    End If
    
    If txtMinute7 <> "" Then
        Call pSave
    End If
    
End Sub

Private Sub txtMinute10_Change()
    Call pAdd
End Sub

Private Sub txtMinute10_GotFocus()
    txtMinute10.SelStart = 0 'begin selection at start
    txtMinute10.SelLength = Len(txtMinute10)
    iFocus8 = 1
End Sub

Private Sub txtMinute10_LostFocus()
    If txtMinute10.Text = "" And txtSecond10 = "00" Then
        txtSecond10 = ""
    End If
    iFocus8 = 0
End Sub

Private Sub txtMinute11_Change()
    Call pAdd
End Sub

Private Sub txtMinute11_GotFocus()
    txtMinute11.SelStart = 0 'begin selection at start
    txtMinute11.SelLength = Len(txtMinute11)
    iFocus8 = 1
End Sub

Private Sub txtMinute11_LostFocus()
    If txtMinute11.Text = "" And txtSecond11 = "00" Then
        txtSecond11 = ""
    End If
    iFocus8 = 0
End Sub

Private Sub txtMinute8_Change()
    Call pAdd
End Sub

Private Sub txtMinute8_GotFocus()
    txtMinute8.SelStart = 0 'begin selection at start
    txtMinute8.SelLength = Len(txtMinute8)
End Sub

Private Sub txtMinute8_LostFocus()
    If txtMinute8.Text = "" And txtSecond8 = "00" Then
        txtSecond8 = ""
    End If
    
    If txtMinute8 <> "" Then
        Call pSave
    End If
    
End Sub

Private Sub txtMinute9_Change()
    If txtMinute9 <> "" Then
        cmdAdditional.Enabled = True
    End If
    Call pAdd
End Sub

Private Sub txtMinute9_GotFocus()
    txtMinute9.SelStart = 0 'begin selection at start
    txtMinute9.SelLength = Len(txtMinute9)
End Sub

Private Sub txtMinute9_LostFocus()
    If txtMinute9.Text = "" And txtSecond9 = "00" Then
        txtSecond9 = ""
    End If
    
    If txtMinute9 <> "" Then
        Call pSave
    End If
    
End Sub

Private Sub txtSecond1_Change()
    Call pAdd
End Sub

Private Sub txtSecond1_GotFocus()
    txtSecond1.SelStart = 0 'begin selection at start
    txtSecond1.SelLength = Len(txtSecond1)
End Sub

Private Sub txtSecond1_LostFocus()
    If txtMinute1.Text <> "" And txtSecond1 = "" Then
        txtSecond1 = "00"
    End If
    txtSecond1.Text = Format$(txtSecond1, "00")
    
    
    If txtSecond1 <> "" Then
        Call pSave
    End If

End Sub

Private Sub txtSecond2_Change()
    Call pAdd
End Sub

Private Sub txtSecond2_GotFocus()

    If Val(txtMinute2) >= 0 Then 'begin selection at start
        txtSecond2.SelStart = 0 'begin selection at start
        txtSecond2.SelLength = Len(txtSecond2)
    Else
        txtSecond2.SelStart = 1
    End If
End Sub

Private Sub txtSecond2_LostFocus()

    If Val(txtMinute2) < 0 And Val(txtSecond2) > 0 Then
        txtSecond2 = ""
        Exit Sub
    End If

    If Val(txtMinute2) > 0 And (txtMinute2.Text <> "" And txtSecond2 = "") Then
        txtSecond2 = "00"
    End If
    txtSecond2.Text = Format$(txtSecond2, "00")
    
    If txtSecond2 <> "" Then
        Call pSave
    End If

End Sub

Private Sub txtSecond3_Change()
    If txtSecond3 <> "" Then
        cmdAdditional.Enabled = True
    End If
    Call pAdd
End Sub

Private Sub txtSecond3_GotFocus()
    txtSecond3.SelStart = 0 'begin selection at start
    txtSecond3.SelLength = Len(txtSecond3)
End Sub

Private Sub txtSecond3_LostFocus()
    If txtMinute3.Text <> "" And txtSecond3 = "" Then
        txtSecond3 = "00"
    End If
    txtSecond3.Text = Format$(txtSecond3, "00")
    
    If txtSecond3 <> "" Then
        Call pSave
    End If

End Sub

Private Sub txtSecond4_Change()
    Call pAdd
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
    
    If txtSecond4 <> "" Then
        Call pSave
    End If
    
End Sub

Private Sub txtSecond5_Change()
    Call pAdd
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
    
    If txtSecond5 <> "" Then
        Call pSave
    End If
    
End Sub

Private Sub txtSecond6_Change()
    Call pAdd
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
    
    If txtSecond6 <> "" Then
        Call pSave
    End If

End Sub

Private Sub txtSecond7_Change()
    Call pAdd
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
    
    If txtSecond7 <> "" Then
        Call pSave
    End If

End Sub

Private Sub txtSecond10_Change()
    Call pAdd
End Sub

Private Sub txtSecond10_GotFocus()
    txtSecond10.SelStart = 0 'begin selection at start
    txtSecond10.SelLength = Len(txtSecond10)
    iFocus8 = 1
End Sub

Private Sub txtSecond10_LostFocus()
    If txtMinute10.Text <> "" And txtSecond10 = "" Then
        txtSecond10 = "00"
    End If
    txtSecond10.Text = Format$(txtSecond10, "00")
    iFocus8 = 0
End Sub

Private Sub txtSecond11_Change()
    Call pAdd
End Sub

Private Sub txtSecond11_GotFocus()
    txtSecond11.SelStart = 0 'begin selection at start
    txtSecond11.SelLength = Len(txtSecond11)
    iFocus8 = 1
End Sub

Private Sub txtSecond11_LostFocus()
    If txtMinute11.Text <> "" And txtSecond11 = "" Then
        txtSecond11 = "00"
    End If
    txtSecond11.Text = Format$(txtSecond11, "00")
    
    If chkSubtract.Value = 1 And iSubOption = 1 Then
    cmdClearEntries.SetFocus
    End If
    
    iFocus8 = 0
    If txtSecond11 <> "" And chkSubtract.Value = 1 Then
        Call pAdd
    End If

End Sub

Private Sub txtSecond8_Change()
    Call pAdd
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
    
    If txtSecond8 <> "" Then
        Call pSave
    End If
    
End Sub

Private Sub txtSecond9_Change()
    If txtSecond9 <> "" Then
        cmdAdditional.Enabled = True
    End If
    Call pAdd
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
    
    If txtSecond9 <> "" Then
        Call pSave
    End If

End Sub

Private Sub pSave()

On Error GoTo HandleError

    If giEntrySave = 0 Then
    
        Open "AddTime.dat" For Output As #3
        Write #3, txtMinute1, txtMinute2, txtMinute3, txtMinute4, txtMinute5, txtMinute6, txtMinute7, txtMinute8, txtMinute9, txtMinute10, txtMinute11, _
        txtSecond1, txtSecond2, txtSecond3, txtSecond4, txtSecond5, txtSecond6, txtSecond7, txtSecond8, txtSecond9, txtSecond10, txtSecond11, txtBlock
        Close #3
    
    ElseIf giEntrySave = 1 Then
    
        Open "AddTime2.dat" For Output As #4
        Write #4, txtMinute1, txtMinute2, txtMinute3, txtMinute4, txtMinute5, txtMinute6, txtMinute7, txtMinute8, txtMinute9, txtMinute10, txtMinute11, _
        txtSecond1, txtSecond2, txtSecond3, txtSecond4, txtSecond5, txtSecond6, txtSecond7, txtSecond8, txtSecond9, txtSecond10, txtSecond11, txtBlock
        Close #4
        
     ElseIf giEntrySave = 2 Then
     
        Open "AddTime3.dat" For Output As #6
        Write #6, txtMinute1, txtMinute2, txtMinute3, txtMinute4, txtMinute5, txtMinute6, txtMinute7, txtMinute8, txtMinute9, txtMinute10, txtMinute11, _
        txtSecond1, txtSecond2, txtSecond3, txtSecond4, txtSecond5, txtSecond6, txtSecond7, txtSecond8, txtSecond9, txtSecond10, txtSecond11, txtBlock
        Close #6
    End If
    Exit Sub
    
HandleError:

    Close #3
    Close #4
    Close #6
End Sub
