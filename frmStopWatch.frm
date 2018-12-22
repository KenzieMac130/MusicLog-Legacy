VERSION 5.00
Begin VB.Form frmStopWatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StopWatch - Timer   F12"
   ClientHeight    =   5115
   ClientLeft      =   10470
   ClientTop       =   2220
   ClientWidth     =   3615
   ControlBox      =   0   'False
   Icon            =   "frmStopWatch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   3615
   Begin VB.Timer timDisplay 
      Interval        =   1000
      Left            =   165
      Top             =   4725
   End
   Begin VB.Frame Frame2 
      Height          =   930
      Left            =   240
      TabIndex        =   11
      Top             =   3795
      Width           =   3165
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close  F12"
         Height          =   375
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   345
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   390
         Left            =   1845
         TabIndex        =   13
         Top             =   135
         Width           =   1170
         Begin VB.Label lblClock 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "00:00:00 PM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   135
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdNote 
         Caption         =   "Read &Me"
         Height          =   270
         Left            =   2003
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   570
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2775
      Left            =   225
      TabIndex        =   4
      Top             =   -60
      Width           =   3165
      Begin VB.CheckBox chkLockReset 
         Caption         =   "&Lock-Out Reset Command"
         Height          =   210
         Left            =   510
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2460
         Width           =   2235
      End
      Begin VB.CommandButton cmdStop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Height          =   360
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   840
         Width           =   1380
      End
      Begin VB.Frame frmCounter 
         ForeColor       =   &H000000FF&
         Height          =   720
         Left            =   1635
         TabIndex        =   24
         Top             =   675
         Width           =   1425
         Begin VB.Label lblRunTime 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmdResume 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Resume "
         Enabled         =   0   'False
         Height          =   360
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1380
         Width           =   1380
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Start Timing"
         Height          =   360
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reset"
         Enabled         =   0   'False
         Height          =   330
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2280
      End
      Begin VB.Label lblHr 
         Caption         =   "hr"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1875
         TabIndex        =   23
         Top             =   1635
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblRunTime2 
         Alignment       =   2  'Center
         Caption         =   "00   00 min 00 sec"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1560
         TabIndex        =   21
         Top             =   1635
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Elapsed Time"
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
         Height          =   240
         Left            =   1830
         TabIndex        =   10
         Top             =   1410
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lblStart 
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
         Height          =   255
         Left            =   2220
         TabIndex        =   9
         Top             =   165
         Width           =   765
      End
      Begin VB.Label lblEnd 
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
         Height          =   255
         Left            =   2235
         TabIndex        =   8
         Top             =   375
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdTime1 
      Caption         =   "Split &1"
      Enabled         =   0   'False
      Height          =   270
      Left            =   270
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2850
      Width           =   645
   End
   Begin VB.CommandButton cmdTime2 
      Caption         =   "Split &2"
      Enabled         =   0   'False
      Height          =   270
      Left            =   1095
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2850
      Width           =   645
   End
   Begin VB.CommandButton cmdTime3 
      Caption         =   "Split &3"
      Enabled         =   0   'False
      Height          =   270
      Left            =   1905
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2850
      Width           =   645
   End
   Begin VB.CommandButton cmdTime4 
      Caption         =   "Split &4"
      Enabled         =   0   'False
      Height          =   270
      Left            =   2730
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2850
      Width           =   645
   End
   Begin VB.Label lblTime4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2670
      MouseIcon       =   "frmStopWatch.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   3195
      Width           =   735
   End
   Begin VB.Label lblTime3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1860
      MouseIcon       =   "frmStopWatch.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   3195
      Width           =   735
   End
   Begin VB.Label lblTime2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1035
      MouseIcon       =   "frmStopWatch.frx":0A56
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3195
      Width           =   735
   End
   Begin VB.Label lblTime1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   225
      MouseIcon       =   "frmStopWatch.frx":0D60
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   3195
      Width           =   735
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
      Left            =   1170
      TabIndex        =   16
      Top             =   4785
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Double-Click Split Text Box to Clear"
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
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   390
      TabIndex        =   15
      ToolTipText     =   "Double-Clicking this line of text will clear all Split text boxes."
      Top             =   3495
      Width           =   2865
   End
   Begin VB.Menu mnuPage 
      Caption         =   "P&age"
      Begin VB.Menu mnuPagePlanner 
         Caption         =   "&Music Planner"
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
      Begin VB.Menu mnuPageSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageAddTime 
         Caption         =   "AddTime &Keypad..."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuPageSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageClose 
         Caption         =   "&Close StopWatch"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPageSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPagePrintPage 
         Caption         =   "&Print a Copy of this Page"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsBeeps 
         Caption         =   "&Beep to Indicate Start, Stop and Resume Timing"
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "frmStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'StopWatch-Timer Revised
'Jim Wright 11/2000

    Dim mStartTime As Variant
    Dim mEndTime As Variant
    Dim mElapsedTime As Variant
    Dim Today As Variant
    Dim mControl As Integer

Option Explicit

Private Sub chkLockReset_Click()

    If chkLockReset.Value = 0 Then
    cmdReset.Caption = "Reset"
        If cmdResume.Enabled = True Then
        cmdReset.Caption = "Reset Elapsed Time to &Zero"
        cmdReset.Enabled = True
        End If
    Else
    cmdReset.Caption = "Reset Locked Out"
    cmdReset.Enabled = False
    End If

End Sub

Private Sub cmdNote_Click()
    frmNote.Show vbModal
    cmdNote.Enabled = False
    cmdClose.Enabled = False
End Sub

Private Sub cmdTime1_Click()
    lblTime1.Caption = Format(mRunTime, "hh:mm:ss")
    cmdStop.SetFocus
    Label5.Enabled = True
End Sub

Private Sub cmdResume_Click()

On Error GoTo HandleErrors
   mControl = 1
       Open "RunTimeSW.dat" For Input As #1
       Input #1, mElapsedTime
    Close #1
    
    If mnuToolsBeeps.Checked = True Then
        Beep
    End If
    
    mStartTime = Now - mElapsedTime
    lblStart.Caption = Format(mStartTime, "hh:mm:ss")
   
    lblEnd.Caption = ""
    frmCounter.ForeColor = &HFF0000
    frmCounter.Caption = "Running"
    Label4.Visible = False
    cmdStop.Enabled = True
    cmdStop.Caption = "&Stop Timing"
    cmdStop.BackColor = &HFFFFFF
    cmdReset.Enabled = False
    cmdReset.Caption = "Reset"
    cmdReset.BackColor = &HC0C0C0
    cmdResume.Caption = "Resume"
    cmdResume.BackColor = &HC0C0C0
    cmdResume.Enabled = False
    cmdStop.SetFocus
    
    lblRunTime2.Visible = False
    lblHr.Visible = False
    cmdTime1.Enabled = True
    cmdTime2.Enabled = True
    cmdTime3.Enabled = True
    cmdTime4.Enabled = True
    frmPlanner!lblTimerLabel.Caption = "Timer Running"
    frmTimeRemain!Frame9.Caption = "Timer Running"
    frmTransmitter!lblTimerLabel.Caption = "Timer Running"
HandleErrors:
End Sub

Private Sub cmdStop_Click()

    'find the ending time, compute the elapsed time,
    mControl = 0
    
    If mnuToolsBeeps.Checked = True Then
        Beep
    End If
   
    mEndTime = Now
    mElapsedTime = mEndTime - mStartTime
    lblRunTime.Caption = Format(mElapsedTime, "hh:mm:ss")

    Open "RunTimeSW.dat" For Output As #1
        Write #1, mElapsedTime
    Close #1
    
    lblEnd.Caption = Format(mEndTime, "hh:mm:ss")
   
    frmCounter.ForeColor = &HFF&
    frmCounter.Caption = "Stopped"
    Label4.Visible = True
    cmdStop.Caption = "Stop"
    cmdStop.BackColor = &HC0C0C0
    
    If chkLockReset.Value = 0 Then
    cmdReset.Enabled = True
    cmdReset.Caption = "Reset ElapsedTime to &Zero"
    Else
     cmdReset.Caption = "Reset Locked Out"
    End If
    cmdReset.BackColor = &HFFFFFF
    cmdResume.Enabled = True
    cmdResume.Caption = "&Resume Timing"
    cmdResume.BackColor = &HFFFFFF
    cmdResume.SetFocus
     cmdStop.Enabled = False
    cmdTime1.Enabled = False
    cmdTime2.Enabled = False
    cmdTime3.Enabled = False
    cmdTime4.Enabled = False
    lblRunTime2.Visible = True
    lblHr.Visible = True
    If lblTime1 = "" And lblTime2 = "" And lblTime3 = "" And lblTime4 = "" Then
        Label5.Enabled = False
    End If
    frmPlanner!lblTimerLabel.Caption = "Timer Paused"
    frmTimeRemain!Frame9.Caption = "Timer Paused"
    frmTransmitter!lblTimerLabel.Caption = "Timer Paused"
End Sub

Private Sub cmdClose_Click()
    frmTimeRemain!cmdStopwatch.Enabled = True
    Unload frmNote
    frmStopWatch.Hide
    
    If cmdClose.Caption = "&Close  F12" Then
        frmTimeRemain!chkStopWatch.Value = 0
    End If
    
End Sub

Private Sub cmdReset_Click()

    Dim iResponse As Integer

    iResponse = MsgBox("Do you want reset elapsed time to zero?", vbYesNo + vbQuestion, "Reset to Zero")

    If iResponse = vbNo Then '
        Exit Sub
    Else
    End If

    If chkLockReset.Value = 0 Then
        mStartTime = 0
        lblRunTime = "00:00:00"
        lblRunTime2.Visible = False
        lblHr.Visible = False
        lblStart.Caption = ""
        lblEnd.Caption = ""
        
        cmdStart.Enabled = True
        cmdStart.Caption = "&Start Timing"
        cmdStart.BackColor = &HFFFFFF
        cmdStart.SetFocus
        Label4.Visible = False
        frmCounter.Caption = ""
         
        cmdStop.Caption = "Stop"
        cmdStop.Enabled = False
        cmdResume.Caption = "Resume"
        cmdResume.BackColor = &HC0C0C0
        cmdResume.Enabled = False
        cmdReset.Caption = "Reset"
        cmdReset.BackColor = &HC0C0C0
        cmdReset.Enabled = False
        cmdTime1.Enabled = False
        cmdTime2.Enabled = False
        cmdTime3.Enabled = False
        cmdTime4.Enabled = False
        cmdClose.Caption = "&Close  F12"
        Label5.Enabled = True
        mRunTime = 0
    Else
    End If
    
End Sub

Private Sub cmdStart_Click()
    mControl = 1
    
    If mnuToolsBeeps.Checked = True Then
        Beep
    End If
   
    mStartTime = Time
    lblStart.Caption = Format(mStartTime, "hh:mm:ss")
    
    lblEnd.Caption = ""
    Label4.Visible = False
    frmCounter.ForeColor = &HFF0000
    frmCounter.Caption = "Running"
    cmdStop.Enabled = True
    cmdStop.Caption = "&Stop Timing"
    cmdStop.BackColor = &HFFFFFF   '&HE0E0E0
    cmdResume.Enabled = False
    cmdStart.Caption = "Start"
    cmdStart.BackColor = &HC0C0C0
    cmdStart.Enabled = False
    cmdStop.SetFocus
    
    lblRunTime2.Visible = False
    lblHr.Visible = False
    cmdTime1.Enabled = True
    cmdTime2.Enabled = True
    cmdTime3.Enabled = True
    cmdTime4.Enabled = True
    
    cmdClose.Caption = "&Hide  F12"
    frmPlanner!lblTimerLabel.Caption = "Timer Running" '•
    frmTimeRemain!Frame9.ForeColor = vbBlue
    frmTimeRemain!Frame9.Caption = "Timer Running"
    frmTransmitter!lblTimerLabel.Caption = "»Timer Running"
End Sub

Private Sub mnuReturn_Click()
    Call cmdClose_Click
End Sub

Private Sub cmdTime2_Click()
    lblTime2.Caption = Format(mRunTime, "hh:mm:ss")
    cmdStop.SetFocus
    Label5.Enabled = True
End Sub

Private Sub cmdTime3_Click()
    lblTime3.Caption = Format(mRunTime, "hh:mm:ss")
    Label5.Enabled = True
    cmdStop.SetFocus
End Sub

Private Sub cmdTime4_Click()
    lblTime4.Caption = Format(mRunTime, "hh:mm:ss")
    cmdStop.SetFocus
    Label5.Enabled = True
End Sub

Private Sub Form_Activate()
    giTimeFocus = 3
'To prevent Run-Time Error if Planner control box 'close' (iHourNow)
'is clicked while stopwatch is selected, StopWatch, AddTime, memos & PlanHelp
'send giTimeFocus = 3 as a control number when any of the forms is activated.

    cmdNote.Enabled = True
    cmdClose.Enabled = True
End Sub

Private Sub Form_Load()
    lblRunTime.Caption = "00:00:00"
End Sub
Private Sub Label5_DblClick()
    lblTime1 = ""
    lblTime2 = ""
    lblTime3 = ""
    lblTime4 = ""
End Sub

Private Sub lblTime1_DblClick()
    lblTime1 = ""
End Sub
Private Sub lblTime2_DblClick()
    lblTime2 = ""
End Sub

Private Sub lblTime3_DblClick()
    lblTime3 = ""
End Sub

Private Sub lblTime4_DblClick()
    lblTime4 = ""
End Sub

Private Sub mnuPageAddTime_Click()
    frmAddTime.Show
End Sub

Private Sub mnuPageClose_Click()
    cmdClose_Click
End Sub

Private Sub mnuPagePlanner_Click()
    Unload frmNote
    frmTransmitter.Hide
    frmTimeRemain.Hide
    frmPlanner.Show
End Sub

Private Sub mnuPagePrintPage_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "Timer")
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

Private Sub mnuPageTimeRemain_Click()
    Unload frmNote
    frmPlanner.Hide
    frmTransmitter.Hide
    frmTimeRemain.Show
End Sub

Private Sub mnuPageXmitter_Click()
    Unload frmNote
    frmPlanner.Hide
    frmTimeRemain.Hide
    frmTransmitter.Show
End Sub

Private Sub mnuToolsBeeps_Click()
    If mnuToolsBeeps.Checked = True Then
        mnuToolsBeeps.Checked = False
    Else
        mnuToolsBeeps.Checked = True
    End If
End Sub

Private Sub timDisplay_Timer()
    Dim Today As Variant
    Today = Now

   lblClock.Caption = Format(Today, "h:mm:ss ampm")
   
'--test to adjust time---
'lblClock.Caption = Time '24 hour time format
'lblClock.Caption = Hour(Time) & ":" & Minute(Time) + 5 & ":" & Second(Time) + 3
'---end test----

    If lblStart = "" Then
        lblRunTime = "00:00:00"
    
        frmPlanner!lblTimer = ""
        frmPlanner!lblTimerLabel.Visible = False
        frmTimeRemain!lblTimer = ""
        frmTimeRemain!Frame9.ForeColor = &H80&
        frmTimeRemain!Frame9.Caption = "Stopwatch: Alternate Source for Program Timing"
       ' frmPlanner!Label4.Visible = True
        frmTransmitter!lblTimer = ""
        frmTransmitter!lblTimerLabel.Visible = False

    Else
        If mControl = 1 Then
            mRunTime = Today - mStartTime
        ElseIf mControl = 0 Then
            mRunTime = mElapsedTime
        End If
        lblRunTime = Format(mRunTime, "hh:mm:ss")
        frmPlanner!lblTimer = Format(mRunTime, "hh:mm:ss")
        frmPlanner!lblTimerLabel.Visible = True
       ' frmPlanner!Label4.Visible = False
        frmTimeRemain!lblTimer = Format(mRunTime, "hh:mm:ss")
        frmTransmitter!lblTimer = Format(mRunTime, "hh:mm:ss")
        frmTransmitter!lblTimerLabel.Visible = True
        lblRunTime2 = Format(mRunTime, "hh     mm") & "min  " & Format(mRunTime, "ss") & "sec"
    End If

End Sub

