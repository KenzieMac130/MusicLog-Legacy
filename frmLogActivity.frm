VERSION 5.00
Begin VB.Form frmLogActivity 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "MusicLog Activity Log"
   ClientHeight    =   11115
   ClientLeft      =   2445
   ClientTop       =   405
   ClientWidth     =   7515
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogActivity.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10425
      Left            =   105
      TabIndex        =   0
      Top             =   585
      Width           =   7290
      Begin VB.ListBox lstActivity 
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
         Height          =   3375
         Left            =   1005
         TabIndex        =   20
         Top             =   2070
         Width           =   5490
      End
      Begin VB.CommandButton cmdClearLog 
         BackColor       =   &H80000018&
         Caption         =   "&Clear Activity Log"
         Height          =   460
         Left            =   5445
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1410
         Width           =   1700
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H80000018&
         Cancel          =   -1  'True
         Caption         =   "E&xit Page"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   485
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   7935
         Width           =   1140
      End
      Begin VB.ListBox lstAccess 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   1905
         TabIndex        =   1
         Top             =   7935
         Width           =   3585
      End
      Begin VB.Image imgHand 
         Height          =   480
         Left            =   6600
         Picture         =   "frmLogActivity.frx":0442
         Top             =   1965
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Caption         =   "25"
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
         Left            =   675
         TabIndex        =   23
         Top             =   9015
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   480
         Picture         =   "frmLogActivity.frx":0884
         Stretch         =   -1  'True
         Top             =   3525
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   427
         TabIndex        =   22
         Top             =   3105
         Width           =   420
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000000&
         Caption         =   $"frmLogActivity.frx":0CC6
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
         Height          =   720
         Left            =   1410
         TabIndex        =   21
         Top             =   1290
         Width           =   3810
      End
      Begin VB.Label lblLast 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblLast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   7560
         Width           =   450
      End
      Begin VB.Label lblNextToLast 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblNextToLast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   16
         Top             =   7275
         Width           =   975
      End
      Begin VB.Label lblSecondLast 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblSecondLast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   15
         Top             =   6975
         Width           =   1005
      End
      Begin VB.Label lblAccessOpened 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Activity Log has been opened 25 times."
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
         Height          =   255
         Left            =   1905
         TabIndex        =   14
         Top             =   9570
         Width           =   3435
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         Caption         =   "Dates && number of times this activity log has been viewed"
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
         Height          =   870
         Left            =   450
         TabIndex        =   13
         Top             =   8115
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000000&
         Caption         =   "Dates of recent previous MusicLog activity"
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
         Height          =   1140
         Left            =   450
         TabIndex        =   12
         Top             =   6015
         Width           =   795
      End
      Begin VB.Label lblThirdLast 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblThirdLast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   6690
         Width           =   810
      End
      Begin VB.Label lblFourthLast 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblFourthLast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   6405
         Width           =   900
      End
      Begin VB.Label lblFifthLast 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblFifthLast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   9
         Top             =   6120
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "Log is cleared and count reset to Ø after 25 entries."
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
         Height          =   255
         Left            =   1710
         TabIndex        =   8
         Top             =   9825
         Width           =   3810
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblCount"
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
         Height          =   285
         Left            =   3315
         TabIndex        =   7
         Top             =   620
         Width           =   720
      End
      Begin VB.Label lblToday 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblToday"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3240
         TabIndex        =   6
         Top             =   285
         Width           =   870
      End
      Begin VB.Label lblPath 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblPath"
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
         Height          =   255
         Left            =   3375
         TabIndex        =   5
         Top             =   10065
         Width           =   480
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblLogActivity"
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
         Height          =   285
         Left            =   3105
         TabIndex        =   4
         Top             =   910
         Width           =   1140
      End
      Begin VB.Label lblSixthLast 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblSixthLast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   5820
         Width           =   795
      End
      Begin VB.Label lblSeventhLast 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "lblSeventhLast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   5535
         Width           =   1050
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000002&
      Caption         =   "Activity Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   240
      TabIndex        =   24
      Top             =   105
      Width           =   1395
   End
End
Attribute VB_Name = "frmLogActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmLogActivity

Dim iCount As Integer
Option Explicit

Private Sub cmdClearLog_Click()
    Dim sTemp As String
    Dim daDate As String
    Dim iResponse As Integer
    
 'manually clear activity screen
 On Error GoTo HandleErrors
 
    iResponse = MsgBox("Do you want to clear the Activity Log and return the activity count to zero at this time?", vbYesNo, "Clear Activity Log?")
    
    If iResponse = vbNo Then
        cmdClose.SetFocus
        Exit Sub
    
    ElseIf iResponse = vbYes Then
        lstActivity.Clear
        cmdClearLog.Enabled = False
        
        daDate = Format(Date, "long Date")
        
        Open "A_LogLog.txt" For Output As #102 '1
        Print #102, ">Begin", daDate
        Close #102
 
        Open "A_LogLog.txt" For Input As #102 'opens text file & restores data to screen
        Do Until EOF(102)
        Line Input #102, sTemp
        lstActivity.AddItem sTemp
        Loop
        Close #102

        lblCount.Caption = "0  MusicLog open/close cycles begin with the entry dated"
    End If
    cmdClearLog.BackColor = &HFFFFFF
    imgHand.Visible = False
    cmdClose.SetFocus
HandleErrors:
End Sub

Private Sub cmdClose_Click()

    If lstActivity.ListCount >= 398 Then  '>= 199 open/close cycles

         MsgBox "The Activity Log has logged 199 or more open/close cycles." & vbCrLf & vbCrLf & _
        "The log will be cleared and begin a new series of activity counts.", vbOKOnly, "Clearing Activity Log"
        
        lstActivity.Clear
        
        Dim daDate As String
        daDate = Format(Date, "long Date")

        Open "A_LogLog.txt" For Output As #102 '2
        Print #102, ">Begin", daDate
        Close #102
    End If

    giActivity = 0
    frmTransmitter!cmdRestoreDefaults.Caption = "Restore &Default Entries"
    frmTransmitter!cmdRestoreDefaults.BackColor = &H8000000F
    frmTransmitter!lblDataMissing.Visible = False
    cmdClearLog.BackColor = &HFFFFFF
    Unload Me
    frmPlanner.Show
 End Sub

Private Sub Form_Activate()

     giTimeFocus = 3
    'To prevent Run-Time Error if Planner control box 'close' (iHourNow)
    'is clicked while AddTime is selected, StopWatch, AddTime, memos & PlanHelp
    'send giTimeFocus = 3 as a control number when any of the forms is activated.

    If lstActivity.ListCount < 7 Then
        cmdClearLog.Enabled = False
    Else
        cmdClearLog.Enabled = True
    End If

    Dim daDate As String
    Dim sTemp As String
    Dim iResponse As Integer
    Dim sTemp1 As String
    
On Error GoTo HandleErrors
    
'--------

    Open "Access.txt" For Input As #104 'opens text file & restores data to screen
    Do Until EOF(104)
    Line Input #104, sTemp1
    lstAccess.AddItem sTemp1
    Loop
    Close #104
    
    'scrolls Access list box to last entry
    If lstAccess.ListCount > 6 Then
        lstAccess.TopIndex = lstAccess.ListCount - 5
    End If
    
    Dim iAcsCount As Integer
    iAcsCount = lstAccess.ListCount
    
    lblAccessOpened = "Activity Log has been opened " _
    & iAcsCount & " times."
    
    Label6 = iAcsCount

'--------
 
'Access Log, clears and begins new access log after 25 cycles

    If lstAccess.ListCount > 25 Then
        Dim aDate As String
        Dim aTime As String

        lstAccess.Clear

        aDate = Format(Now, "long Date")
        aTime = Format(Time, "hh:mm") '24 hour format

        Open "Access.txt" For Output As #104 '> 25 cycles, activity page open dates
        Print #104, aDate, aTime
        Close #104
   
        Open "Access.txt" For Input As #104 'opens text file & restores data to screen
        Do Until EOF(104)
        Line Input #104, sTemp1
        lstAccess.AddItem sTemp1
        Loop
        Close #104

        lblAccessOpened = ""
        Label4.BackColor = vbWhite
        Label4.Caption = " Log entries exceeded 25. Count reset. "
        Label6 = ""
    End If
    cmdClose.SetFocus
    Exit Sub
    
HandleErrors:
    Close #104
End Sub

Private Sub Form_Load()
    Dim sTemp As String
    Dim dDate As Date
    
    dDate = Now()
    
    lblToday = "Today is " & Format(dDate, "long Date")
    
    lblPath = CurDir$
    
On Error GoTo HandleErrors

    '--------------------
    Dim iOpens As Variant
    Dim idate As String

    Open "A_LogLog.txt" For Input As #102 'inputs activity text lines
    Input #102, iOpens, idate
    Close #102

    lblDate.Caption = Format(idate, "long Date") 'shows series beginning date
    
'========================================
  
    lstActivity.Clear
   
    Open "A_LogLog.txt" For Input As #102 'opens text file & restores data to screen
    Do Until EOF(102)
    Line Input #102, sTemp
    lstActivity.AddItem sTemp
    Loop
    Close #102

    lblSeventhLast = lstActivity.List(lstActivity.ListCount - 17)
    lblSixthLast = lstActivity.List(lstActivity.ListCount - 15)
    lblFifthLast = lstActivity.List(lstActivity.ListCount - 13)
    lblFourthLast = lstActivity.List(lstActivity.ListCount - 11)
    lblThirdLast = lstActivity.List(lstActivity.ListCount - 9)
    lblSecondLast = lstActivity.List(lstActivity.ListCount - 7) 'next-to-last
    lblNextToLast = lstActivity.List(lstActivity.ListCount - 5) 'next-to-last
    lblLast = lstActivity.List(lstActivity.ListCount - 3) 'shows last activity

    iCount = (Int(lstActivity.ListCount) - 1) / 2
    lblCount.Caption = iCount & "   MusicLog open/close cycles beginning with the entry dated"
    
    Label5 = iCount

    If lstActivity.ListCount > 17 Then
        lstActivity.TopIndex = lstActivity.ListCount - 17
    End If

    Exit Sub
    '-----------------------
HandleErrors:
    Close #102
    Close #104
    
    lstActivity.Clear
    
    Dim daDate As String
    daDate = Format(Date, "long Date")

    Open "A_LogLog.txt" For Output As #102 '3
    Print #102, ">Begin", daDate
    Close #102
    
    Dim aDate As String
    Dim aTime As String
    
    aDate = Format(Now, "long Date")
    aTime = Format(Time, "hh:mm") '24 hour format
   
    Open "Access.txt" For Output As #104 'activity page open dates /HandleError
    Print #104, aDate, aTime
    Close #104
   
    MsgBox "MusicLog activity log has been initialized. The activity log records the dates and times the MusicLog program is opened and closed." _
    & vbCrLf & vbCrLf & "This completes the the basic setup of the MusicLog program.", _
    vbOKOnly + vbInformation, "Activity Count Set to Ø"
    frmPlanner!txtHour.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #102
    Close #104
End Sub
