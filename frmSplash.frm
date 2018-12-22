VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5955
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSplash.frx":000C
   MousePointer    =   1  'Arrow
   ScaleHeight     =   6051.567
   ScaleMode       =   0  'User
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5400
      Left            =   315
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   315
      Width           =   8445
      Begin VB.Timer Timer60SecPause 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   4590
         Top             =   270
      End
      Begin VB.Timer timeAccessCode 
         Interval        =   50000
         Left            =   3420
         Top             =   270
      End
      Begin VB.Timer TimerDelayOpen 
         Interval        =   400
         Left            =   3930
         Top             =   270
      End
      Begin VB.TextBox txtAccess 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   420
         MaxLength       =   6
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "123456"
         Top             =   4875
         Width           =   500
      End
      Begin VB.Timer timeOpen 
         Interval        =   15000
         Left            =   2910
         Top             =   270
      End
      Begin VB.Image Image10 
         Height          =   1065
         Left            =   6270
         Picture         =   "frmSplash.frx":044E
         Stretch         =   -1  'True
         Top             =   945
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Image Image9 
         Height          =   930
         Left            =   6615
         Picture         =   "frmSplash.frx":8B6F
         Stretch         =   -1  'True
         Top             =   885
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image Image8 
         Height          =   1125
         Left            =   6345
         Picture         =   "frmSplash.frx":15E231
         Stretch         =   -1  'True
         Top             =   855
         Width           =   990
      End
      Begin VB.Image Image7 
         Height          =   765
         Left            =   6810
         Picture         =   "frmSplash.frx":2B45D3
         Stretch         =   -1  'True
         Top             =   975
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Image Image6 
         Height          =   900
         Left            =   870
         Picture         =   "frmSplash.frx":2B7671
         Stretch         =   -1  'True
         Top             =   1035
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgCD 
         Height          =   1425
         Left            =   3127
         MouseIcon       =   "frmSplash.frx":2B9E65
         MousePointer    =   99  'Custom
         Picture         =   "frmSplash.frx":2BA16F
         Stretch         =   -1  'True
         Top             =   2475
         Width           =   2130
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Music Log"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   38.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   930
         Left            =   2310
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   990
         Width           =   3750
      End
      Begin VB.Image Image5 
         Height          =   900
         Left            =   7005
         Picture         =   "frmSplash.frx":2BB8FA
         Stretch         =   -1  'True
         Top             =   900
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Image Image4 
         Height          =   750
         Left            =   6510
         OLEDropMode     =   1  'Manual
         Picture         =   "frmSplash.frx":2BE84A
         Stretch         =   -1  'True
         Top             =   1110
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Image Image3 
         Height          =   780
         Left            =   7590
         Picture         =   "frmSplash.frx":2C26B6
         Stretch         =   -1  'True
         Top             =   945
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   6375
         Picture         =   "frmSplash.frx":2C33A0
         Top             =   270
         Width           =   1995
      End
      Begin VB.Image imgExit 
         Height          =   240
         Left            =   5820
         MousePointer    =   14  'Arrow and Question
         Picture         =   "frmSplash.frx":2C6902
         Top             =   4980
         Width           =   240
      End
      Begin VB.Label lblAccessCode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select and enter an Access Code of 3 to 6 digits, then click the CD Icon to continue."
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
         Height          =   480
         Left            =   1065
         TabIndex        =   5
         Top             =   4845
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Double-click to exit program"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   6135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   4973
         Width           =   2130
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "And Transmitter Power Computation Page"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   1905
         Width           =   3120
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Click CD Icon for Music Program Planning Page"
         Height          =   480
         Left            =   3240
         MouseIcon       =   "frmSplash.frx":2C6A04
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   4095
         Width           =   1905
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ElectroVoice RE 20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   420
         TabIndex        =   6
         Top             =   3645
         Width           =   1365
      End
      Begin VB.Image imgMicRE20 
         Height          =   1155
         Left            =   495
         MousePointer    =   1  'Arrow
         Picture         =   "frmSplash.frx":2C6D0E
         ToolTipText     =   " Click to pause program opening "
         Top             =   2415
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "RCA 77DX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6930
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         ToolTipText     =   " Click Mic for Transmitter Power Page "
         Top             =   4185
         Width           =   750
      End
      Begin VB.Image imgMic 
         Height          =   2400
         Left            =   6720
         MouseIcon       =   "frmSplash.frx":2CBF20
         MousePointer    =   14  'Arrow and Question
         Picture         =   "frmSplash.frx":2CC22A
         Stretch         =   -1  'True
         ToolTipText     =   " Click Mic for Transmitter Power Page "
         Top             =   1830
         Width           =   1140
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version 2.5  June 10, 2017"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   420
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   4440
         Width           =   1395
      End
      Begin VB.Label lblWelcome 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Welcome to..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   420
         MouseIcon       =   "frmSplash.frx":2D18CA
         TabIndex        =   1
         Top             =   420
         Width           =   2505
      End
      Begin VB.Label lblPause 
         BackColor       =   &H00FFFFFF&
         Caption         =   "click mic to pause program opening"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   405
         Left            =   420
         TabIndex        =   8
         Top             =   3885
         Width           =   1500
      End
      Begin VB.Image Image2 
         Height          =   750
         Left            =   6240
         Picture         =   "frmSplash.frx":2D1BD4
         Stretch         =   -1  'True
         Top             =   195
         Visible         =   0   'False
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim iTimeOut As Integer
Dim iTimeDelay As Integer
Option Explicit

Private Sub Form_Load()

'======= DualV ---- Select correct version. Comment out incorrect version


    lblVersion.Caption = "Version  " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.Comments
    
  '===========

    giTimeFocus = 2 'to set frmPlanner opening focus on txtHour

    Dim sAccess As String
On Error GoTo HandleErrors

    Open "Access.dat" For Input As #1
            Input #1, sAccess
    Close #1

    giAccess = sAccess 'Val(sAccess)
    txtAccess = sAccess
    
 '------------------------------Seasonal Images
    Dim Today As Variant
    Today = Now
    
    Dim DayVariable As Variant
    DayVariable = Format$(Today, "d")
    
    Dim MonthVariable As Variant
    MonthVariable = Format$(Today, "m")
    
    Dim MonthVariable2 As Integer
    MonthVariable2 = Val(MonthVariable)
    
    '------------December and Christmas
    
    If MonthVariable2 = 12 Then 'December
        Image6.Visible = True 'Christmas bell
    ElseIf MonthVariable2 <> 12 Then
        Image6.Visible = False
    End If
    
    If MonthVariable2 = 12 And DayVariable >= 10 Then 'December 10th or later
        Image1.Visible = False 'music banner
        Image2.Visible = True 'Christmas music banner
    Else
        Image2.Visible = False 'Christmas music banner
        Image1.Visible = True 'music banner
    End If
        
        '--------------
    
    If MonthVariable2 = 1 Then 'January
         Image4.Visible = True 'snow flake
    ElseIf MonthVariable2 <> 1 Then
        Image4.Visible = False
    End If
    
    If MonthVariable2 = 4 Then ' April
         Image3.Visible = True 'tulips
    ElseIf MonthVariable2 <> 4 Then
        Image3.Visible = False
    End If
    
     If MonthVariable2 = 5 Then 'May
         Image9.Visible = True 'blue flowers
    ElseIf MonthVariable2 <> 5 Then
        Image9.Visible = False
    End If

    If MonthVariable2 = 7 Then 'July
         Image8.Visible = True 'patriotic badge
    ElseIf MonthVariable2 <> 7 Then
        Image8.Visible = False
    End If
    
    If MonthVariable2 = 9 Then 'September
         Image5.Visible = True 'leaf
    ElseIf MonthVariable2 <> 9 Then
        Image5.Visible = False
    End If
    
    If MonthVariable2 = 10 Then 'October
         Image7.Visible = True 'fall leaves
    ElseIf MonthVariable2 <> 10 Then
        Image7.Visible = False
    End If
    
    If MonthVariable2 = 3 And DayVariable >= 10 Then  'March 10 or later
         Image10.Visible = True 'wind
    ElseIf MonthVariable2 <> 3 Then
        Image10.Visible = False
    End If

    '----------------------------------------
    
    Exit Sub

HandleErrors:
  
     timeOpen.Enabled = False
     timeAccessCode = False
     MsgBox "WELCOME to MusicLog" & vbCrLf & vbCrLf & "To begin, two setup steps are required: " & vbCrLf & vbCrLf & _
    "(1) Create a system Access Code, and" & vbCrLf & vbCrLf & "(2)  Enter call letters for at least the Flagship Radio Station." _
    & vbCrLf & vbCrLf & "The Access Code is necessary for certain program management activities. The Access Coded always is displayed in the lower" _
    & vbCrLf & "left hand corner of the opening or 'Welcome to MusicLog' page." _
    & vbCrLf & vbCrLf & _
    " The code box shows an initial setup code of 1234. To continue setup, overwrite 1234 with a code of 3 to 6 digits of your choice." _
    & vbCrLf & vbCrLf & "The Access code can be changed at any time by double-clicking the Version/Date information at the lower left of the 'Welcome" _
    & vbCrLf & "to MusicLog' page and overwriting the existing code.", _
    vbOKOnly + vbInformation, "MusicLog Setup"
    
    timeAccessCode.Enabled = True
   
    txtAccess.Enabled = True
    txtAccess = "1234"
    lblAccessCode.Visible = True
    txtAccess.BorderStyle = 1
    txtAccess.SelStart = 0 'begin selection at start
    txtAccess.SelLength = Len(txtAccess) 'selects # of characters
    txtAccess.MousePointer = 3
    Label5.Visible = False 'note to double-click to exit program
    imgExit.Visible = False
    Frame1.MousePointer = 1
    Label2.MousePointer = 1
    lblProductName.MousePointer = 1
   ' imgMic.MousePointer = 1
    'imgMic.Enabled = False 'mic icon
    Label1.Enabled = False 'mic label
    Label2.Enabled = False 'disc label
    
    txtAccess.ToolTipText = "Set Access Code, up to 6 digits. For future changes, first double-click the Version Data on the opening page, then overwrite the existing code."
    Close #1

  End Sub

Private Sub Form_Unload(Cancel As Integer)
    If txtAccess = "" Then
        txtAccess = "4321"
    End If
End Sub

Private Sub Frame1_DblClick()
    If Frame1.MousePointer = 11 Then
        lblWelcome_DblClick
    End If
End Sub

Private Sub imgExit_DblClick()

'to abort or close out opening process of the program
    
    Unload frmAbout
    Unload frmAddHelp
    Unload frmAddTime
    Unload frmDefaults
    Unload frmEditHelp
    Unload frmF4Help
    Unload frmLogActivity
    Unload frmMemos
    Unload frmNote
    Unload frmPlanHelp
    Unload frmPlanner
    Unload frmPrintHelp
   ' Unload frmSplash
    Unload frmStaff
    Unload frmStopWatch
    Unload frmTimeRemain
    Unload frmTransmitter
    Unload frmTransmitterHints
    lblWelcome.BorderStyle = 0
    Unload Me
   
End Sub

Private Sub imgMicRE20_Click()

 lblWelcome.BorderStyle = 0
    If timeOpen.Enabled = True Then
        timeOpen.Enabled = False
        timeAccessCode.Enabled = False
        Timer60SecPause.Enabled = True
        
        Frame1.MousePointer = 11
        imgMicRE20.MousePointer = 13
        lblWelcome.MousePointer = 1 'arrow
        lblVersion.ToolTipText = "Double-Click (twice) to set Access Code (below)"
        lblPause.ForeColor = &H8000& 'green
        lblPause.Caption = "60 sec opening pause. Click CD to continue."
        imgMicRE20.ToolTipText = " Click CD to continue "
    Else
        timeOpen.Enabled = True
        timeAccessCode.Enabled = True
        Timer60SecPause.Enabled = False
        Frame1.MousePointer = 1
        lblWelcome.MousePointer = 1
        lblVersion.ToolTipText = ""
        txtAccess.ToolTipText = ""
        txtAccess.Enabled = False
        lblPause.ForeColor = &HC0C0C0   'lite gray    '&H808080    'gray
        lblPause.Caption = "click mic to pause program opening"
        imgMicRE20.MousePointer = 1
    End If

     ' iTimeDelay = 1 'RE20 mic, an opening delay to prevent double-click's second click from landing on
                'MusicLog page and putting focus onto lstList rather than txtHour box.

End Sub

Private Sub imgCD_DblClick()
    iTimeDelay = 1
End Sub

Private Sub ImgMic_Click()
    lblPause.Visible = False
    frmTransmitter!cmdPrevious.Caption = "&Music Planner Page F1"
    Unload Me
    frmTransmitter.Show
End Sub
Private Sub ImgCD_Click()
    lblPause.Visible = False
    iTimeDelay = 1 'an opening delay to prevent double-click's second click from landing on
                'MusicLog page and putting focus onto lstList rather than txtHour box.
End Sub

Private Sub Label1_Click()
    frmTransmitter!cmdPrevious.Caption = "&Music Planner Page F1"
    Unload Me
    frmTransmitter.Show
End Sub

Private Sub Label2_Click()
    frmPlanner.Show
    Unload Me
End Sub

Private Sub Label5_DblClick()

'to abort or close out opening process of the program
    
    Unload frmAbout
    Unload frmAddHelp
    Unload frmAddTime
    Unload frmDefaults
    Unload frmEditHelp
    Unload frmF4Help
    Unload frmLogActivity
    Unload frmMemos
    Unload frmNote
    Unload frmPlanHelp
    Unload frmPlanner
    Unload frmPrintHelp
   ' Unload frmSplash
    Unload frmStaff
    Unload frmStopWatch
    Unload frmTimeRemain
    Unload frmTransmitter
    Unload frmTransmitterHints
    lblWelcome.BorderStyle = 0
    Unload Me
   
End Sub

Private Sub lblVersion_dblClick()

    If timeOpen.Enabled = True Then
        timeOpen.Enabled = False
        timeAccessCode.Enabled = True
        txtAccess.Enabled = True
        txtAccess.ToolTipText = "Set Access Code (6 digits max.)"
        txtAccess.MousePointer = 3 'I beam
        Frame1.MousePointer = 13 'arrow & hourglass
        lblVersion.ToolTipText = ""
        lblAccessCode.Visible = True
        txtAccess.BorderStyle = 1
        txtAccess.SelStart = 0 'begin selection at start
        txtAccess.SelLength = Len(txtAccess)  'selects # of characters
        Label5.Visible = False
        imgExit.Visible = False

    Else
        timeOpen.Enabled = True
        timeAccessCode.Enabled = False
        txtAccess.Enabled = False
        txtAccess.ToolTipText = ""
        Frame1.MousePointer = 1 'arrow
        lblAccessCode.Visible = False
        Label5.Visible = True
        imgExit.Visible = True
        txtAccess.BorderStyle = 0
    End If
End Sub

Private Sub lblWelcome_DblClick()

    If lblWelcome.BorderStyle = 0 Then
        lblWelcome.BorderStyle = 1
        timeOpen.Enabled = False
        timeAccessCode.Enabled = False
        lblWelcome.MousePointer = 11
        Frame1.MousePointer = 11
    ElseIf lblWelcome.BorderStyle = 1 Then
        lblWelcome.BorderStyle = 0
        timeOpen.Enabled = True
        timeAccessCode.Enabled = True
        Frame1.MousePointer = 1
        lblWelcome.MousePointer = 1
    End If
    lblPause.Visible = False
    
End Sub

Private Sub timeOpen_Timer()
    'on timer, to go to planner form after 1500 milliseconds (15 sec), normal procedure
    
    If txtAccess <> "" And lblAccessCode.Visible = False Then
        frmPlanner.Show
        Unload Me
    End If
End Sub

Private Sub timeAccessCode_Timer()
    'unloads after 50,000 milliseconds (50 sec)
    
    Dim iLengthAccess As Integer
    iLengthAccess = Len(txtAccess)
    
    If txtAccess <> "" And (iLengthAccess > 2) And lblAccessCode.Visible = True Then
        MsgBox "Timed out." & vbCrLf & vbCrLf & "Current Access Code of " & txtAccess & " is retained.", vbOKOnly + vbExclamation, "Access Code Retained"
        
    ElseIf (txtAccess = "" Or (iLengthAccess <= 2)) And lblAccessCode.Visible = True Then
        MsgBox "Timed out." & vbCrLf & vbCrLf & "An Access Code of at least 3 digits has not been entered." _
        & vbCrLf & vbCrLf & "The Access Code 3456 is assigned.", vbOKOnly + vbExclamation, "Access Code Assigned"
        txtAccess = "3456"
    End If
  
        Open "Access.dat" For Output As #1
        Write #1, txtAccess
        Close #1
        giAccess = txtAccess
        
        frmPlanner.Show
        Unload Me
End Sub

Private Sub Timer60SecPause_Timer()
    lblPause.Visible = False
    
    If lblWelcome.BorderStyle = 0 Then
        frmPlanner.Show
        Unload Me
    ElseIf lblWelcome.BorderStyle = 1 Then
        Unload Me
    End If
   
End Sub

Private Sub TimerDelayOpen_Timer()
 'delays opening for 400 milliseconds (0.4 sec)

    If iTimeDelay <> 0 Then '400

On Error GoTo HandleErrors

        If txtAccess = "" Then
            txtAccess = "5678"
        End If

        If txtAccess <> "" And lblAccessCode.Visible = True Then
    
            Open "Access.dat" For Output As #1
            Write #1, txtAccess
            Close #1
            giAccess = txtAccess
        End If
    
        lblAccessCode.Visible = False
        frmPlanner.Show
        Unload Me
        Exit Sub

HandleErrors:
        Unload frmPlanner
        Unload frmDefaults
        Unload Me
    End If
End Sub

Private Sub txtAccess_Change()
    If Not IsNumeric(txtAccess) And txtAccess <> "" And txtAccess <> "?" Then
        MsgBox "Access Code must be numeric, no more than 6 numbers.", vbOKOnly, "Entry Error"
        txtAccess = ""
        txtAccess.SetFocus
        Exit Sub
    End If

    If txtAccess <> "" Then
        Open "Access.dat" For Output As #1
        Write #1, txtAccess
        Close #1
        giAccess = txtAccess
    End If
    Label2.Enabled = True
End Sub

''Multi-version search word: Dual
'
'    'Station
'    lblVersion.Caption = "Version  " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.Comments
'
'   'Home
'   'lblVersion.Caption = "Version  " & App.Major & "." & App.Minor & "." & App.Revision & " DE" & vbCrLf & App.Comments
''-------------
'    giTimeFocus = 2 'to set frmPlanner opening focus on txtHour
'
'    Dim sAccess As String
'On Error GoTo HandleErrors
'
'    Open "Access.dat" For Input As #1
'            Input #1, sAccess
'    Close #1
'
'    giAccess = sAccess 'Val(sAccess)
'    txtAccess = sAccess
'    Exit Sub
'
'HandleErrors:
'
'     timeOpen.Enabled = False
'     timeAccessCode = False
'     MsgBox "WELCOME to MusicLog" & vbCrLf & vbCrLf & "To begin, two setup steps are required: " & vbCrLf & vbCrLf & _
'    "(1) Create a system Access Code, and" & vbCrLf & vbCrLf & "(2)  Enter call letters for at least the Flagship Radio Station." _
'    & vbCrLf & vbCrLf & "The Access Code is necessary for certain program management activities. The Access Coded always is displayed in the lower" _
'    & vbCrLf & "left hand corner of the opening or 'Welcome to MusicLog' page." _
'    & vbCrLf & vbCrLf & _
'    " The code box shows an initial setup code of 1234. To continue setup, overwrite 1234 with a code of 3 to 6 digits of your choice." _
'    & vbCrLf & vbCrLf & "The Access code can be changed at any time by double-clicking the Version/Date information at the lower left of the 'Welcome" _
'    & vbCrLf & "to MusicLog' page and overwriting the existing code.", _
'    vbOKOnly + vbInformation, "MusicLog Setup"
'
'    timeAccessCode.Enabled = True
'
'    txtAccess.Enabled = True
'    txtAccess = "1234"
'    lblAccessCode.Visible = True
'    txtAccess.BorderStyle = 1
'    txtAccess.SelStart = 0 'begin selection at start
'    txtAccess.SelLength = Len(txtAccess) 'selects # of characters
'    txtAccess.MousePointer = 3
'    Frame1.MousePointer = 1
'    Label2.MousePointer = 1
'    Label5.MousePointer = 1
'    lblProductName.MousePointer = 1
'    imgMic.MousePointer = 1
'    imgMic.Enabled = False 'mic icon
'    Label5.Enabled = False 'mic label
'    Label1.Enabled = False 'mic label
'    Label2.Enabled = False 'disc label
'
'    txtAccess.ToolTipText = "Set Access Code, up to 6 digits. For future changes, first double-click the Version Data on the opening page, then overwrite the existing code."
'    Close #1
'
'  End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If txtAccess = "" Then
'        txtAccess = "4321"
'    End If
'End Sub
'
'Private Sub imgMicRE20_DblClick()
'
'    lblWelcome.BorderStyle = 0
'    If timeOpen.Enabled = True Then
'        timeOpen.Enabled = False
'        timeAccessCode.Enabled = False
'        Timer60SecPause.Enabled = True
'
'        Frame1.MousePointer = 13
'        lblWelcome.MousePointer = 1 'arrow
'        lblVersion.ToolTipText = "Double-Click (twice) to set Access Code (below)"'
'        lblPause.Visible = True
'    Else
'        timeOpen.Enabled = True
'        timeAccessCode.Enabled = True
'        Timer60SecPause.Enabled = False
'
'        Frame1.MousePointer = 1
'        lblWelcome.MousePointer = 1
'        lblVersion.ToolTipText = ""
'        txtAccess.ToolTipText = ""
'        txtAccess.Enabled = False'
'        lblPause.Visible = False
'    End If
'
'End Sub
'
'Private Sub imgCD_DblClick()
'    iTimeDelay = 1
'End Sub
'
'Private Sub ImgMic_Click()
'    lblPause.Visible = False
'    frmTransmitter!cmdPrevious.Caption = "&Music Planner Page F1"
'    Unload Me
'    frmTransmitter.Show
'End Sub
'Private Sub ImgCD_Click()
'    lblPause.Visible = False
'    iTimeDelay = 1 'an opening delay to prevent double-click's second click from landing on
'                'MusicLog page and putting focus onto lstList rather than txtHour box.
'End Sub
'
'Private Sub Label1_Click()
'    frmTransmitter!cmdPrevious.Caption = "&Music Planner Page F1"
'    Unload Me
'    frmTransmitter.Show
'End Sub
'
'Private Sub Label2_Click()
'    frmPlanner.Show
'    Unload Me
'End Sub
'
'Private Sub Label5_Click()
'    frmTransmitter!cmdPrevious.Caption = "&Music Planner Page F1"
'    Unload Me
'    frmTransmitter.Show
'End Sub
'
'Private Sub lblVersion_dblClick()
'
'    If timeOpen.Enabled = True Then
'        timeOpen.Enabled = False
'        timeAccessCode.Enabled = True
'        txtAccess.Enabled = True
'        txtAccess.ToolTipText = "Set Access Code (6 digits max.)"
'        txtAccess.MousePointer = 3 'I beam
'        Frame1.MousePointer = 13 'arrow & hourglass
'        lblVersion.ToolTipText = ""
'        lblAccessCode.Visible = True
'        txtAccess.BorderStyle = 1
'        txtAccess.SelStart = 0 'begin selection at start
'        txtAccess.SelLength = Len(txtAccess)  'selects # of characters
'
'    Else
'        timeOpen.Enabled = True
'        timeAccessCode.Enabled = False
'        txtAccess.Enabled = False
'        txtAccess.ToolTipText = ""
'        Frame1.MousePointer = 1 'arrow
'        lblAccessCode.Visible = False
'        txtAccess.BorderStyle = 0
'    End If
'End Sub
'
'Private Sub lblWelcome_Click()
'
'    If lblWelcome.BorderStyle = 0 Then
'        lblWelcome.BorderStyle = 1
'        timeOpen.Enabled = False
'        timeAccessCode.Enabled = False
'        lblWelcome.MousePointer = 14
'    ElseIf lblWelcome.BorderStyle = 1 Then
'        lblWelcome.BorderStyle = 0
'        timeOpen.Enabled = True
'        timeAccessCode.Enabled = True
'        Frame1.MousePointer = 1
'        lblWelcome.MousePointer = 1
'    End If
'    lblPause.Visible = False
'
'End Sub
'
'Private Sub lblWelcome_DblClick()
'    'to abort or close out opening process of the program
'    lblWelcome.BorderStyle = 1
'    Unload frmAbout
'    Unload frmAddHelp
'    Unload frmAddTime
'    Unload frmDefaults
'    Unload frmEditHelp
'    Unload frmF4Help
'    Unload frmLogActivity
'    Unload frmNote
'    Unload frmPlanHelp
'    Unload frmPlanner
'    Unload frmPrintHelp
'   ' Unload frmSplash
'    Unload frmStaff
'    Unload frmStopWatch
'    Unload frmTimeRemain
'    Unload frmTransmitter
'    Unload frmTransmitterHints
'    lblWelcome.BorderStyle = 0
'    Unload Me
'End Sub
'
'Private Sub timeOpen_Timer()
'    'on timer, to go to planner form after 1500 milliseconds (15 sec), normal procedure
'
'    If txtAccess <> "" And lblAccessCode.Visible = False Then
'        frmPlanner.Show
'        Unload Me
'    End If
'End Sub
'
'Private Sub timeAccessCode_Timer()
'    'unloads after 50,000 milliseconds (50 sec)
'
'    Dim iLengthAccess As Integer
'    iLengthAccess = Len(txtAccess)
'
'    If txtAccess <> "" And (iLengthAccess > 2) And lblAccessCode.Visible = True Then
'        MsgBox "Timed out." & vbCrLf & vbCrLf & "Current Access Code of " & txtAccess & " is retained.", vbOKOnly + vbExclamation, "Access Code Retained"
'
'    ElseIf (txtAccess = "" Or (iLengthAccess <= 2)) And lblAccessCode.Visible = True Then
'        MsgBox "Timed out." & vbCrLf & vbCrLf & "An Access Code of at least 3 digits has not been entered." _
'        & vbCrLf & vbCrLf & "The Access Code 3456 is assigned.", vbOKOnly + vbExclamation, "Access Code Assigned"
'        txtAccess = "3456"
'    End If
'
'        Open "Access.dat" For Output As #1
'        Write #1, txtAccess
'        Close #1
'        giAccess = txtAccess
'
'        frmPlanner.Show
'        Unload Me
'End Sub
'
'Private Sub Timer60SecPause_Timer()
'    lblPause.Visible = False
'
'    If lblWelcome.BorderStyle = 0 Then
'        frmPlanner.Show
'        Unload Me
'    ElseIf lblWelcome.BorderStyle = 1 Then
'        Unload Me
'    End If
'
'End Sub
'
'Private Sub TimerDelayOpen_Timer()
' 'delays opening for 400 milliseconds (0.4 sec)
'
'    If iTimeDelay <> 0 Then '400
'
'On Error GoTo HandleErrors
'
'        If txtAccess = "" Then
'            txtAccess = "5678"
'        End If
'
'        If txtAccess <> "" And lblAccessCode.Visible = True Then
'
'            Open "Access.dat" For Output As #1
'            Write #1, txtAccess
'            Close #1
'            giAccess = txtAccess
'        End If
'
'        lblAccessCode.Visible = False
'        frmPlanner.Show
'        Unload Me
'        Exit Sub
'
'HandleErrors:
'        Unload frmPlanner
'        Unload frmDefaults
'        Unload Me
'    End If
'End Sub
'
'Private Sub txtAccess_Change()
'    If Not IsNumeric(txtAccess) And txtAccess <> "" And txtAccess <> "?" Then
'        MsgBox "Access Code must be numeric, no more than 6 numbers.", vbOKOnly, "Entry Error"
'        txtAccess = ""
'        txtAccess.SetFocus
'        Exit Sub
'    End If
'
'    If txtAccess <> "" Then
'        Open "Access.dat" For Output As #1
'        Write #1, txtAccess
'        Close #1
'        giAccess = txtAccess
'    End If
'    Label2.Enabled = True
'End Sub
