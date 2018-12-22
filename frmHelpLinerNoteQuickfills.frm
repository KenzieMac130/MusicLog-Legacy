VERSION 5.00
Begin VB.Form frmHelpLinerNoteQuickfills 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Liner Note Quickfills:  al\  t\  a\  p\         Title-Page Remarks Quickfill:   r          Inserts:  •   »  «    ¶     "
   ClientHeight    =   10950
   ClientLeft      =   4515
   ClientTop       =   3480
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelpLinerNoteQuickfills.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrintForm 
      Caption         =   "&Print a copy of this page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7110
      TabIndex        =   3
      Top             =   10140
      Width           =   2460
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4875
      TabIndex        =   0
      Top             =   10140
      Width           =   1365
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Liner Notes:    \ space space = [     :    ]          z\ = time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   705
      TabIndex        =   23
      Top             =   10200
      Width           =   3945
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "• WHAT ARE QUICKFILLS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   405
      TabIndex        =   22
      Top             =   270
      Width           =   2880
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHelpLinerNoteQuickfills.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   705
      TabIndex        =   21
      Top             =   7260
      Width           =   10035
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHelpLinerNoteQuickfills.frx":0530
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   720
      TabIndex        =   20
      Top             =   300
      Width           =   9795
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHelpLinerNoteQuickfills.frx":068A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   720
      TabIndex        =   19
      Top             =   1530
      Width           =   9795
   End
   Begin VB.Label lblMusicLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MusicLog"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7905
      TabIndex        =   18
      Top             =   10485
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The ""tilde/accent"" key normally is at the left end of the line of number keys."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   705
      TabIndex        =   17
      Top             =   8865
      Width           =   6690
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "• To insert a Chevron », Reverse Chevron «, or Paragraph Symbol ¶"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   420
      TabIndex        =   16
      Top             =   9240
      Width           =   7290
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHelpLinerNoteQuickfills.frx":07AE
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   720
      TabIndex        =   13
      Top             =   5145
      Width           =   9795
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHelpLinerNoteQuickfills.frx":085E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   720
      TabIndex        =   12
      Top             =   3375
      Width           =   9795
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "• LINER NOTE QUICKFILLS:   al\   t\   a\   p\"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   405
      TabIndex        =   11
      Top             =   1170
      Width           =   5130
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHelpLinerNoteQuickfills.frx":092C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   705
      TabIndex        =   10
      Top             =   6450
      Width           =   9660
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHelpLinerNoteQuickfills.frx":0A30
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   405
      TabIndex        =   9
      Top             =   8280
      Width           =   9630
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "After the program lineup is printed each music selection's actual start time can be hand-written in.  Example:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   4365
      Width           =   9795
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmHelpLinerNoteQuickfills.frx":0AEB
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   705
      TabIndex        =   7
      Top             =   9540
      Width           =   9795
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "• The ""QUARTERLY MUSIC REPORT"" remarks quickfill:  r\   prints on the title page just below     the page title line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   405
      TabIndex        =   6
      Top             =   5865
      Width           =   9885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The three [START-TIME] QUICKFILLS:   t\   a\  p\  are for use at the beginning of a quarterly music report  liner note "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   600
      TabIndex        =   5
      Top             =   2835
      Width           =   9720
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Album: Leon Fleisher / Cleveland Orchestra • Beethoven Piano Concertos 2 && 4 • Sony SBK 4865"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   900
      TabIndex        =   4
      Top             =   2310
      Width           =   9420
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[Start Time]   Performer(s) / Orchestra • Album Name • Record Label"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   885
      TabIndex        =   2
      Top             =   7800
      Width           =   6675
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "10:24"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1095
      TabIndex        =   1
      Top             =   4725
      Width           =   525
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[      :       ] Leon Fleisher / Cleveland Orchestra • Beethoven Piano Concertos 2 && 4 • Sony SBK 4865"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   900
      TabIndex        =   14
      Top             =   3945
      Width           =   9765
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "[      :       ] Leon Fleisher / Cleveland Orchestra • Beethoven Piano Concertos 2 && 4 • Sony SBK 4865"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   15
      Top             =   4725
      Width           =   9735
   End
End
Attribute VB_Name = "frmHelpLinerNoteQuickfills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    If frmPlanner!txtEdit.Text = " Double-Clicking a line of text on the screen highlights it and copies it into this box for editing." Then

        frmPlanner!cmdReplaceList.Visible = False
        frmPlanner!txtEdit.Visible = False
        frmPlanner!cmdEditClose.Visible = False
        frmPlanner!cmdDelete.Visible = False
        frmPlanner!cmdInsert.Visible = False
        frmPlanner!cmdDelete.Visible = False
        frmPlanner!cmdInsertSpot.Visible = False
        frmPlanner!cmdReplaceEdit.Visible = False
        frmPlanner!lblEditBox.Visible = False
        frmPlanner!shpRunTime(0).Visible = False
        frmPlanner!lblEdit.Visible = True
        frmPlanner!lblLineCount.Visible = True
        frmPlanner!cmdTypeSize.Visible = True
        frmPlanner!cmdReplaceList.Visible = True
        frmPlanner!cmdAddList.Visible = True
        frmPlanner!lstListNotes.ListIndex = -1
    End If
    
    frmPlanner!txtEdit.ForeColor = &H80000008
    Unload Me
End Sub

Private Sub cmdPrintForm_Click()
On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "Quickfills")
    If iResponse = vbNo Then
        Exit Sub
    ElseIf iResponse = vbYes Then
        lblMusicLog.Visible = True
        PrintForm
        
    End If
        lblMusicLog.Visible = False
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

