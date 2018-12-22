VERSION 5.00
Begin VB.Form frmPlanHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entering Data and Format Options                                           Click Alt+P to print a text copy of this page"
   ClientHeight    =   10860
   ClientLeft      =   1545
   ClientTop       =   330
   ClientWidth     =   10755
   Icon            =   "frmPlanHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleMode       =   0  'User
   ScaleWidth      =   10635.17
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print a text copy of this page"
      Height          =   300
      Left            =   7425
      TabIndex        =   1
      Top             =   10155
      Width           =   2460
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4695
      TabIndex        =   0
      Top             =   10155
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   435
      TabIndex        =   15
      Top             =   1760
      Width           =   9705
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   $"frmPlanHelp.frx":0551
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   435
      TabIndex        =   14
      Top             =   4220
      Width           =   9705
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":064A
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
      Left            =   435
      TabIndex        =   13
      Top             =   240
      Width           =   9705
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":0792
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   435
      TabIndex        =   12
      Top             =   5915
      Width           =   9705
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":0926
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   435
      TabIndex        =   11
      Top             =   9390
      Width           =   9705
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":0AE5
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   435
      TabIndex        =   10
      Top             =   8045
      Width           =   9705
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":0BCE
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   435
      TabIndex        =   9
      Top             =   5350
      Width           =   9705
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":0CA7
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   435
      TabIndex        =   8
      Top             =   6690
      Width           =   9705
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":0E11
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   435
      TabIndex        =   7
      Top             =   7465
      Width           =   9705
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":0EF6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   435
      TabIndex        =   6
      Top             =   8610
      Width           =   9705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":1046
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
      Left            =   435
      TabIndex        =   5
      Top             =   1000
      Width           =   9705
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":118E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   435
      TabIndex        =   4
      Top             =   4785
      Width           =   9705
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":1281
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   435
      TabIndex        =   3
      Top             =   2490
      Width           =   9705
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPlanHelp.frx":145F
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   435
      TabIndex        =   2
      Top             =   3445
      Width           =   9705
   End
End
Attribute VB_Name = "frmPlanHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    'close (unload) form
    Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo HandleErrors
     Dim iResponse As Integer

    iResponse = MsgBox("Print a text copy of this Help information?", vbYesNo, "Entering Data")
    If iResponse = vbNo Then
        cmdClose.SetFocus
        Exit Sub
        
    ElseIf iResponse = vbYes Then

        Printer.FontName = "Times New Roman"
        Printer.FontSize = 11
        Printer.Print
        Printer.Print
             
        Printer.FontBold = True
        Printer.Print Tab(7); "MusicLog: Entering Data and Format Options"
        
        Printer.FontSize = 6
        Printer.Print
        Printer.FontSize = 11
       
        Printer.FontBold = False
'---------------

        Printer.Print Tab(7); "LISTS: The Composer, Composition, CD, Disc, Track numbers and Min and Sec text entry boxes make up the"
        Printer.Print Tab(8); "'Playlist'. The screen below the playlist is the 'Program List' which can be printed. It is important when entering"
        Printer.Print Tab(8); "data into the playlist to use the TAB Key to progress through the text entry boxes in a sequential flow."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "HOUR BLOCKS: Program planning normally is done in hour-blocks. Begin by entering your program's starting"
        Printer.Print Tab(8); "hour into the text box at the top-right of the page. Enter the hour as (9, 10, etc.) which will print as 9:00 or"
        Printer.Print Tab(8); "10:00. An option allows arranging your playlist as a sequence (Hour 1, Hour 2, etc.) rather than by time."
       
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "THE 'PLANNED MUSIC TIME' box at the page lower right lists the number of minutes of music planned for"
        Printer.Print Tab(8); "the period. Since most programming is planned in hour-blocks, 56 minutes of music is the default, but you can"
        Printer.Print Tab(8); "change the time period to any number up to 120 minutes."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "ENTRIES: Use the TAB Key to advance the cursor through the text entry boxes. A character is each letter,"
        Printer.Print Tab(8); "number or space. The total number of characters in a line is limited to the number that can be contained"
        Printer.Print Tab(8); "within the width of a printed page. The small number that appears briefly upon completion of a line is the"
        Printer.Print Tab(8); "number of unused characters that still can be accommodated in the line. Double-clicking the number will"
        Printer.Print Tab(8); "place the cursor into the Composition text entry box."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "TRACK NUMBERS: To include CD/LP track numbers, click the 'Include CD Track Numbers' option. The"
        Printer.Print Tab(8); "Track and CD numbers entry boxes overlay each other. The Tab key normal sequence opens the Track# entry"
        Printer.Print Tab(8); "box first, then closes it and opens the CD# box. Double-clicking either the Track# or CD# text entry box"
        Printer.Print Tab(8); "will close that box and open the other one."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "LP: If you are playing an LP rather than a CD, begin the 'CD#' text box entry with the keyword, LP (no leading"
        Printer.Print Tab(8); "spaces).  Follow that with the album number or station file number. The printout now will identify the entry"
        Printer.Print Tab(8); "with LP# rather than CD#."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "LINER NOTES: You can add an additional line to each entry for brief program notes, such as performer,"
        Printer.Print Tab(8); "orchestra, or album information. To view a liner note click the composer's name in the composer text"
        Printer.Print Tab(8); "entry box which will open the note."
                
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "NON-MUSIC ENTRY: If your program includes a non-music item, such as an events-calendar or theater review,"
        Printer.Print Tab(8); "enter it on Line 8. Click the 'Format Line 8 as Non-Music Entry' check box to format line 8 for the entry."
             
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "WHEN A PLAYLIST IS COMPLETE, click the 'Add Playlist to Printable Program List' button to transfer"
        Printer.Print Tab(8); "the list to the screen below and clear the text boxes for the next hour's entries. The playlist will be ADDED"
        Printer.Print Tab(8); "at the bottom of any existing list. If the 'Replace List' command is clicked instead, the new playlist will"
        Printer.Print Tab(8); "REPLACE the existing list. Selecting 'Undo Replace' will restore the previous list."
               
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "PROGRAM LIST FORMATS:  Click the 'Insert a Space Between Entries' check box to insert a blank line"
        Printer.Print Tab(8); "between entries. The Program List can be displayed and printed in several formats. Options include two"
        Printer.Print Tab(8); "voice track formats, used when preparing pre-recorded voice tracks for later programming by a board operator."
        Printer.Print Tab(8); "What you see on the screen is what will print."
                
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "EDITING: To edit a Program List line of text, Double-Click the line to highlight it and open a text edit box"
        Printer.Print Tab(8); "and edit controls. For additional information, read 'Editing the Lower Screen' from the 'How to use this"
        Printer.Print Tab(8); "page' menu."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "ADDTIME CALCULATOR: Music Log includes a calculator that is useful if you need to add times, such as"
        Printer.Print Tab(8); "totaling a CD or LP selection's track playing times. Click the F9 key or the 'AddTime Calculator' button to"
        Printer.Print Tab(8); "open the calculator."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "SAVE: A program can be saved to an archives folder. To save a program to file, click the 'Save Program List"
        Printer.Print Tab(8); "As' button or click the F7 key. You also can select 'Save Current Program As'  from the 'File' menu. To open"
        Printer.Print Tab(8); "a previously saved program, click the F8 key, or select 'Open Saved Program List' from the 'File' menu."
                      
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(7); "NOTE: Double-clicking a 'Composer' text entry box copies that entire line of data into temporary memory."
        Printer.Print Tab(8); "Double-clicking any empty 'Composer' text entry box will paste the entire line of data into the empty line."
        Printer.Print Tab(8); "This allows you to shuffle positions in the lineup without necessarily having to retype an entry. Note,"
        Printer.Print Tab(8); "deleting a Composer entry will delete the entire line of accompanying data: Composition, Track, CD,"
        Printer.Print Tab(8); "Min, Sec and Liner Note."
        
        Printer.FontSize = 8
        Printer.Print
        Printer.FontItalic = True
        Printer.Print Tab(38); "MusicLog 'How to Use this Page' (program version " & App.Comments; ")"
        Printer.FontItalic = False
        
        Printer.EndDoc
        End If
    
        cmdClose.SetFocus
        Exit Sub

HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"

End Sub

Private Sub Form_Activate()
    giTimeFocus = 3
    'To prevent Run-Time Error if Planner control box 'close' (iHourNow)
    'is clicked while PlanHelp is selected, StopWatch, AddTime, memos & PlanHelp
    'send giTimeFocus = 3 as a control number when any of the forms is activated.
End Sub

