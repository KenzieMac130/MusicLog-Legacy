VERSION 5.00
Begin VB.Form frmAddHelp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Add Time  -  Help"
   ClientHeight    =   8190
   ClientLeft      =   4800
   ClientTop       =   2655
   ClientWidth     =   7035
   Icon            =   "frmAddHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7035
   Begin VB.CommandButton cmdPrintPage 
      Caption         =   "&Print a text copy of this page"
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
      Left            =   4395
      TabIndex        =   6
      Top             =   7515
      Width           =   2355
   End
   Begin VB.CommandButton cmdReturn 
      Cancel          =   -1  'True
      Caption         =   "&Close Help"
      Height          =   360
      Left            =   2842
      TabIndex        =   3
      Top             =   7515
      Width           =   1350
   End
   Begin VB.Label lblHelp6 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAddHelp.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   495
      TabIndex        =   8
      Top             =   6060
      Width           =   6045
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The ""Block/Remain"" section shows what remains of a reference block of time as times entered on the keypad are subtracted from it."
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
      Left            =   495
      TabIndex        =   7
      Top             =   3291
      Width           =   6045
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAddHelp.frx":0578
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   495
      TabIndex        =   5
      Top             =   1452
      Width           =   6045
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAddHelp.frx":063E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   495
      TabIndex        =   4
      Top             =   285
      Width           =   6045
   End
   Begin VB.Label lblHelp5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAddHelp.frx":072E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   495
      TabIndex        =   2
      Top             =   3963
      Width           =   6045
   End
   Begin VB.Label lblHelp4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAddHelp.frx":0843
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
      Left            =   495
      TabIndex        =   1
      Top             =   5130
      Width           =   6045
   End
   Begin VB.Label lblHelp3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAddHelp.frx":08F8
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
      Left            =   495
      TabIndex        =   0
      Top             =   2394
      Width           =   6045
   End
End
Attribute VB_Name = "frmAddHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrintPage_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a text copy of this Help information?", vbYesNo, "AddTime")
    If iResponse = vbNo Then
        cmdReturn.SetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then
    
        'Printer.FontName = "Arial"
        Printer.FontName = "Times New Roman"
        Printer.FontSize = 12
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.Print Tab(8); "Add Times Calculator"
        Printer.FontBold = False
'        Printer.FontSize = 8
'        Printer.Print Tab(12); "Music Log, Help Page #6  (Version " & App.Major & "." & App.Minor & "." & App.Revision & ", " & App.Comments & ")"
        Printer.FontSize = 11
        Printer.Print
        Printer.Print Tab(9); "'AddTime' is an adding calculator which produces a running total of individual times entered as minutes"
        Printer.Print Tab(9); "and seconds. Examples are adding up a selection's CD or LP track times, or totaling the times of"
        Printer.Print Tab(9); "individual selections in a program."
        Printer.Print
        Printer.Print Tab(9); "Always use the 'Tab' key to advance the cursor in sequence through the Minute and Second data entry"
        Printer.Print Tab(9); "boxes. As you enter the individual times, they are totaled and displayed on the screens below."
        Printer.Print
        Printer.Print Tab(9); "Total time is displayed in four formats: (1) Minutes and Seconds, (2) Hours and Minutes, (3) Hours and"
        Printer.Print Tab(9); "Decimal Fraction of Hours, and (4) Total Seconds."
        Printer.Print
        Printer.Print Tab(9); "The 'Block/Remain' section shows what remains of a reference block of time as times entered on the"
        Printer.Print Tab(9); "keypad are subtracted from it."
        Printer.Print
        Printer.Print Tab(9); "For example, if you plan 56 minutes of music for the hour and enter 56 into the 'Block/Remain' text entry"
        Printer.Print Tab(9); "'Min' box, a running total of what remains of that 56 minutes will be displayed as the individual CD/LP"
        Printer.Print Tab(9); "playing times you enter into the keypad are subtracted from it."
        Printer.Print
        Printer.Print Tab(9); "The 'Clear Entries' button clears all data from the 'Min' and 'Sec' boxes. The 'Set Block Time' or 'Clear Block'"
        Printer.Print Tab(9); "button clears data only from the 'Block/Remain' data entry box."
        Printer.Print
        Printer.Print Tab(9); "The 'Additional Entries' command totals the existing entries, places the total in the first min-sec boxes,"
        Printer.Print Tab(9); "and clears the remaining time boxes for additional entries. This is useful if you have more than eleven"
        Printer.Print Tab(9); "entries to add, or simply want to total what you have and add additional entries to that total."
        
        Printer.FontSize = 8
        Printer.Print
        Printer.FontItalic = True
        Printer.Print Tab(38); "MusicLog, 'Adding Times', help page #7  (program version " & App.Comments; ")"
        Printer.FontItalic = False
        'Printer.Print Tab(64); "###"
        Printer.EndDoc
    End If
    cmdReturn.SetFocus
    
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

Private Sub cmdReturn_Click()
    Unload Me
End Sub

