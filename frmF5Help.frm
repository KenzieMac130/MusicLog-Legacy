VERSION 5.00
Begin VB.Form frmF4Help 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Total Time Estimate"
   ClientHeight    =   6195
   ClientLeft      =   2310
   ClientTop       =   2985
   ClientWidth     =   10935
   Icon            =   "frmF5Help.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrintForm 
      Caption         =   "&Print a text copy of this page"
      Height          =   330
      Left            =   7710
      TabIndex        =   5
      Top             =   5610
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   5610
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmF5Help.frx":0442
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
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Width           =   10215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmF5Help.frx":0641
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
      Left            =   360
      TabIndex        =   3
      Top             =   3228
      Width           =   10215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label 3"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2257
      Width           =   10215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label 2"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1286
      Width           =   10215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmF5Help.frx":0739
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
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   10215
   End
End
Attribute VB_Name = "frmF4Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdPrintForm_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a text copy of this page?", vbYesNo, "Estimating Program Time")
    If iResponse = vbNo Then
        cmdExit.SetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then
       ' Printer.ColorMode = 1 'monochrome
        
        Printer.FontName = "Arial"
        Printer.FontSize = 11
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.Print Tab(6); "MusicLog:  Estimating Program Total Time"
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print
        Printer.Print Tab(7); "Selecting the F4 Function Key, or 'Include Estimated Announce Times' from the 'Format-Options' menu, gives"
        Printer.Print Tab(7); "you more flexibility and accuracy in estimating the total time required for a program segment. It does this by"
        Printer.Print Tab(7); "including with each music entry a text box for entering an estimated time needed to announce and back-announce"
        Printer.Print Tab(7); "that selection."
        Printer.Print
        Printer.Print Tab(7); "As a selection is entered, an announce time of " & Val(frmTimeRemain!txtIntro) & " seconds is automatically entered. However, this is just a"
        Printer.Print Tab(7); "beginning estimate. You can overwrite it with whatever you feel is more accurate. Incidentally, Double-Clicking"
        Printer.Print Tab(7); "an announce-time box reduces the time shown by half."
        Printer.Print
        Printer.Print Tab(7); "At the lower right-hand corner of the page there is a text entry box where you can enter the number of " & Val(frmTimeRemain!txtSpotLength) & " second"
        Printer.Print Tab(7); "(average length) spot or weather announcements in the hour.  " & Val(frmTimeRemain!txtSpotLength) & " seconds is shown as the current spot average"
        Printer.Print Tab(7); "length. On the 'Time Remain' page you can overwrite this value to change it."
        Printer.Print
        Printer.Print Tab(7); "Selecting the F4 Function Key links 'Music Planning' page to the 'Time Remain' page, where estimated announce"
        Printer.Print Tab(7); "and total program time computations actually occur. However, it is not necessary to go to the 'Time Remain'"
        Printer.Print Tab(7); "page to use this function."
        Printer.Print
        Printer.Print Tab(7); "On the 'Time Remain' page, an entry with 2 letters or less in the 'Lineup' text box is called a trial entry and will"
        Printer.Print Tab(7); "NOT be saved with the lineup. This feature allows you to enter temporary, trial entries into an existing lineup,"
        Printer.Print Tab(7); "then, as long as the entry's 'Music Lineup' text  box contains 2 letters or less, clear the temporary entries with"
        Printer.Print Tab(7); "the 'Remove Trial Entries' button without affecting the other entries in the lineup. (Note: the 'Clear Lineup'"
        Printer.Print Tab(7); "button clears and saves ALL entries in the lineup.)"
        
        Printer.FontSize = 8
        Printer.FontItalic = True
        Printer.Print
        Printer.Print Tab(30); "MusicLog 'How to Use this Page', help page #6  (program version " & App.Comments; ")"
        Printer.FontItalic = False
        'Printer.Print Tab(64); "###"
        Printer.EndDoc
    End If
    cmdExit.SetFocus
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

Private Sub Form_Load()
    Label2.Caption = "As a selection is entered, an announce time of " & Val(frmTimeRemain!txtIntro) _
    & " seconds is automatically entered. However, this is just a beginning estimate. You can overwrite it with whatever" _
    & " you feel is more accurate. Incidentally, Double-Clicking an announce-time box reduces the time shown by half."

    Label3.Caption = "At the lower right-hand corner of the page there is a text entry box where you can enter the number of " _
    & Val(frmTimeRemain!txtSpotLength) & " second (average length) spot or weather announcements in the hour. " & Val(frmTimeRemain!txtSpotLength) & _
    " seconds is shown as the current spot average length. On the 'Time Remain' page you can overwrite this value to change it."

End Sub

