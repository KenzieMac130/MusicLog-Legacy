VERSION 5.00
Begin VB.Form frmEditHelp 
   BackColor       =   &H00FFFFFF&
   Caption         =   " Editing the Lower Screen"
   ClientHeight    =   4110
   ClientLeft      =   3345
   ClientTop       =   315
   ClientWidth     =   11250
   Icon            =   "frmEditHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   4875
      TabIndex        =   4
      Top             =   3540
      Width           =   1380
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print a  text copy of this page"
      Height          =   300
      Left            =   7845
      TabIndex        =   5
      Top             =   3525
      Width           =   2385
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmEditHelp.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1709
      Width           =   10185
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmEditHelp.frx":04D3
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   3
      Top             =   2895
      Width           =   10650
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmEditHelp.frx":05A6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2301
      Width           =   10395
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmEditHelp.frx":06AA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   240
      TabIndex        =   1
      Top             =   892
      Width           =   10545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmEditHelp.frx":07D0
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   10650
   End
End
Attribute VB_Name = "frmEditHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdPrint_Click()

'On Error GoTo HandleErrors
'
'    Dim iResponse As Integer
'
'    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "Editing")
'    If iResponse = vbNo Then
'        Exit Sub
'    ElseIf iResponse = vbYes Then
'        PrintForm
'    End If
'
'    Exit Sub
'
'HandleErrors:
'
'    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
'    vbOKOnly, "Printing Error"

On Error GoTo HandleErrors

    Dim iResponse As Integer

    iResponse = MsgBox("Print a text copy of this Help information?", vbYesNo, "Editing Lower Screen")
    If iResponse = vbNo Then
        cmdClose.SetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then

        Printer.FontName = "Arial"
        Printer.FontSize = 12
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.Print Tab(5); "Editing the Lower Screen"
        Printer.FontBold = False
        
        Printer.Print
        Printer.FontSize = 11
        Printer.Print Tab(6); "SELECTING A LINE ON THE LOWER SCREEN TO EDIT: To edit a line of text on the lower screen,"
        Printer.Print Tab(7); "Double-Click the line to be edited. This will both highlight the line as well as copy it into a text edit"
        Printer.Print Tab(7); "box that will open above the screen."
        
        Printer.Print
        Printer.Print Tab(6); "EDITING: Make the desired text changes in the 'Edit Box', then select the command button labeled:"
        Printer.Print Tab(7); " 'Click to Replace Highlighted Line's Text with Edit Box Text'. The highLighted line's text will be"
        Printer.Print Tab(7); "replaced. If you decide not to make a change or have finished, click the 'Close Edit' button."
'
        Printer.Print
        Printer.Print Tab(6); "CAUTION: Editing a line of text in one view does not affect the corresponding text in the other views."
        Printer.Print Tab(7); "Each view must be edited individually."
        
        Printer.Print
        Printer.Print Tab(6); "The lower screen is not a spreadsheet. It is like a printed sheet of paper. Changes to the playing"
        Printer.Print Tab(7); "time of any music selection will not be reflected in the list's total playing time. You must type in"
        Printer.Print Tab(7); "any change if you want an accurate total time displayed."
        
        Printer.Print
        Printer.Print Tab(6); "VIEWS: The lower screen displays several views of the same playlist, depending upon the 'Lineup"
        Printer.Print Tab(7); "& List Option' selected. Each option can be selected or de-selected without affecting other"
        Printer.Print Tab(7); "options and views."
        
        Printer.Print
        Printer.FontSize = 8
        Printer.FontItalic = True
        Printer.Print
        Printer.Print Tab(30); "MusicLog 'How to Use this Page', help page #4  (program version " & App.Comments; ")"
        Printer.FontItalic = False
        Printer.EndDoc
    End If
    cmdClose.SetFocus

    Exit Sub

HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

