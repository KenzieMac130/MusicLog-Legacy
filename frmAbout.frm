VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Music Log                           Click Alt+P to print a text copy of this page"
   ClientHeight    =   10440
   ClientLeft      =   3840
   ClientTop       =   1830
   ClientWidth     =   8505
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7205.87
   ScaleMode       =   0  'User
   ScaleWidth      =   7986.635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   360
      Left            =   3540
      TabIndex        =   14
      Top             =   9615
      Width           =   1410
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print a text copy of this page"
      Height          =   330
      Left            =   5775
      TabIndex        =   13
      Top             =   9585
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   630
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   1035
      ScaleWidth      =   1320
      TabIndex        =   2
      Top             =   195
      Width           =   1320
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Use or distribution of this program without the expressed permission of the developer is a violation of copyright laws."
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
      Left            =   383
      TabIndex        =   20
      Top             =   10170
      Width           =   7515
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "All Rights Reserved"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   19
      Top             =   9270
      Width           =   1530
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "MusicLog"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3780
      TabIndex        =   18
      Top             =   8625
      Width           =   930
   End
   Begin VB.Label Label12 
      Caption         =   "Select one of the two Voice Track options if preparing pre-recorded voice tracks for later programming by a board operator."
      Height          =   435
      Left            =   585
      TabIndex        =   17
      Top             =   3552
      Width           =   7335
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      Caption         =   "Tamarack Rd,  Newtown, CT 06470"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3173
      TabIndex        =   16
      Top             =   9083
      Width           =   2145
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   1
      X1              =   1612.352
      X2              =   6345.172
      Y1              =   5849.593
      Y2              =   5839.24
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   " Copywrite © 2017  James W. Wright"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3060
      TabIndex        =   15
      Top             =   8895
      Width           =   2385
   End
   Begin VB.Label Label10 
      Caption         =   "MusicLog Planning Page is to assist you to plan, time, and print a classical music program playlist."
      Height          =   270
      Left            =   585
      TabIndex        =   12
      Top             =   1380
      Width           =   7110
   End
   Begin VB.Label Label9 
      Caption         =   $"frmAbout.frx":457C
      Height          =   1005
      Left            =   585
      TabIndex        =   11
      Top             =   4035
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":4732
      Height          =   450
      Left            =   585
      TabIndex        =   10
      Top             =   2571
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":4803
      Height          =   435
      Left            =   585
      TabIndex        =   9
      Top             =   3069
      Width           =   7335
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAbout.frx":4896
      Height          =   1035
      Left            =   585
      TabIndex        =   8
      Top             =   6459
      Width           =   7335
   End
   Begin VB.Label Label7 
      Caption         =   $"frmAbout.frx":4A4A
      Height          =   825
      Left            =   585
      TabIndex        =   7
      Top             =   1698
      Width           =   7335
   End
   Begin VB.Label lblRevision 
      Height          =   840
      Left            =   6300
      TabIndex        =   6
      Top             =   345
      Width           =   1590
   End
   Begin VB.Label Label6 
      Caption         =   $"frmAbout.frx":4BAB
      Height          =   450
      Left            =   585
      TabIndex        =   5
      Top             =   7830
      Width           =   7335
   End
   Begin VB.Label Label4 
      Caption         =   $"frmAbout.frx":4C50
      Height          =   435
      Left            =   585
      TabIndex        =   4
      Top             =   5976
      Width           =   7335
   End
   Begin VB.Label Label5 
      Caption         =   "A ""Stop Watch-Timer"" is called up by clicking the F12 Function Key."
      Height          =   240
      Left            =   585
      TabIndex        =   3
      Top             =   7542
      Width           =   7110
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      DrawMode        =   16  'Merge Pen
      Index           =   0
      X1              =   1640.523
      X2              =   6359.258
      Y1              =   5818.533
      Y2              =   5818.533
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":4D11
      Height          =   840
      Left            =   585
      TabIndex        =   1
      Top             =   5088
      Width           =   7335
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Music Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   480
      Left            =   2055
      TabIndex        =   0
      Top             =   480
      Width           =   2100
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a text copy of this Help information?", vbYesNo, "About Music Log")
    If iResponse = vbNo Then
        cmdOK.SetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then

        Printer.FontName = "Arial"
        Printer.FontSize = 12
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.Print Tab(7); "About MusicLog"
        Printer.FontBold = False
        
        Printer.FontSize = 10
        Printer.Print
        Printer.Print Tab(9); "MusicLog Planning Page is to assist you to plan, time, and print a classical music program Playlist."
        
        Printer.Print
        Printer.Print Tab(9); "Planning is done in hour blocks. Lines are provided for up to eight entries. Enter each composer's name,"
        Printer.Print Tab(9); "composition, disc information and playing time. Track numbers are optional. A non-music item, such as an"
        Printer.Print Tab(9); "events calendar or theater review, should be entered on Line 8 after selecting the 'Format Line 8 for a"
        Printer.Print Tab(9); "Non-Music Entry' checkbox."
        
        Printer.Print
        Printer.Print Tab(9); "If you would like to add a program note to an entry, select the 'Include Liner Or Album Notes' check box."
        Printer.Print Tab(9); "This adds an accompanying line to the entry for additional information such as performer or album."
        
        Printer.Print
        Printer.Print Tab(9); "You can choose whether to view (and print) or not view (and not print) the notes by selecting or de-selecting"
        Printer.Print Tab(9); "the liner notes check box."
        
        Printer.Print
        Printer.Print Tab(9); "Select one of the two Voice Track options if preparing pre-recorded voice tracks for later programming"
        Printer.Print Tab(9); "by a board operator."
        
        Printer.Print
        Printer.Print Tab(9); "When an hour's music list is complete, it is transferred from the Playlist to the Program List below with the"
        Printer.Print Tab(9); "'Add Playlist to Printable Program List' button. To insert (or remove) a space between Program List entires check"
        Printer.Print Tab(9); "(or uncheck) the 'Insert a Space Between; Entries' box. Whatever you see on this lower screen is what will print."
        Printer.Print Tab(9); "It can be printed in either a portrait or landscape orientation, Arial or Times New Roman font."
        
        Printer.Print
        Printer.Print Tab(9); "If you need to add up individual times, such as totaling a CD or LP track times, the Add Time calculator is"
        Printer.Print Tab(9); "designed for that. It totals time entered as minutes and seconds. It aso can be used to display the time"
        Printer.Print Tab(9); "remaining in a block-of-time as CD/LP playing times are subtracted from it. Click the 'AddTime Calculator'"
        Printer.Print Tab(9); "button or the F9 Key to call up the calculator."
        
        Printer.Print
        Printer.Print Tab(9); "The 'Transmitter' page (F3) computes transmitter power. Fill in the amp and volt readings and the power"
        Printer.Print Tab(9); "(wattage) is calculated and displayed. It warns of out-of-limit transmitter readings."
        
        Printer.Print
        Printer.Print Tab(9); "The 'Time Remain' page (F2) can be used to estimate a program's total running time, music and talk. It does"
        Printer.Print Tab(9); "this by adding to the music playing time an estimate of the time needed to introduce and back-announce each"
        Printer.Print Tab(9); "selection. Once the program is running, it can be used to estimate the time remaining in an hour as selections"
        Printer.Print Tab(9); "are played. The F4 function key links the Music Program Planning page and Time Remain page music lineups."
                        
        Printer.Print
        Printer.Print Tab(9); "A 'Stop Watch-Timer' is called up by clicking the F12 Function Key."
        
        Printer.Print
        Printer.Print Tab(9); "Each of these features is a separate page in the program. In addition to command buttons, they also can be"
        Printer.Print Tab(9); "selected from the 'Page' menu at the top of each page."
        
        Printer.FontSize = 7
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.Print Tab(30); "About Music Log, help page #9,  Version " & App.Comments
        Printer.FontBold = False
        Printer.EndDoc
    End If
    cmdOK.SetFocus
    
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

Private Sub Form_Load()

'======= DualV ===== select correct version


    lblRevision.Caption = "Version  " & App.Major & "." & App.Minor & "." & App.Revision _
    & vbCrLf & "Revision Date:" & vbCrLf & App.Comments
    Label8.Caption = App.LegalCopyright

End Sub


