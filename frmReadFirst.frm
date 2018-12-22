VERSION 5.00
Begin VB.Form frmReadFirst 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Getting Started with MusicLog                                                      Click Alt+P to print a text copy of this page"
   ClientHeight    =   10830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   Icon            =   "frmReadFirst.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print a text copy of this page"
      Height          =   345
      Left            =   7290
      TabIndex        =   1
      Top             =   10275
      Width           =   2475
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   330
      Left            =   4845
      TabIndex        =   0
      Top             =   10320
      Width           =   1350
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Formats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5250
      TabIndex        =   15
      Top             =   7110
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   13
      Left            =   690
      TabIndex        =   20
      Top             =   6435
      Width           =   9765
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":05E0
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
      Index           =   2
      Left            =   690
      TabIndex        =   19
      Top             =   1140
      Width           =   10020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":072E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Index           =   1
      Left            =   690
      TabIndex        =   18
      Top             =   180
      Width           =   10020
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Playing an LP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4927
      TabIndex        =   17
      Top             =   6165
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Entering Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4950
      TabIndex        =   16
      Top             =   1770
      Width           =   1395
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":090E
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
      Left            =   690
      TabIndex        =   14
      Top             =   7875
      Width           =   9825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":0A28
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   9
      Left            =   690
      TabIndex        =   13
      Top             =   9270
      Width           =   10020
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":0AED
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   690
      TabIndex        =   12
      Top             =   7395
      Width           =   10020
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":0BE2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   690
      TabIndex        =   11
      Top             =   8550
      Width           =   10020
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Continuing and Printing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4567
      TabIndex        =   10
      Top             =   9015
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":0CC4
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
      Index           =   10
      Left            =   690
      TabIndex        =   9
      Top             =   5670
      Width           =   9780
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":0DB5
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
      Index           =   14
      Left            =   690
      TabIndex        =   8
      Top             =   3930
      Width           =   10020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":0F58
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
      Index           =   3
      Left            =   690
      TabIndex        =   7
      Top             =   2025
      Width           =   10020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":10AC
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
      Index           =   4
      Left            =   690
      TabIndex        =   6
      Top             =   2725
      Width           =   10020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":118F
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
      Index           =   5
      Left            =   690
      TabIndex        =   5
      Top             =   4645
      Width           =   10140
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":1265
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
      Index           =   6
      Left            =   690
      TabIndex        =   4
      Top             =   3230
      Width           =   10020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":1400
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
      Index           =   7
      Left            =   690
      TabIndex        =   3
      Top             =   5165
      Width           =   10020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmReadFirst.frx":14B6
      Height          =   450
      Index           =   11
      Left            =   690
      TabIndex        =   2
      Top             =   9735
      Width           =   9360
   End
End
Attribute VB_Name = "frmReadFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo HandleErrors
     Dim iResponse As Integer

    iResponse = MsgBox("Print a text copy of this Help information?", vbYesNo, "Getting Started")
    
    If iResponse = vbNo Then
        cmdClose.SetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then

        Printer.FontName = "Times New Roman"
        Printer.FontSize = 11
        Printer.Print
             
        Printer.FontBold = True
        Printer.Print Tab(7); "Getting Started with MusicLog"
        
        Printer.FontSize = 8
        Printer.Print Tab(12); "program version " & App.Comments
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
       
        Printer.FontBold = False
        
        Printer.Print Tab(9); "The MusicLog planning page is designed to aid you to plan and organize a program of recorded classical"
        Printer.Print Tab(9); "music. The page is divided into three sections. At the top-left there are eight rows of text entry boxes called"
        Printer.Print Tab(9); "the Playlist. Below the playlist is a printable screen called the Program List. To the right are control"
        Printer.Print Tab(9); "buttons and checkboxes for choosing format options."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "Planning is done in hour blocks of time. The computer assumes an average 56 minutes of music and"
        Printer.Print Tab(9); "4 minutes of talk per hour. This 56 minutes is called the planned music time. However, using the 'Planned"
        Printer.Print Tab(9); "Music Time' box at the lower right of the page, you can change or reset the number of minutes of music"
        Printer.Print Tab(9); "planned for the period."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.FontBold = True
        Printer.Print Tab(7); "Entering Data"
        Printer.FontBold = False
        
        Printer.Print Tab(9); "Data entry begins in the upper right hand corner of the page with a text box for entering the start time of your"
        Printer.Print Tab(9); "show. If it starts at 8 o'clock, enter 8, then click the Tab key. The Tab key is the key to the successful use of"
        Printer.Print Tab(9); "this page. It moves the cursor in proper sequence through the text entry boxes."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "If you prefer to identify playlists by sequence rather than air times, select that option from the 'Format-Options'"
        Printer.Print Tab(9); "menu. Start time now will be identified on the printout as 'Hour #1' rather than (previous example) '8:00'."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "After entering the start hour, click the Tab key. The cursor will move to the first line Composer text entry box."
        Printer.Print Tab(9); "Enter the name of the composer of your first selection. Continuing with the Tab key, enter in sequence the"
        Printer.Print Tab(9); "composition title, the track number (optional), the CD or LP album number or station file number, the CD disc"
        Printer.Print Tab(9); "or LP side number and the composition's running time in minutes and seconds."
          
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "If the 'Liner or Album Notes' option is selected, a yellow text entry line will open following each entry for"
        Printer.Print Tab(9); "writing in additional information such as performer, orchestra or album information. This text entry line is"
        Printer.Print Tab(9); "limited to 106 characters (letters and spaces). Double-click the line to extend the limit to 140 characters."
        Printer.Print Tab(9); "Liner-album notes print as an additional line below the data entry line they accompany."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "To insert a bullet • into a liner note, a line-edit, or a title-page remarks line click the keyboard upper left"
        Printer.Print Tab(9); "tilde/ accent [~,]  key. To insert a chevron » click 'Shift' plus the same tilde/accent [~,] key."
            
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "When the first entry is complete, tab to the second line for the next entry. Remember, the Tab key is your"
        Printer.Print Tab(9); "guide, moving you along in proper sequence through the text entry boxes."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "As you complete each hour's playlist, click the 'Add Playlist to Printable Program List' button to add that hour's "
        Printer.Print Tab(9); "lineup to the printable Program List screen below, and to clear the Playlist text entry boxes for the next"
        Printer.Print Tab(9); "hour's entries."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.FontBold = True
        Printer.Print Tab(7); "Playing an LP"
        Printer.FontBold = False
        
        Printer.Print Tab(9); "If you are playing an LP rather than a CD, begin the 'CD Number' entry with the keyword, 'LP' (no quote marks,"
        Printer.Print Tab(9); "no leading spaces). The column heading will change from 'CD Number' to 'LP Number' and 'LP#' will appear in"
        Printer.Print Tab(9); "the text box. Follow that with the album number or station file number. For the Disc number enter the LP side"
        Printer.Print Tab(9); "number. The printout will identify the entry as an LP rather than a CD."
         
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        Printer.FontBold = True
        Printer.Print Tab(7); "Formats"
        Printer.FontBold = False
        
        Printer.Print Tab(9); "The Program List can be displayed and printed in several formats. From the 'Format-Options' menu you can"
        Printer.Print Tab(9); "choose to print the Program List in either Times New Roman or Arial type. Additional options are available from"
        Printer.Print Tab(9); "the Format-Options menu."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "From the 'Lineup & List Options' box you can select to space or not space entries and to include or not include"
        Printer.Print Tab(9); "liner notes. Two voice track formats are available. They are designed to be used when preparing pre-recorded"
        Printer.Print Tab(9); "voice tracks for later programming by a board operator."
                
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "All 'Lineup & List Options' as well as the Times New Roman or Arial option can be selected or de-selected at"
        Printer.Print Tab(9); "anytime. The Program List screen will show the current selection. What you see on the screen is what will print."
        
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.FontBold = True
        Printer.Print Tab(7); "Continuing and Printing"
        Printer.FontBold = False
        
        Printer.Print Tab(9); "As you add compositions to the playlist, time-readout boxes will display a running account of how much time"
        Printer.Print Tab(9); "you have scheduled in the hour and how much of the hour's planned music time remains."
                
        Printer.FontSize = 5
        Printer.Print
        Printer.FontSize = 11
        
        Printer.Print Tab(9); "When you have completed planning, click the 'Print List As...' button. What you see on the Program List screen is"
        Printer.Print Tab(9); "what will print. Additional instructions for entering, editing, and printing data are available under the 'How to"
        Printer.Print Tab(9); "use this page' menu."
       
        Printer.EndDoc
    End If

    cmdClose.SetFocus
    Exit Sub

HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub
