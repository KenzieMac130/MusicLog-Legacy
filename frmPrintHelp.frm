VERSION 5.00
Begin VB.Form frmPrintHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing the Lower Screen Program List"
   ClientHeight    =   8955
   ClientLeft      =   4095
   ClientTop       =   3960
   ClientWidth     =   8820
   Icon            =   "frmPrintHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print a  text copy of this page"
      Height          =   300
      Left            =   5820
      TabIndex        =   1
      Top             =   8355
      Width           =   2385
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   8430
      Width           =   1380
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPrintHelp.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   450
      TabIndex        =   11
      Top             =   4995
      Width           =   7740
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPrintHelp.frx":05E1
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
      Left            =   450
      TabIndex        =   10
      Top             =   4175
      Width           =   7665
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selecting print Planning Worksheet bypasses the print dialog box."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   450
      TabIndex        =   9
      Top             =   7295
      Width           =   6540
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPrintHelp.frx":06D5
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   450
      TabIndex        =   8
      Top             =   2550
      Width           =   7800
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPrintHelp.frx":07D5
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
      Left            =   450
      TabIndex        =   7
      Top             =   1730
      Width           =   7800
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPrintHelp.frx":08C2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   450
      TabIndex        =   6
      Top             =   685
      Width           =   7800
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "There are two printing choices, ""Music Logbook page"" and ""Triple-Spaced Worksheet."""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   450
      TabIndex        =   5
      Top             =   300
      Width           =   7530
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPrintHelp.frx":09F1
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   450
      TabIndex        =   4
      Top             =   6250
      Width           =   7800
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPrintHelp.frx":0B44
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
      Left            =   450
      TabIndex        =   3
      Top             =   3355
      Width           =   7800
   End
   Begin VB.Label lblSignature 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmPrintHelp.frx":0C51
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   465
      TabIndex        =   2
      Top             =   7665
      Width           =   7800
   End
End
Attribute VB_Name = "frmPrintHelp"
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
    
    iResponse = MsgBox("Print a text copy of this Help information?", vbYesNo, "Printing Lower Screen")
    If iResponse = vbNo Then
    
     cmdClose.SetFocus
     Exit Sub
     
     ElseIf iResponse = vbYes Then
     ' Printer.ColorMode = 1 'monochrome
     
     Printer.FontName = "Arial"
     Printer.FontSize = 12
     
     Printer.Print
     Printer.Print
     Printer.Print
     Printer.FontBold = True
     Printer.Print Tab(7); "Printing the Lower Screen Program List"
     Printer.FontBold = False
     Printer.FontSize = 10
     
     Printer.Print
     Printer.Print
     Printer.FontBold = True
     Printer.Print Tab(8); "Options"
     Printer.FontBold = False
     
     Printer.Print
     Printer.Print Tab(10); "There are two printing choices, 'Music Logbook page' and 'Tripple Spaced Worksheet.'"
     
     Printer.Print
     Printer.Print Tab(10); "Music Logbook page is the standard format. It prints the screen as shown. The header includes program"
     Printer.Print Tab(10); "and host names and the date. There is an option of adding a remarks line printed just below the page title"
     Printer.Print Tab(10); "line. Music Logbook page can be printed either portrait or landscape orientation."
     
     Printer.Print
     Printer.Print Tab(10); "Tripple Spaced Worksheet inserts extra space between entries for writing in additional information, notes"
     Printer.Print Tab(10); "or talking poindts. The header contains blank lines for writing in program and host names and date."
     
     Printer.Print
     Printer.FontBold = True
     Printer.Print Tab(8); "Voice Tracks"
     Printer.FontBold = False
     
    Printer.Print
    Printer.Print Tab(10); "There are two voice track formats designed to be used when preparing pre-recorded voice tracks for"
    Printer.Print Tab(10); "later programming by a board operator. Either voice track format can be selected or de-selected at"
    Printer.Print Tab(10); "any time for any playlist on the screen."
    
    Printer.Print
    Printer.FontBold = True
    Printer.Print Tab(8); "Multiple Pages"
    Printer.FontBold = False
    
    Printer.Print
    Printer.Print Tab(10); "A program list that is too long to print on a single page will print 'continued' at the page bottom and"
    Printer.Print Tab(10); "number the successive pages. The 'Line Count' text block at the lower right below the program list"
    Printer.Print Tab(10); "screen indicates the number of pages that will print. "
    
    Printer.Print
    Printer.Print Tab(10); "If more than one page will be printed, the number of pages is shown next to the words 'Line Count'."
    Printer.Print Tab(10); "Normal type size prints 58 lines to a page. If the enlarge or increase print type size option is selected,"
    Printer.Print Tab(10); "52 lines will be printed to a page."
     
    Printer.Print
    Printer.FontBold = True
    Printer.Print Tab(8); "Printing"
    Printer.FontBold = False
    
    Printer.Print
    Printer.Print Tab(10); "What you see on the Program List screen is what will print. From the 'Format-Options' menu you"
    Printer.Print Tab(10); "can choose to print the list in either Times New Roman or Arial type. If you have included liner notes or"
    Printer.Print Tab(10); "spaces, you may choose to view (and print) the list with or without the notes or spaces. Selecting from"
    Printer.Print Tab(10); "'Lineup & List Options' you can cycle through note, space and voice track options at any time"
          
    Printer.Print
    Printer.FontBold = True
    Printer.Print Tab(8); "Print Dialog Box"
    Printer.FontBold = False
    
    Printer.Print
    Printer.Print Tab(10); "Selecting print 'Music Logbook page' will open a print dialog box. On the print dialog box you can enter"
    Printer.Print Tab(10); "or select from drop down lists program and host names for the title line, change date settings if needed,"
    Printer.Print Tab(10); "enter remarks to be printed on a line below the title line, and select the number of copies to be printed."
         
    Printer.Print
    Printer.Print Tab(10); "Selecting print 'Planning Worksheet' directly bypasses the print dialog box."
    
    Printer.Print
    Printer.FontBold = True
    Printer.Print Tab(8); "Signature"
    Printer.FontBold = False
    
    Printer.Print
    Printer.Print Tab(10); "The signature line is printed as the last line on the page. It is shown at the bottom of the print dialog box."
    Printer.Print Tab(10); "Double-click the signature line to temporarily edit, overwrite, or delete it. Use the Defaults page to make"
    Printer.Print Tab(10); "permanent changes. The Defaults page requires the access code."
    
    Printer.FontSize = 8
    Printer.Print
    Printer.Print
    Printer.FontItalic = True
    Printer.Print Tab(35); "MusicLog 'How to Use this Page'(program version " & App.Comments; ")"
    Printer.FontItalic = False
    
    Printer.EndDoc
    End If
    cmdClose.SetFocus
    
    Exit Sub
    
HandleErrors:
    
    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub


