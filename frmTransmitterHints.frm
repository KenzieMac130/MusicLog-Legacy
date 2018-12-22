VERSION 5.00
Begin VB.Form frmTransmitterHints 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transmitter Power Comuptation Page Hints"
   ClientHeight    =   10230
   ClientLeft      =   2235
   ClientTop       =   2055
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrintForm 
      Caption         =   "&Print a text copy of this page"
      Height          =   360
      Left            =   6045
      TabIndex        =   9
      Top             =   9600
      Width           =   2520
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   405
      Left            =   4118
      TabIndex        =   0
      Top             =   9600
      Width           =   1065
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTransmitterHints.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   450
      TabIndex        =   12
      Top             =   5070
      Width           =   8400
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "• The Set Up page normally is opened from the Music Planning page ""Tools"" menu."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   465
      TabIndex        =   11
      Top             =   9120
      Width           =   8400
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTransmitterHints.frx":0148
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   450
      TabIndex        =   10
      Top             =   7530
      Width           =   8400
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "• Use the 'Settings' menu to update station efficiency values and transmitter power limits."
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
      TabIndex        =   8
      Top             =   270
      Width           =   8400
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTransmitterHints.frx":038B
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   450
      TabIndex        =   7
      Top             =   730
      Width           =   8400
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTransmitterHints.frx":0436
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
      TabIndex        =   6
      Top             =   3070
      Width           =   8400
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTransmitterHints.frx":0542
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
      Left            =   450
      TabIndex        =   5
      Top             =   1965
      Width           =   8400
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTransmitterHints.frx":06E6
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
      TabIndex        =   4
      Top             =   6220
      Width           =   8400
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   3
      Top             =   7100
      Width           =   8400
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTransmitterHints.frx":07CB
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
      Left            =   420
      TabIndex        =   2
      Top             =   3950
      Width           =   8400
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTransmitterHints.frx":090A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   450
      TabIndex        =   1
      Top             =   1370
      Width           =   8400
   End
End
Attribute VB_Name = "frmTransmitterHints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
    frmTransmitter.Show
End Sub

Private Sub cmdPrintForm_Click()
On Error GoTo HandleErrors

 Dim iResponse As Integer
    
    iResponse = MsgBox("Print a text copy of this Help information?", vbYesNo, "Printing Transmitter Hints")
    If iResponse = vbNo Then
        cmdClose.SetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then
     '   Printer.ColorMode = 1 'monochrome
        
        Printer.FontName = "Arial"
        Printer.FontSize = 11
        Printer.Print
        Printer.Print
        Printer.FontBold = True
        Printer.Print Tab(7); "Music Log:  Transmitter Power Computation Page Hints"
        Printer.FontBold = False

        Printer.FontSize = 10
        Printer.Print
        Printer.Print Tab(8); "• Use the 'Settings' menu to update station efficiency values and transmitter power limits."
        Printer.Print
        Printer.Print Tab(8); "• When updating Transmitter Efficiencies, a quick way to close the settings mode at the end of the update"
        Printer.Print Tab(8); "is to click the mouse in any side or lower margin open space."
        Printer.Print
        Printer.Print Tab(8); "• Pause the mouse cursor over a station's call letters for a readout of minimum and maximum power limit"
        Printer.Print Tab(8); "information for that station."
        Printer.Print
        Printer.Print Tab(8); "• If you do not want to take readings for a particular station, Double-Click the station's call letters to remove"
        Printer.Print Tab(8); "the station from the Tab Key data entry sequence. The call letters and data entry boxes will change color to"
        Printer.Print Tab(8); "gray, indicating the station is bypassed, and the Tab Key will skip the Volts and Amps data entry boxes for"
        Printer.Print Tab(8); "that station. Double-Click the call letters a second time to restore normal function."
        Printer.Print
        Printer.Print Tab(8); "• A quick way to clear data from an entry box is to enter a non-numeric character into the box.  A pop-up"
        Printer.Print Tab(8); "warning, 'Non - Numeric Entry' will appear. Click the 'Space Bar' or 'Enter' key, or 'Esc' key. This will both"
        Printer.Print Tab(8); "close the pop-up and clear the data entry box."
        Printer.Print
        Printer.Print Tab(8); "• Double-Clicking any line of text on the transmitter data display screen will cause the current transmitter"
        Printer.Print Tab(8); "readings to be copied to the screen. By selecting a check box on the music planning page print dialogue box,"
        Printer.Print Tab(8); "transmitter readings as shown on the screen can be printed at the bottom of the music list page."
        Printer.Print
        Printer.Print Tab(8); "• A convenient feature is limiting the digits to be highlighted as you tab through the Volt-Amp text entry"
        Printer.Print Tab(8); "boxes to the ones most likely to change with transmitter changing volt amp values. The digits to be"
        Printer.Print Tab(8); "highlighted (by selecting those less likely to change and NOT to be highlighted) can be set from the"
        Printer.Print Tab(8); "'Settings' menu."
        Printer.Print
        Printer.Print Tab(8); "v If entry data or settings are lost or corrupted, the 'Restore Default Entries' command button will restore Volts"
        Printer.Print Tab(8); "and Amps readings, Transmitter Efficiency and Power Limit values, and Station Call Letters to default entries."
        Printer.Print
        Printer.Print Tab(8); "• " & frmTransmitter!staPower.Panels(2).Text
        Printer.Print
        Printer.Print Tab(8); "v Note: When initiality setting up program, it is necessary to enter on the Set Up page at least the flag ship "
        Printer.Print Tab(8); "station call letters. If known, the station transmitter minimum and maximum powers should also be set,"
        Printer.Print Tab(8); "either as an approved percentage of Ideal power, or by direct entry."
        Printer.Print
        Printer.Print Tab(8); "• Next go to the transmitter page. From the 'Settings' menu, select 'Change Transmitter Efficiencies'."
        Printer.Print Tab(8); "Enter the efficiencies for each transmitter. Efficiencies normally will range from a low of about .50 (50%)"
        Printer.Print Tab(8); "to a high of about .90 (90%). In all cases efficiencies will be less than .99 (99%)."
        
        Printer.Print
        Printer.Print Tab(7); "• The Set Up page normally is opened from the Music Planning page 'Tools' menu."
                
        Printer.FontSize = 8
        Printer.Print
        Printer.FontItalic = True
        Printer.Print Tab(38); "MusicLog 'Transmitter Hints', help page #8  (program version " & App.Comments; ")"
        Printer.FontItalic = False
        Printer.EndDoc
    End If
    cmdClose.SetFocus
    
    Exit Sub
    
HandleErrors:

MsgBox "Printing Error. Check to be certain a printer is installed and selected.", vbOKOnly, "Printing Error"

End Sub

Private Sub Form_Load()
'Dim sDate As Date 'date of default transmitter data

    Label5 = "• " & frmTransmitter!staPower.Panels(2).Text '= "Default values current as of " & Format(rDate, "Long Date")
End Sub

