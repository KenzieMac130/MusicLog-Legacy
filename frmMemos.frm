VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMemos 
   BorderStyle     =   0  'None
   Caption         =   "Memos"
   ClientHeight    =   9660
   ClientLeft      =   2535
   ClientTop       =   3030
   ClientWidth     =   11310
   Icon            =   "frmMemos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrintPage 
      Caption         =   "Print a copy of this page"
      Height          =   330
      Left            =   8040
      TabIndex        =   5
      Top             =   9120
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1665
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   7200
      Width           =   10695
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1665
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5280
      Width           =   10695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1665
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3360
      Width           =   10695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1665
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1440
      Width           =   10695
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close Memo"
      Height          =   330
      Left            =   4927
      TabIndex        =   0
      Top             =   9120
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMemos.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   360
      Width           =   9975
   End
   Begin VB.Menu mnuEnterMemos 
      Caption         =   "Enter memos into spaces below"
   End
End
Attribute VB_Name = "frmMemos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    If Text1 = "" Then
        Text1 = "Enter memos into these text boxes"
    End If
    
    If Text1 <> "" Or Text2 <> "" Or Text3 <> "" Or Text4 <> "" Then
        Open "Memo.dat" For Output As #404
        Write #404, Text1, Text2, Text3, Text4
        Close #404
    End If
    
    Unload Me
End Sub

Private Sub cmdPrintPage_Click()
On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "Memos")
    If iResponse = vbNo Then
        Exit Sub
        
    ElseIf iResponse = vbYes Then
        With CommonDialog1 'common dialog print box
            .CancelError = True
            .Flags = cdlPDAllPages + cdlPDHidePrintToFile + cdlPDNoPageNums + cdlPDNoSelection
            .ShowPrinter
        End With
    
        Printer.Copies = CommonDialog1.Copies
        Printer.Orientation = CommonDialog1.Orientation
    
        PrintForm
    End If
Exit Sub
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

Private Sub Form_Activate()
    giTimeFocus = 3
       'To prevent Run-Time Error if Planner control box 'close' (iHourNow)
       'is clicked while AddTime is selected, StopWatch, AddTime, memos & PlanHelp
       'send giTimeFocus = 3 as a control number when any of the forms is activated.
End Sub

Private Sub Form_Load()

On Error GoTo HandleErrors

    Dim sText1 As String
    Dim sText2 As String
    Dim sText3 As String
    Dim sText4 As String
    
    Open "Memo.dat" For Input As #404
    Input #404, sText1, sText2, sText3, sText4
    Close #404
    
    Text1 = sText1
    Text2 = sText2
    Text3 = sText3
    Text4 = sText4
    
HandleErrors:
    Close #404
    Exit Sub
End Sub
