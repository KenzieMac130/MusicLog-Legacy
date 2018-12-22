VERSION 5.00
Begin VB.Form frmNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A Note About Clock Accuracy..."
   ClientHeight    =   7635
   ClientLeft      =   7680
   ClientTop       =   2010
   ClientWidth     =   6465
   Icon            =   "frmNote.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   6465
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   6870
   End
   Begin VB.CommandButton cmdPrintPage 
      Caption         =   "&Print a Copy of This Note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4845
      TabIndex        =   5
      Top             =   6915
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close Note"
      Height          =   450
      Left            =   2662
      TabIndex        =   3
      Top             =   6915
      Width           =   1110
   End
   Begin VB.Label Label8 
      Caption         =   "Computer clock time:"
      Height          =   255
      Left            =   390
      TabIndex        =   10
      Top             =   135
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   $"frmNote.frx":0442
      Height          =   690
      Left            =   390
      TabIndex        =   9
      Top             =   6120
      Width           =   5670
   End
   Begin VB.Label Label6 
      Caption         =   $"frmNote.frx":052E
      Height          =   660
      Left            =   390
      TabIndex        =   8
      Top             =   3197
      Width           =   5670
   End
   Begin VB.Label Label5 
      Caption         =   $"frmNote.frx":05FF
      Height          =   1410
      Left            =   390
      TabIndex        =   7
      Top             =   4593
      Width           =   5670
   End
   Begin VB.Label lblClock 
      Height          =   285
      Left            =   2010
      TabIndex        =   6
      Top             =   135
      Width           =   1290
   End
   Begin VB.Label Label4 
      Caption         =   $"frmNote.frx":07DC
      Height          =   510
      Left            =   390
      TabIndex        =   4
      Top             =   3970
      Width           =   5670
   End
   Begin VB.Label Label3 
      Caption         =   $"frmNote.frx":0865
      Height          =   855
      Left            =   390
      TabIndex        =   2
      Top             =   2229
      Width           =   5670
   End
   Begin VB.Label Label2 
      Caption         =   $"frmNote.frx":099A
      Height          =   645
      Left            =   390
      TabIndex        =   1
      Top             =   503
      Width           =   5670
   End
   Begin VB.Label Label1 
      Caption         =   $"frmNote.frx":0A58
      Height          =   855
      Left            =   390
      TabIndex        =   0
      Top             =   1261
      Width           =   5670
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrintPage_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a copy of this Note?", vbYesNo, "Print Clock Note")
    If iResponse = vbNo Then
        cmdClose.SetFocus
        Exit Sub
    ElseIf iResponse = vbYes Then
        PrintForm
    End If
    cmdClose.SetFocus
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"
End Sub

Private Sub Timer1_Timer()
 Dim Today As Variant
    Today = Now
    lblClock.Caption = Format(Today, "h:mm:ss ampm")
End Sub
