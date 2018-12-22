VERSION 5.00
Begin VB.Form frmHelpCharacterAllotment 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"frmHelpCharacterAllotment.frx":0000
   ClientHeight    =   10335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelpCharacterAllotment.frx":0087
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print a copy of this page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7470
      TabIndex        =   10
      Top             =   9495
      Width           =   2280
   End
   Begin VB.CommandButton cmdCloseHint 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   420
      Left            =   4695
      TabIndex        =   9
      Top             =   9480
      Width           =   1620
   End
   Begin VB.Label lblMusicLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MusicLog"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   12
      Top             =   9435
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Text box character limits and the number that can help you complete an unfinished Composition entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   11
      Top             =   315
      Width           =   9960
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000005&
      Caption         =   $"frmHelpCharacterAllotment.frx":04C9
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   690
      TabIndex        =   8
      Top             =   8205
      Width           =   9600
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Caption         =   "And this is where the little number at the end of the line comes into play again. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   690
      TabIndex        =   7
      Top             =   7772
      Width           =   9285
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      Caption         =   $"frmHelpCharacterAllotment.frx":063D
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
      Left            =   690
      TabIndex        =   6
      Top             =   6891
      Width           =   9900
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   $"frmHelpCharacterAllotment.frx":0742
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   690
      TabIndex        =   5
      Top             =   5280
      Width           =   9735
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   $"frmHelpCharacterAllotment.frx":09D4
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   690
      TabIndex        =   4
      Top             =   3929
      Width           =   9870
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Caption         =   $"frmHelpCharacterAllotment.frx":0BA7
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
      Left            =   690
      TabIndex        =   3
      Top             =   3018
      Width           =   9720
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   $"frmHelpCharacterAllotment.frx":0C9C
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
      Left            =   690
      TabIndex        =   2
      Top             =   2092
      Width           =   9675
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   $"frmHelpCharacterAllotment.frx":0DBE
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   690
      TabIndex        =   1
      Top             =   1406
      Width           =   9630
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   $"frmHelpCharacterAllotment.frx":0E76
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
      Left            =   690
      TabIndex        =   0
      Top             =   750
      Width           =   9795
   End
End
Attribute VB_Name = "frmHelpCharacterAllotment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCloseHint_Click()
'close (unload) form
    Unload Me

End Sub

Private Sub cmdPrint_Click()

On Error GoTo HandleErrors

    Dim iResponse As Integer
    
    iResponse = MsgBox("Print a copy of this page?", vbYesNo, "Text Box Line Limits")
    If iResponse = vbNo Then
        Exit Sub
    ElseIf iResponse = vbYes Then
        lblMusicLog.Visible = True
        PrintForm
    End If
    lblMusicLog.Visible = False
    Exit Sub
    
HandleErrors:

    MsgBox "Printing Error. Check to be certain a printer is installed and selected.", _
    vbOKOnly, "Printing Error"

End Sub

