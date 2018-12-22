VERSION 5.00
Begin VB.Form frmTimeRemainPad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TimeRemain Pad"
   ClientHeight    =   3330
   ClientLeft      =   405
   ClientTop       =   2445
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   2745
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   330
      Left            =   952
      TabIndex        =   1
      Top             =   2925
      Width           =   840
   End
   Begin VB.ListBox lstList 
      Height          =   2790
      Left            =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2640
   End
End
Attribute VB_Name = "frmTimeRemainPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
    frmTimeRemain!chkPad = 0
End Sub

Private Sub Form_Activate()
    giTimeFocus = 3
'To prevent Run-Time Error if Planner control box 'close' (X)
'is clicked while stopwatch is selected, StopWatch, AddTime & TimeRemain Pad
'send giTimeFocus = 3 as a control number when any of the forms is activated.
    
End Sub

Private Sub Form_Load()
    Dim l As String 'lstLog horizontal scroll
    Dim lDummy As String 'lstLog  horizontal scroll
    l = SendMessage(lstList.hwnd, LB_SETHORIZONTALEXTENT, 400, lDummy)
End Sub
