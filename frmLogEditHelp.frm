VERSION 5.00
Begin VB.Form frmLogEditHelp 
   Caption         =   "Screen Editing Help"
   ClientHeight    =   3855
   ClientLeft      =   2745
   ClientTop       =   345
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7665
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close Edit Help"
      Height          =   345
      Left            =   3135
      TabIndex        =   2
      Top             =   3420
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   $"frmLogEditHelp.frx":0000
      Height          =   1020
      Left            =   180
      TabIndex        =   4
      Top             =   2310
      Width           =   7290
   End
   Begin VB.Label Label4 
      Caption         =   $"frmLogEditHelp.frx":01D4
      Height          =   675
      Left            =   191
      TabIndex        =   3
      Top             =   1535
      Width           =   7290
   End
   Begin VB.Label Label3 
      Caption         =   $"frmLogEditHelp.frx":0305
      Height          =   795
      Left            =   198
      TabIndex        =   1
      Top             =   640
      Width           =   7290
   End
   Begin VB.Label Label2 
      Caption         =   $"frmLogEditHelp.frx":0468
      Height          =   420
      Left            =   198
      TabIndex        =   0
      Top             =   120
      Width           =   7290
   End
End
Attribute VB_Name = "frmLogEditHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    If frmPlanner!txtEdit.Text = "Double-Clicking a line of text on the screen highlights it and copies it into this box for editing." Then

        frmPlanner!cmdReplaceList.Visible = False
        frmPlanner!txtEdit.Visible = False
        frmPlanner!cmdEditClose.Visible = False
        frmPlanner!cmdDelete.Visible = False
        frmPlanner!cmdInsert.Visible = False
        frmPlanner!shpEdit.Visible = False
        frmPlanner!cmdDelete.Visible = False
        frmPlanner!cmdInsert.Visible = False
 
        frmPlanner!lblTotal2.Visible = True

        frmPlanner!lblNote.Visible = False
        frmPlanner!lblNote.Caption = "Note:"
        frmPlanner!cmdReplaceList.Visible = True
        frmPlanner!cmdAddList.Visible = True
        frmPlanner!lstList.ListIndex = -1
    End If

    Unload Me

End Sub




