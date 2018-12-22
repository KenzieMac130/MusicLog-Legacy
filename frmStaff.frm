VERSION 5.00
Begin VB.Form frmStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add or Remove Names on Print-Dialogue-Box Program and Host Dropdown Lists"
   ClientHeight    =   7590
   ClientLeft      =   2565
   ClientTop       =   1695
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Program List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5595
      Left            =   315
      TabIndex        =   8
      Top             =   435
      Width           =   3420
      Begin VB.ListBox cboProgram 
         Height          =   3570
         Left            =   255
         Sorted          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1395
         Width           =   2880
      End
      Begin VB.TextBox txtProgram 
         Height          =   315
         Left            =   263
         MaxLength       =   35
         TabIndex        =   12
         ToolTipText     =   "Enter name of program to be added to Program list."
         Top             =   360
         Width           =   2865
      End
      Begin VB.CommandButton cmdProgramRemove 
         Caption         =   "Delete Selected Progra&m"
         Height          =   315
         Left            =   660
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Select Program to be deleted."
         Top             =   5115
         Width           =   2070
      End
      Begin VB.CommandButton cmdProgramAdd 
         Caption         =   "Add &Program"
         Height          =   360
         Left            =   255
         TabIndex        =   10
         ToolTipText     =   "Add the above name to the list of Programs."
         Top             =   795
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelProgram 
         Caption         =   "C&ancel"
         Height          =   360
         Left            =   2175
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   795
         Width           =   855
      End
      Begin VB.Shape shpProgram 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   360
         Top             =   825
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Host List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5595
      Left            =   4185
      TabIndex        =   3
      Top             =   450
      Width           =   3420
      Begin VB.CommandButton cmdCancelHost 
         Caption         =   "Ca&ncel"
         Height          =   360
         Left            =   2130
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   780
         Width           =   855
      End
      Begin VB.ListBox cboYou 
         Height          =   3570
         Left            =   270
         Sorted          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1380
         Width           =   2880
      End
      Begin VB.TextBox txtHost 
         Height          =   315
         Left            =   270
         MaxLength       =   26
         TabIndex        =   0
         ToolTipText     =   "Enter name of host to be added to Host list."
         Top             =   360
         Width           =   2880
      End
      Begin VB.CommandButton cmdHostAdd 
         Caption         =   "Add &Host"
         Height          =   360
         Left            =   285
         TabIndex        =   1
         ToolTipText     =   "Add the above name to the list of Hosts."
         Top             =   780
         Width           =   1155
      End
      Begin VB.CommandButton cmdHostRemove 
         Caption         =   "Delete Selected H&ost"
         Height          =   315
         Left            =   675
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Select Host to be deleted."
         Top             =   5115
         Width           =   2070
      End
      Begin VB.Shape shpHost 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   390
         Top             =   810
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdLog 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   465
      Left            =   3375
      TabIndex        =   2
      Top             =   6840
      Width           =   1185
   End
   Begin VB.Label lblHostCount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5490
      TabIndex        =   15
      Top             =   6090
      Width           =   810
   End
   Begin VB.Label lblProgramCount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1605
      TabIndex        =   14
      Top             =   6090
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "A dropdown list name cannot be edited. It can only be deleted and a corrected replacement added."
      Height          =   285
      Left            =   330
      TabIndex        =   7
      Top             =   6420
      Width           =   7275
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MusicLog, form: Staff

Dim miCount As Integer
Option Explicit

Private Sub cmdCancelHost_Click()
    txtHost = ""
    cboYou.ListIndex = -1 'removes highlight from List line
    frmStaff!shpProgram.Visible = False
    frmStaff!shpHost.Visible = False
    
    'lblHostCount = cboYou.ListCount & " entries"
    
If cboYou.ListCount = 0 Then
   lblHostCount = ""
ElseIf cboYou.ListCount = 1 Then
    lblHostCount = cboYou.ListCount & " entry"
ElseIf cboYou.ListCount > 1 Then
    lblHostCount = cboYou.ListCount & " entries"
End If
    
End Sub

Private Sub cmdCancelProgram_Click()
    txtProgram = ""
    cboProgram.ListIndex = -1 'removes highlight from List line
    frmStaff!shpProgram.Visible = False
    frmStaff!shpHost.Visible = False
    'lblProgramCount = cboProgram.ListCount & " entries"
    
If cboProgram.ListCount = 0 Then
    lblProgramCount = ""
ElseIf cboProgram.ListCount = 1 Then
    lblProgramCount = cboProgram.ListCount & " entry"
ElseIf cboProgram.ListCount > 1 Then
    lblProgramCount = cboProgram.ListCount & " entries"
End If

End Sub

Private Sub cmdHostAdd_Click()
    miCount = miCount + 1
    If txtHost = "" Then
       MsgBox "There is no entry to add", vbOKOnly, "No Data"
       miCount = 0
    ElseIf txtHost <> "(Enter Host's Name)" And txtHost <> "" Then
      cboYou.AddItem txtHost
      txtHost = ""
    Else
        MsgBox "(Enter Host's Name) is an invalid entry", vbOKOnly, "Invalid Entry"
        txtHost = ""
    End If

    frmStaff!shpProgram.Visible = False
    frmStaff!shpHost.Visible = False
    
    frmPlanner!cmdProgramHosts.Width = 3960
    frmPlanner!cmdProgramHosts.Left = 427
    frmPlanner!cmdProgramHosts.BackColor = &HFFFFFF  'white' &H8000000F 'gray
    frmPlanner!cmdProgramHosts.Caption = "Click to Edit or Add Name to Program or Host Dropdown List"
    frmPlanner!cmdProgramHosts.TabStop = False
    
If cboYou.ListCount = 1 Then
    lblHostCount = cboYou.ListCount & " entry "
ElseIf cboYou.ListCount > 1 Then
    lblHostCount = cboYou.ListCount & " entries "
End If
    
   ' lblHostCount = cboYou.ListCount & " entries"
        
    If cboYou.ListCount > 20 Then
        MsgBox "The Host List contains more than 20 entries." & vbCrLf & vbCrLf & _
        "It is suggested using the Delete Selected Host button you reduce the list to no more than 20 entries", _
        vbOKOnly, "Host List Exceeds 20 Entries"
    End If
    
    frmStaff!shpHost.Visible = False
    cmdLog.SetFocus
End Sub

Private Sub cmdHostRemove_Click()
  miCount = miCount + 1
      Dim iResponse As Integer
    'remove selected line
    If cboYou.ListIndex >= 0 Then 'if a line is selected
        iResponse = MsgBox("Remove '" & cboYou.List(cboYou.ListIndex) & "' from the host list?", vbYesNo, "Confirm Delete")
        If iResponse = vbYes Then
            cboYou.RemoveItem cboYou.ListIndex
            cboYou.ListIndex = -1
        End If
    Else
        MsgBox "No host name has been selected to delete.", vbOKOnly, "No Selection"
        miCount = 0
    End If
    
    'lblHostCount = cboYou.ListCount & " entries"
    
If cboYou.ListCount = 0 Then
   lblHostCount = ""
ElseIf cboYou.ListCount = 1 Then
    lblHostCount = cboYou.ListCount & " entry"
ElseIf cboYou.ListCount > 1 Then
    lblHostCount = cboYou.ListCount & " entries"
End If
    
    cmdLog.SetFocus
End Sub

Private Sub cmdLog_Click()
    txtProgram = ""
    txtHost = ""

    frmPlanner!cmdSaveProgramName.Visible = False
    
    frmPlanner!cmdProgramHosts.Width = 4170
    frmPlanner!cmdProgramHosts.Left = 540
    frmPlanner!cmdProgramHosts.BackColor = &HFFFFFF  'white'
    frmPlanner!cmdProgramHosts.Caption = "Click to Add a Name or Edit Program or Host Dropdown List"

    frmStaff!shpProgram.Visible = False
    frmStaff!shpHost.Visible = False
    
    Unload frmStaff
    frmStaff!shpProgram.Visible = False
    frmStaff!shpHost.Visible = False
    giHost = 0
    
End Sub

Private Sub cmdProgramAdd_Click()
    miCount = miCount + 1
    If txtProgram = "" Then
       MsgBox "There is no entry to add", vbOKOnly, "No Data"
       miCount = 0
    Else
        If txtProgram <> "(Enter Program Name)" And txtProgram <> "" Then
            cboProgram.AddItem txtProgram
            txtProgram = ""
        Else
            MsgBox "(Enter Program Name) is an invalid entry", vbOKOnly, "Invalid Entry"
            txtProgram = ""
        End If
    End If
    
    'lblProgramCount = cboProgram.ListCount & " entries"

If cboProgram.ListCount = 0 Then
    lblProgramCount = ""
ElseIf cboProgram.ListCount = 1 Then
    lblProgramCount = cboProgram.ListCount & " entry"
ElseIf cboProgram.ListCount > 1 Then
    lblProgramCount = cboProgram.ListCount & " entries"
End If
    
    If cboProgram.ListCount > 20 Then
        MsgBox "The Program List contains more than 20 entries." & vbCrLf & vbCrLf & _
        "It is suggested using the Delete Selected Program button you reduce the list to no more than 20 entries", _
        vbOKOnly, "Program List Exceeds 20 Entries"
    End If
        
    frmStaff!shpProgram.Visible = False
    frmStaff!shpHost.Visible = False
    
    frmStaff!shpProgram.Visible = False
    cmdLog.SetFocus
    
End Sub

Private Sub cmdProgramRemove_Click()
  miCount = miCount + 1
  Dim iResponse As Integer
  'remove selected line
  If cboProgram.ListIndex >= 0 Then 'if a line is selected
     iResponse = MsgBox("Remove '" & cboProgram.List(cboProgram.ListIndex) & "' from the program list?", vbYesNo, "Confirm Delete")
     If iResponse = vbYes Then
         cboProgram.RemoveItem cboProgram.ListIndex
         cboProgram.ListIndex = -1
     End If
  Else
     MsgBox "No program has been selected to delete.", vbOKOnly, "No Selection"
     miCount = 0
  End If
  
  'lblProgramCount = cboProgram.ListCount & " entries"
  
If cboProgram.ListCount = 0 Then
    lblProgramCount = ""
ElseIf cboProgram.ListCount = 1 Then
    lblProgramCount = cboProgram.ListCount & " entry"
ElseIf cboProgram.ListCount > 1 Then
    lblProgramCount = cboProgram.ListCount & " entries"
End If
  
  cmdLog.SetFocus
  
End Sub

Private Sub Form_Activate()

    If giHost = 1 Then
        cmdProgramAdd.SetFocus
        frmStaff!shpProgram.Visible = True
        frmStaff!shpHost.Visible = False
        
    ElseIf giHost = 2 Then
        cmdHostAdd.SetFocus
        frmStaff!shpHost.Visible = True
        frmStaff!shpProgram.Visible = False
        
    ElseIf giHost = 3 Then
        txtProgram.SetFocus
        
    ElseIf giHost = 4 Then
        txtHost.SetFocus
    Else
        txtProgram.SetFocus
        frmStaff!shpProgram.Visible = False
        frmStaff!shpHost.Visible = False
    End If
        
    'lblProgramCount = cboProgram.ListCount & " entries"
    
If cboProgram.ListCount = 0 Then
    lblProgramCount = ""
ElseIf cboProgram.ListCount = 1 Then
    lblProgramCount = cboProgram.ListCount & " entry"
ElseIf cboProgram.ListCount > 1 Then
    lblProgramCount = cboProgram.ListCount & " entries"
End If
    
    'lblHostCount = cboYou.ListCount & " entries"
    
If cboYou.ListCount = 0 Then
   lblHostCount = ""
ElseIf cboYou.ListCount = 1 Then
    lblHostCount = cboYou.ListCount & " entry"
ElseIf cboYou.ListCount > 1 Then
    lblHostCount = cboYou.ListCount & " entries"
End If
    
    If cboProgram.ListCount > 20 Then
        MsgBox "The Program List contains more than 20 entries." & vbCrLf & vbCrLf & _
        "It is suggested using the Delete Selected Program button you reduce the list to no more than 20 entries", _
        vbOKOnly, "Program List Exceeds 20 Entries"
    End If
        
    If cboYou.ListCount > 20 Then
        MsgBox "The Host List contains more than 20 entries." & vbCrLf & vbCrLf & _
        "It is suggested using the Delete Selected Host button you reduce the list to no more than 20 entries", _
        vbOKOnly, "Host List Exceeds 20 Entries"
    End If
    
End Sub

Private Sub Form_Click()
    cboYou.ListIndex = -1 'removes highlight from List line
    cboProgram.ListIndex = -1
    txtProgram.SetFocus
End Sub

Private Sub Form_Load()
    miCount = 0
    Dim sTemp As String
    Dim sTemp2 As String
     
On Error GoTo HandleErrors
    cboProgram.Clear
    cboYou.Clear
    
    Open "Program.dat" For Input As #1 'opens text file & restores data to screen
    Do Until EOF(1)
    Line Input #1, sTemp
    
If sTemp <> "" Then
    cboProgram.AddItem sTemp
End If
    
    Loop
    Close #1
    
    Open "Host.dat" For Input As #2 'opens text file & restores data to screen
    Do Until EOF(2)
    Line Input #2, sTemp2
    
If sTemp2 <> "" Then
    cboYou.AddItem sTemp2
End If

    Loop
    Close #2
    
    Exit Sub
HandleErrors:
    Close #1
    Close #2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
 On Error GoTo HandleErrors
 
    If miCount <> 0 Then
        Dim i As Integer
        Dim sTemp As String
        Dim sTemp2 As String

        Open "Program.dat" For Output As #1 'saves logged data
        For i = 0 To cboProgram.ListCount - 1
        Print #1, cboProgram.List(i)
        Next
        Close #1

        Open "Host.dat" For Output As #2 'saves logged data
        For i = 0 To cboYou.ListCount - 1
        Print #2, cboYou.List(i)
        Next
        Close #2
        '------------
        'updates frmLog Program & Host Lists (and frmPlanner lists)
        frmPlanner!cboProgram.Clear
        frmPlanner!cboYou.Clear

        Open "Program.dat" For Input As #1
        Do Until EOF(1)
        Line Input #1, sTemp
        frmPlanner!cboProgram.AddItem sTemp
        Loop
        Close #1

        Open "Host.dat" For Input As #2 'opens text file & restores data to screen
        Do Until EOF(2)
        Line Input #2, sTemp2
        
'If sTemp2 <> "" Then
    frmPlanner!cboYou.AddItem sTemp2
'End If
        
        Loop
        Close #2

    End If

    '------------
    Dim you, program As String 'maintains current entries in planner page Host & Program print text boxes
    Open "PrintHost.dat" For Input As #403
    Input #403, you, program
    Close #403
    frmPlanner!cboYou = you
    frmPlanner!cboProgram = program

    Exit Sub

HandleErrors:
    Dim sHostName As String
    Dim sProgramName As String
    sHostName = "(Enter Host's Name)"
    sProgramName = "(Enter Program Name)"

    Open "PrintHost.dat" For Output As #403
    Write #403, sHostName, sProgramName
    Close #403
    
    frmPlanner!cboYou = sHostName
    frmPlanner!cboProgram = sProgramName

    Close #1
    Close #2
End Sub
    
Private Sub Frame1_Click()
    cboYou.ListIndex = -1 'removes highlight from List line
    cboProgram.ListIndex = -1
    txtProgram.SetFocus
End Sub

Private Sub Frame2_Click()
    cboYou.ListIndex = -1 'removes highlight from List line
    cboProgram.ListIndex = -1
    txtHost.SetFocus
End Sub

Private Sub txtProgram_Change()
    If Len(txtProgram) > 35 Then
       txtProgram = ""
    End If
End Sub

Private Sub txtProgram_Click()
    cboYou.ListIndex = -1 'removes highlight from List line
    cboProgram.ListIndex = -1
End Sub

Private Sub txtProgram_DblClick()
    If txtProgram = "" Then
        txtProgram = gsProgram
    End If
End Sub

Private Sub txtHost_Change()
    If Len(txtHost) > 26 Then
        txtHost = ""
    End If
End Sub

Private Sub txtHost_Click()
    cboYou.ListIndex = -1 'removes highlight from List line
    cboProgram.ListIndex = -1
End Sub

Private Sub txtHost_DblClick()
    If txtHost = "" Then
        txtHost = gsHost
    End If
End Sub
