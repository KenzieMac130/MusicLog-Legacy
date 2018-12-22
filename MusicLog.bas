Attribute VB_Name = "Module1"
Option Explicit
Public mRunTime As Variant
Public sPlannerProgramTime As String
Public gcMCal1 As Currency
Public gcTotalMinNT As Currency
Public gcTotalMinNA As Currency
Public giExit As Integer 'used in exiting without frmPlanner Form QueryUnload

Public giClockShow As Integer 'prevents log form focus error when clock engaged
                             ' selects between frmLog &frmList when clock closed
Public giPlannedTime As String 'distributes defaulted planned time
Public giStation1 'distributes frmTransmitter station1 call letters
'Public giSave 'saves TimeRemain data when form activated from Log Page
Public giTimeFocus 'sets focus upon return from AddTime,StopWatch,TR Pad to prevent RT Error
'Public giDefaultsFocus 'sets focus upon opening defaults page
Public gipAnnounce 'used to call pAnnounce in frmTimeRemain

Public gcMinA As Currency 'minutes in MusicLineup summary line
Public gcSecA As Integer 'seconds in MusicLineup summary line
Public giSpots As Integer 'counts # of spots for MusicLineup summary line
Public gcMCal3 As Currency 'compares Planning Page total time with Time Remain Page
Public gcMinNA As Currency
Public giCodeRequired As Integer

Public giHost As Integer  'controls saving hosts name
Public giAccess As String 'sets Access Code value
Public giActivity As Integer 'controls recording activity log
Public gsRemain1 As String 'to copy data to Planner Page without ampersand showing &&
Public gsHost As String 'to carry host name to staff page for double-click
Public gsProgram As String 'to carry program name to staff page for double-click
Public giChkPrintXmitter As Integer 'prevents runtime error when xmitter screen cleared
Public giTimesDiffer As Integer 'controls planning and program lineup times differ message
Public giPlannerTxtSpot As Integer 'indicates to TimeRemain page whether Planner page txtSpots contains data
Public gcProgramMinAdd As Integer 'program line printout
Public gcSecPAdd As Integer 'program line printout

'for public version
'Public giOpenCount As Integer ' counts times program opened, and returns to 1 when activity log cleared.

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_SETHORIZONTALEXTENT = &H194 'installs horizontal scroll bar at bottom of list box
Public gStation1Min, gStation1Max, gStation2Min, gStation2Max, gStation3Min, _
        gStation3Max, gStation4Min, gStation4Max As String

Public Annc4 As String 'this series is used in saving annc times to file
Public Annc5 As String
Public Annc6 As String
Public Annc7 As String
Public Annc8 As String
Public Annc9 As String
Public Annc10 As String
Public Annc11 As String

Public F4Link As Integer

Public giTxtCopy As Integer 'allows track box to show in sequence if export txt from frmTimeRemain page to frmPlanner lineup
Public giEntrySave As Integer 'alternates saves for previous save command
Public giMemo As Integer 'prevents requiring password after subsequent openings of frmMenos
Public giNumPages As Integer
