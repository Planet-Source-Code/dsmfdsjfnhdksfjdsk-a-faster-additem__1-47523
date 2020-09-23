VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AddItem Competition"
   ClientHeight    =   2760
   ClientLeft      =   1950
   ClientTop       =   2745
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4350
   Begin VB.CommandButton cmdComboAPI 
      Caption         =   "Add items using Win32 API"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton cmdComboVB 
      Caption         =   "Add items using AddItem method"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.ComboBox cmbTest 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''
'    A Faster AddItem                          '
'  This will give you a completly faster way   '
'  of using SendMessage to add 5000 items in   '
'  less than 1 second while the VB add item    '
'   takes over 2-3 seconds. Pretty cool stuff  '
''''''''''''''''''''''''''''''''''''''''''''''''




Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_ADDSTRING = &H143

Private Sub cmdComboAPI_Click()
    Dim strItemText     As String
    Dim intIndex        As Integer
    Dim dblTimer        As Double
    
    cmbTest.Clear
    dblTimer = Timer
    For intIndex = 1 To 5000
        strItemText = "Item number " & CStr(intIndex)
        SendMessage cmbTest.hWnd, CB_ADDSTRING, 0, ByVal strItemText
    Next
    MsgBox Format(Timer - dblTimer, "0.000") & " seconds", , "SendMessage"

End Sub

Private Sub cmdComboVB_Click()
    Dim strItemText     As String
    Dim intIndex        As Integer
    Dim dblTimer        As Double
    
    cmbTest.Clear
    dblTimer = Timer
    For intIndex = 1 To 5000
        strItemText = "item number " & CStr(intIndex)
        cmbTest.AddItem strItemText
    Next
    MsgBox Format(Timer - dblTimer, "0.000") & " seconds" & vbLf & "SendMessage is better, right?", , "AddItem"

End Sub
