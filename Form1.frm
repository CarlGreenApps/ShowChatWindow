VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show Chat Window"
   ClientHeight    =   765
   ClientLeft      =   3540
   ClientTop       =   1500
   ClientWidth     =   5925
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const GWL_STYLE = (-16)
Const GWL_EXSTYLE = (-20)
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Const GWL_HWNDPARENT = (-8)
Const WM_COMMAND = &H111
Const MIN_ALL = 419
Const MIN_ALL_UNDO = 416
Const MAX_PATH = 260
Const SW_SHOW = 5
Const SW_RESTORE = 9
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWNORMAL = 1

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type WINDOWPLACEMENT
    Length           As Long
    FLAGS            As Long
    showCmd          As Long
    ptMinPosition    As POINTAPI
    ptMaxPosition    As POINTAPI
    rcNormalPosition As RECT
End Type

Private lpwndpl   As WINDOWPLACEMENT
Private CursorLoc As POINTAPI

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
    (lpPoint As POINTAPI) As Long
    
Private Declare Function WindowFromPoint Lib "user32" _
    (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    
Private Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As Long, lpdwProcessId As Long) As Long
    
Private Declare Function GetWindowPlacement Lib "user32" _
    (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
    
Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Sub Form_Load()
Dim lState     As Long
Dim lHandle    As Long
Dim WindowTitles As Variant, WindowTitle As Variant

'WIN10X64 - XYZ Remote Control - Waiting for your host...
'WIN10X64 - XYZ Remote Control - Connected

WindowTitles = GetWindowTitlesFromFile
If IsArray(WindowTitles) = True Then
    
    For Each WindowTitle In WindowTitles
    Debug.Print "Window Title: " & WindowTitle
    
    lpwndpl.Length = 44
    lHandle = FindWindow(vbNullString, CStr(WindowTitle))
      If lHandle <> 0 Then
        Debug.Print "Found Window: " & WindowTitle
    ' Get the window's state and activate it.
        lState = GetWindowPlacement(lHandle, lpwndpl)
        Select Case lpwndpl.showCmd
            Case SW_SHOWMINIMIZED
                Call ShowWindow(lHandle, SW_RESTORE)
            Case SW_SHOWNORMAL, SW_SHOWMAXIMIZED
                Call ShowWindow(lHandle, SW_SHOW)
        End Select
        Call SetTopMostWindow(lHandle, True) 'force always ontop.
        Call SetTopMostWindow(lHandle, False) 'disable always ontop.
        Call SetForegroundWindow(lHandle) 'this brings chat window back into focus.
      End If
    
    Next WindowTitle
End If


End 'end the program

End Sub

Private Function GetWindowTitlesFromFile() As Variant
Dim TextLine As String
Dim Filename As String
Dim FreeFileNo As Integer
Dim tmpArray() As Variant, Counter As Integer

FreeFileNo = FreeFile

Filename = App.Path & "\" & App.EXEName & ".txt"
If Dir(Filename) <> "" Then
    Open Filename For Input As #FreeFileNo
        Do While Not EOF(FreeFileNo)   ' Loop until end of file.
           Line Input #FreeFileNo, TextLine   ' Read line into variable.
           Debug.Print "Original TextLine: " & TextLine   ' Print to the Immediate window.
           
           'Replace any windows variables with their true values
           ReDim Preserve tmpArray(Counter)
           tmpArray(Counter) = ReplaceSysVars(TextLine)
           Debug.Print "New TextLine: " & tmpArray(Counter)
           Counter = Counter + 1
        Loop
    Close #FreeFileNo   ' Close file.
    
    GetWindowTitlesFromFile = tmpArray
End If
End Function

Public Function ReplaceSysVars(RunLine As String) As String
Dim StartVar As Integer
Dim EndVar As Integer
Dim tmpstr As String
Dim tmpRunLine As String
Dim Fieldname As String, FieldValue As String

On Error GoTo SysVarFailure
StartVar = 1
Debug.Print RunLine
StartVar = InStr(StartVar, RunLine, "%")
Do Until StartVar = 0
    EndVar = InStr(StartVar + 1, RunLine, "%")
    Fieldname = Mid(RunLine, StartVar + 1, EndVar - StartVar - 1)
    FieldValue = Environ$(Fieldname)
    RunLine = Replace(RunLine, "%" & Fieldname & "%", FieldValue)
    Debug.Print RunLine

    StartVar = InStr(StartVar, RunLine, "%")

Loop

ReplaceSysVars = RunLine

Exit Function
SysVarFailure:
    ReplaceSysVars = RunLine
End Function


