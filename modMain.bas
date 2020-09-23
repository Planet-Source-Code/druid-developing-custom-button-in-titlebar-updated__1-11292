Attribute VB_Name = "modMain"
Option Explicit

'*********************
'* API Declarations  *
'*********************
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, ByVal hmod&, ByVal dwThreadId&) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'*********************
'* Type Declarations *
'*********************
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hwnd As Long
End Type

'*********************
'* Consts            *
'*********************
Const WM_MOVE = &H3
Const WM_SETCURSOR = &H20
Const WM_NCPAINT = &H85
Const WM_COMMAND = &H111

Const SWP_FRAMECHANGED = &H20
Const GWL_EXSTYLE = -20

'*********************
'* Vars              *
'*********************
Private WHook&
Private ButtonHwnd As Long

Public Sub Init()
    'Create the button that is going to be placed in the Titlebar
    ButtonHwnd& = CreateWindowEx(0&, "Button", "i", &H40000000, 50, 50, 14, 14, Form1.hwnd, 0&, App.hInstance, 0&)
    'Show the button cause it´s invisible
    Call ShowWindow(ButtonHwnd&, 1)
    'Initialize the window hooking for the button
    WHook = SetWindowsHookEx(4, AddressOf HookProc, 0, App.ThreadID)
    Call SetWindowLong(ButtonHwnd&, GWL_EXSTYLE, &H80)
    Call SetParent(ButtonHwnd&, GetParent(Form1.hwnd))
End Sub

Public Sub Terminate()
    'Terminate the window hooking
    Call UnhookWindowsHookEx(WHook)
    Call SetParent(ButtonHwnd&, Form1.hwnd)
End Sub

Public Function HookProc&(ByVal nCode&, ByVal wParam&, Inf As CWPSTRUCT)
    Dim FormRect As Rect
    Static LastParam&
    If Inf.hwnd = GetParent(ButtonHwnd&) Then
        If Inf.Message = WM_COMMAND Then
            Select Case LastParam
                'If the LastParam is cmdInTitlebar call the Click-Procedure
                'of the button
                Case ButtonHwnd&: Call Form1.cmdInTitlebar_Click
            End Select
        ElseIf Inf.Message = WM_SETCURSOR Then
            LastParam = Inf.wParam
        End If
        ElseIf Inf.hwnd = Form1.hwnd Then
        If Inf.Message = WM_NCPAINT Or Inf.Message = WM_MOVE Then
            'Get the size of the Form
            Call GetWindowRect(Form1.hwnd, FormRect)
            'Place the button int the Titlebar
            Call SetWindowPos(ButtonHwnd&, 0, FormRect.Right - 75, FormRect.Top + 6, 17, 14, SWP_FRAMECHANGED)
        End If
    End If
End Function
