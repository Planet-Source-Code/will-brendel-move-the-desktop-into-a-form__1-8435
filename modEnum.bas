Attribute VB_Name = "modEnum"
' SetParent sets or resets the parent window of another window.
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

' FindWindow is used to find the handle of a window.
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' EnumChildWindows enumerates all the child windows of a window. In this case
' we are enumerating the child windows of Progman.
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

' MoveWindow moves a window :) We are using it to make sure the desktop
' displays properly.
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

' These hold the handles of the four desktop-related windows.
Public Progman, SHELLDLL_DefView, SysListView32, SysHeader32 As Long

' Declare the counter for use with EnumChildProc().
Public iCount As Integer

Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) As Long
    ' I'm sure there's a more efficient way to do this, but
    ' this basically checks the value of the counter and
    ' assigns the variables a value (The windows are found
    ' in order so we can do this without errors (I hope)).
    '
    ' NOTE: We get the handle to Progman in the Form_Load event.
    If iCount = 0 Then
        SHELLDLL_DefView = lhWnd
        iCount = iCount + 1
    ElseIf iCount = 1 Then
        SysListView32 = lhWnd
        iCount = iCount + 1
    ElseIf iCount = 2 Then
        SysHeader32 = lhWnd
        iCount = iCount + 1
    Else
        ' We've found too many windows! End the program
        ' before we do some real damage :)
        MsgBox "Too many windows were found!"
        End
    End If
    
    ' Continue the enumeration
    EnumChildProc = True
End Function


