VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Desktop Capture"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDesktop 
      Height          =   5865
      Left            =   0
      ScaleHeight     =   5805
      ScaleWidth      =   7530
      TabIndex        =   2
      Top             =   450
      Width           =   7590
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "Capture Desktop"
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1440
   End
   Begin VB.CommandButton cmdRelease 
      Caption         =   "Release Desktop"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1500
      TabIndex        =   1
      Top             =   0
      Width           =   1440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MoveWindow is needed for the repositioning of the desktop windows.
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

' GetWindowRect gets the rectangle of a window. This is needed to set the
' heights and widths of the desktop windows.
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

' The RECT data type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' This holds the rectangle of Progman returned by GetWindowRect.
Dim PMRect As RECT
    
Private Sub cmdCapture_Click()
    ' All we do here is set Progman's parent to picDesktop
    ' and then set Progman's children's parents to the window
    ' above it. Maybe looking at the code will better explain it:
    SetParent Progman, picDesktop.hwnd
    SetParent SHELLDLL_DefView, Progman
    SetParent SysListView32, SHELLDLL_DefView
    SetParent SysHeader32, SysListView32
    
    ' Reposition the windows
    RefreshDesktop
    
    ' Enable the release button
    cmdRelease.Enabled = True
    cmdCapture.Enabled = False
End Sub

Private Sub cmdRelease_Click()
    ' This does the same thing that cmdCapture does except
    ' this one sets Progman's parent to 0 (which removes
    ' it from picDesktop.
    SetParent Progman, 0
    SetParent SHELLDLL_DefView, Progman
    SetParent SysListView32, SHELLDLL_DefView
    SetParent SysHeader32, SysListView32
    
    ' Reposition the windows
    RefreshDesktop
    
    ' Enable the Capture button and disable the Release button
    cmdRelease.Enabled = False
    cmdCapture.Enabled = True
End Sub

Private Sub Form_Load()
    ' When we load the program, find the handles of the four windows
    ' having to do with the desktop: Progman, SHELLDLL_DefView,
    ' SysListView32, and SysHeader32.
    '
    ' WARNING: There is an API call named GetDesktopWindow(). This
    ' returns the handle to the actual desktop. DON'T USE IT HERE. It will
    ' crash VB if you try to set that window's parent.
    '
    ' Once we have the handles, we can use SetParent to make picDesktop
    ' the parent window.
    
    ' Set the value of the counter (See EnumChildProc() in module).
    iCount = 0
    
    ' The Progman window is a top-level window so we can just use
    ' FindWindow to get the handle.
    Progman = FindWindow("Progman", "Program Manager")
    
    ' Get the dimesions of Progman
    GetWindowRect Progman, PMRect
    
    ' For the remaining 3 windows under Progman we have to use EnumChildWindows().
    EnumChildWindows Progman, AddressOf EnumChildProc, 0
    
    ' Reposition the windows
    RefreshDesktop
End Sub

Private Sub RefreshDesktop()
    ' This function makes sure the desktop windows
    ' are displayed correctly. I kept getting errors when
    ' I minimized the form, so I made this to fix them.
    ' It also resizes picDesktop to fit the form.
    MoveWindow Progman, 0, 0, PMRect.Right, PMRect.Bottom, 1
    MoveWindow SHELLDLL_ViewDef, 0, 0, PMRect.Right, PMRect.Bottom, 1
    MoveWindow SysListView32, 0, 0, PMRect.Right, PMRect.Bottom, 1
    MoveWindow SysHeader32, 0, 0, PMRect.Right, PMRect.Bottom, 1
    
    ' Resize picDesktop
    picDesktop.Move 5, 400, frmMain.Width - 100, frmMain.Height - 800
End Sub

Private Sub Form_Paint()
    ' Reposition the windows
    RefreshDesktop
End Sub

Private Sub Form_Resize()
    ' Reposition the windows
    RefreshDesktop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' MAKE SURE YOU RELEASE THE WINDOWS!!! Not releasing them will cause
    ' Explorer to close!
    cmdRelease_Click
End Sub
