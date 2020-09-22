Attribute VB_Name = "Module1"

Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const GWL_WNDPROC = (-4)
Const WM_HOTKEY = &H312

Public Enum ModKeys
  MOD_ALT = &H1
  MOD_CONTROL = &H2
  MOD_SHIFT = &H4
  MOD_WIN = &H8
End Enum

Dim iAtom As Integer
Dim OldProc As Long, hOwner As Long
Public sFile As String

Public Function SetHotKey(hWin As Long, ModKey As ModKeys, vKey As Long) As Boolean
  If hOwner > 0 Then Exit Function
  hOwner = hWin
  iAtom = GlobalAddAtom("MyHotKey")
  SetHotKey = RegisterHotKey(hOwner, iAtom, ModKey, vKey)
  OldProc = SetWindowLong(hOwner, GWL_WNDPROC, AddressOf WndProc)
End Function

Public Sub RemoveHotKey()
  If hOwner = 0 Then Exit Sub
  Call UnregisterHotKey(hOwner, iAtom)
  Call SetWindowLong(hOwner, GWL_WNDPROC, OldProc)
End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If wMsg = WM_HOTKEY And wParam = iAtom Then
     Form1.Show
  Else
     WndProc = CallWindowProc(OldProc, hwnd, wMsg, wParam, lParam)
  End If
End Function

Public Sub AddToLog(sText As String)
   Dim nFile As Integer
   nFile = FreeFile
   Open sFile For Append As #nFile
        Print #nFile, sText
   Close #nFile
End Sub

