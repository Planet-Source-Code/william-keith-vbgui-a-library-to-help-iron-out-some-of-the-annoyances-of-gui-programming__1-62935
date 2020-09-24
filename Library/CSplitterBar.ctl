VERSION 5.00
Begin VB.UserControl CSplitterBar 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   MousePointer    =   7  'Size N S
   ScaleHeight     =   4110
   ScaleWidth      =   5310
End
Attribute VB_Name = "CSplitterBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_ASYNCWINDOWPOS As Long = &H4000


Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Event Move(ByVal x As Long, ByVal y As Long)

Dim OldPt As POINTAPI

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbLeftButton = vbLeftButton Then
        Dim pt As POINTAPI
        GetCursorPos pt
        
        If Not (OldPt.x = pt.x And OldPt.y = pt.y) Then
            RaiseEvent Move(pt.x, pt.y)
        End If
    End If
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Sub SizeTopControl(c As Object, ByVal y As Long, ByVal minHeight As Long)
    Dim rc As RECT
    
    On Error Resume Next
    GetWindowRect c.hwnd, rc
    If y - rc.Top >= minHeight Then
        c.height = (y - rc.Top) * Screen.TwipsPerPixelY
    End If
End Sub

Public Sub MoveTopControl(ByVal c As Object, ByVal y As Long)
    Dim pt As POINTAPI
    
    On Error Resume Next
    
    pt.y = y
    ScreenToClient GetParent(c.hwnd), pt
    c.Top = pt.y * Screen.TwipsPerPixelY
End Sub

Public Sub SizeBottomControl(c As Object, ByVal y As Long, ByVal minHeight As Long)
    Dim rc As RECT
    Dim pt As POINTAPI
    Dim newHeight As Long
    
    On Error Resume Next
    GetWindowRect c.hwnd, rc
    
    pt.y = y
    ScreenToClient GetParent(c.hwnd), pt
    newHeight = rc.Bottom - y

    If newHeight >= minHeight Then
        c.Top = pt.y * Screen.TwipsPerPixelY
        c.height = newHeight * Screen.TwipsPerPixelY
    End If
End Sub
