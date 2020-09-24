VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMinMaxSize"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const WM_GETMINMAXINFO As Long = &H24

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Implements ISubclass

Dim mHwnd As Long

Public MaxWidth As Long
Public MaxHeight As Long
Public MinWidth As Long
Public minHeight As Long

Public Sub AttachWind(ByVal hwnd As Long)
    '-- Make sure we're not already attached to some other window
    If mHwnd = 0 Then
        mHwnd = hwnd
        '-- Subclass
        modSubclass.SubWnd hwnd, Me, modSubclass.NewCSubclass(mHwnd), "CMinMaxSize"
    End If
End Sub

Public Sub DetachWind()
    '-- Make sure we're not already detached
    If Not mHwnd = 0 Then
        '-- Unsubclass
        modSubclass.GetCSubclass(mHwnd).RemoveHandler "CMinMaxSize"
        mHwnd = 0
    End If
End Sub

Private Function ISubclass_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean) As Long
    If uMsg = WM_GETMINMAXINFO Then
        Dim GMMI As MINMAXINFO
        
        '-- Get our MINMAXINFO
        CopyMemory GMMI, ByVal lParam, ByVal Len(GMMI)
        
        '-- Make sure everything is in order
        If MinWidth > 0 Then GMMI.ptMinTrackSize.x = MinWidth
        If minHeight > 0 Then GMMI.ptMinTrackSize.y = minHeight
        If MaxWidth > 0 Then GMMI.ptMaxTrackSize.x = MaxWidth
        If MaxHeight > 0 Then GMMI.ptMaxTrackSize.y = MaxHeight
        
        '-- Copy our MINMAXINFO back to modify the window position
        CopyMemory ByVal lParam, GMMI, ByVal Len(GMMI)

        bHandled = True
    End If
End Function
