VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMouseEvents"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Const WM_MOUSEHOVER As Long = &H2A1
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEMOVE As Long = &H200
Private Const TME_HOVER As Long = &H1
Private Const TME_LEAVE As Long = &H2

Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long

Public Event MouseHover()
Public Event MouseLeave()
Public Event MouseWheel(ByVal Dist As Long, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)

Dim dwHoverTime As Long
Dim mHwnd As Long

Public Property Get HoverTime() As Long
    HoverTime = dwHoverTime
End Property

Public Property Let HoverTime(ByVal value As Long)
    dwHoverTime = value
    RequestTracking
End Property

Public Sub AttachWind(ByVal m_hwnd As Long)
    If mHwnd = 0 Then
        mHwnd = m_hwnd
        modSubclass.SubWnd m_hwnd, Me, modSubclass.NewCSubclass(mHwnd), "CMouseEvents"
        RequestTracking
    End If
End Sub

Public Sub DetachWind()
    If Not mHwnd = 0 Then
        modSubclass.GetCSubclass(mHwnd).RemoveHandler "CMouseEvents"
        mHwnd = 0
    End If
End Sub

Private Sub Class_Initialize()
    dwHoverTime = 1
End Sub

Private Sub Class_Terminate()
    DetachWind
End Sub

Private Function ISubclass_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean) As Long
    Static bOver As Boolean
    
    Select Case uMsg
        Case WM_MOUSEHOVER
            If Not bOver Then RaiseEvent MouseHover
            bOver = True
        Case WM_MOUSELEAVE
            RaiseEvent MouseLeave
            bOver = False
        Case WM_MOUSEMOVE
            RequestTracking
            
        Case WM_MOUSEWHEEL
            RaiseEvent MouseWheel(HiWord(wParam) \ 120, LoWord(wParam), LoWord(lParam), HiWord(lParam))
    End Select
End Function

Private Sub RequestTracking()
    Dim TME As tagTRACKMOUSEEVENT
    With TME
        .cbSize = Len(TME)
        .dwFlags = TME_HOVER Or TME_LEAVE
        .dwHoverTime = dwHoverTime
        .hwndTrack = mHwnd
    End With
    TrackMouseEvent TME
End Sub
