VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SB_VERT = 1
Private Const SB_HORZ = 0
Private Const SB_LEFT As Long = 6
Private Const SB_RIGHT As Long = 7
Private Const SB_LINELEFT As Long = 0
Private Const SB_LINERIGHT As Long = 1
Private Const SB_THUMBPOSITION As Long = 4
Private Const SB_THUMBTRACK As Long = 5
Private Const SB_PAGELEFT As Long = 2
Private Const SB_PAGERIGHT As Long = 3
Private Const SB_LINEUP As Long = 0
Private Const SB_LINEDOWN As Long = 1
Private Const SB_PAGEUP As Long = 2
Private Const SB_PAGEDOWN As Long = 3

Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS As Long = &H4

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long

Implements ISubclass

Public Event VScroll()
Public Event HScroll()

Dim mHMin As Long
Dim mHMax As Long
Dim mHValue As Long
Dim mVMin As Long
Dim mVMax As Long
Dim mVValue As Long
Dim mHwnd As Long

Private Sub SetRange(ByVal nBar As Long, ByVal Min As Long, ByVal Max As Long, Optional ByVal PageSize As Long = -1)
    Dim SI As SCROLLINFO
    
    '-- Set the SI struct up
    With SI
        .cbSize = Len(SI)
        .fMask = SIF_RANGE Or IIf(PageSize > -1, SIF_PAGE, 0)
        .nMax = Max
        .nMin = Min
        .nPage = PageSize
    End With
    
    '-- Set the range
    SetScrollInfo mHwnd, nBar, SI, True
End Sub

Private Sub GetRange(ByVal nBar As Long, ByRef Min As Long, ByRef Max As Long)
    Dim SI As SCROLLINFO
    
    '-- Set the SI struct up
    With SI
        .cbSize = Len(SI)
        .fMask = SIF_RANGE
        .nMax = Max
        .nMin = Min
        
        '-- Get the range
        GetScrollInfo mHwnd, nBar, SI
        
        '-- "Return" the values
        Min = .nMin
        Max = .nMax
    End With
End Sub

Private Function GetPageSize(ByVal nBar As Long) As Long
    Dim SI As SCROLLINFO
    
    '-- Set the SI struct up
    With SI
        .cbSize = Len(SI)
        .fMask = SIF_PAGE
    End With
    '-- Get the page size
    GetScrollInfo mHwnd, nBar, SI
    
    '-- Return it
    GetPageSize = SI.nPage
End Function

Private Sub SetPageSize(ByVal nBar As Long, ByVal Value As Long)
    Dim SI As SCROLLINFO
    
    '-- Set the SI struct up
    With SI
        .cbSize = Len(SI)
        .fMask = SIF_PAGE
        .nPage = Value
    End With
    '-- Set the page size
    SetScrollInfo mHwnd, nBar, SI, True
End Sub

Public Property Let HPageSize(ByVal Value As Long)
    '-- Set the page size
    SetPageSize SB_HORZ, Value
End Property

Public Property Get HPageSize() As Long
    '-- Get the page size
    HPageSize = GetPageSize(SB_HORZ)
End Property

Public Property Let VPageSize(ByVal Value As Long)
    '-- Set the page size
    SetPageSize SB_VERT, Value
End Property

Public Property Get VPageSize() As Long
    '-- Get the page size
    VPageSize = GetPageSize(SB_VERT)
End Property

Public Property Let HMax(ByVal Value As Long)
    '-- Set the H-max value
    mHMax = Value
    '-- Set the range
    SetRange SB_HORZ, mHMin, mHMax
End Property

Public Property Get HMax() As Long
    Dim Min As Long, Max As Long
    '-- Get the range and return the max
    GetRange SB_HORZ, Min, Max
    HMax = Max
End Property

Public Property Let VMax(ByVal Value As Long)
    '-- Set the V-max value
    mVMax = Value
    '-- Set the range
    SetRange SB_VERT, mVMin, mVMax
End Property

Public Property Get VMax() As Long
    Dim Min As Long, Max As Long
    '-- Get the range and return the max
    GetRange SB_VERT, Min, Max
    VMax = Max
End Property

Public Property Get hwnd() As Long
    '-- Return currently attached the hwnd
    hwnd = mHwnd
End Property

Public Sub ShowHBar(ByVal bShow As Boolean)
    '-- Show the horizontal bar
    ShowScrollBar hwnd, SB_HORZ, CLng(bShow) * -1
End Sub

Public Sub ShowVBar(ByVal bShow As Boolean)
    '-- Show the vertical bar
    ShowScrollBar hwnd, SB_VERT, CLng(bShow) * -1
End Sub

Private Sub Class_Terminate()
    '-- Clean up
    DetachWind
End Sub

Public Sub AttachWind(ByVal hwnd As Long)
    '-- Make sure we're not already attached to some other window
    If mHwnd = 0 Then
        mHwnd = hwnd
        '-- Subclass
        modSubclass.SubWnd hwnd, Me, modSubclass.NewCSubclass(mHwnd), "CScrollbar"
    End If
End Sub

Public Sub DetachWind()
    '-- Make sure we're not already detached
    If Not mHwnd = 0 Then
        '-- Hide the scrollbars
        ShowHBar False
        ShowVBar False
        '-- Unsubclass
        modSubclass.GetCSubclass(mHwnd).RemoveHandler "CScrollbar"
        mHwnd = 0
    End If
End Sub

Public Property Get HPos() As Long
    Dim Pos As Long
    '-- Get our scroll position
    Pos = GetScrollPos(mHwnd, SB_HORZ)
    '-- Amount that can be scrolled = max - page size + 1, therefore we interpolate as follows:
    Pos = Pos + (Pos / (HMax - HPageSize + 1) * HPageSize)
    '-- Adjust for the extra 1 that comes out at the max end of the range by interpolating again
    Pos = Pos - Pos / HMax
    '-- Return the position
    HPos = Pos
End Property

Public Property Let HPos(ByVal Value As Long)
    '-- Set the position
    SetScrollPos mHwnd, SB_HORZ, Value, True
End Property

Public Property Get VPos() As Long
    Dim Pos As Long
    Pos = GetScrollPos(mHwnd, SB_VERT)
    Pos = Pos + (Pos / (VMax - VPageSize + 1) * VPageSize)
    Pos = Pos - Pos / VMax
    VPos = Pos
End Property

Public Property Let VPos(ByVal Value As Long)
    SetScrollPos mHwnd, SB_VERT, Value, True
End Property

Private Function ISubclass_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef bHandled As Boolean) As Long
    Dim CurPos As Long
    Dim SI As SCROLLINFO
    
    Select Case uMsg
    Case WM_VSCROLL
        '-- Scrolling vertically
        
        '-- Get the scroll info
        SI.cbSize = Len(SI)
        SI.fMask = SIF_PAGE Or SIF_POS Or SIF_RANGE
        GetScrollInfo hwnd, SB_VERT, SI
        
        '-- See what type of scrolling it is and adjust the scroll pos accordingly
        Select Case LoWord(wParam)
        Case SB_THUMBPOSITION, SB_THUMBTRACK
            SetScrollPos hwnd, SB_VERT, HiWord(wParam), True
        Case SB_LINEUP
            On Error Resume Next
            SetScrollPos hwnd, SB_VERT, SI.nPos - 1, True
        Case SB_LINEDOWN
            On Error Resume Next
            SetScrollPos hwnd, SB_VERT, SI.nPos + 1, True
        Case SB_PAGEUP
            SetScrollPos hwnd, SB_VERT, SI.nPos - SI.nPage, True
        Case SB_PAGEDOWN
            SetScrollPos hwnd, SB_VERT, SI.nPos + SI.nPage, True
        End Select
        
        '-- Raise the VScroll event
        RaiseEvent VScroll
        
    Case WM_HSCROLL
        '-- Scrolling horizontally
        
        '-- Get the scroll info
        SI.cbSize = Len(SI)
        SI.fMask = SIF_PAGE Or SIF_POS Or SIF_RANGE
        GetScrollInfo hwnd, SB_HORZ, SI
        
        '-- See what type of scrolling it is and adjust the scroll pos accordingly
        Select Case LoWord(wParam)
        Case SB_THUMBPOSITION, SB_THUMBTRACK
            SetScrollPos hwnd, SB_HORZ, HiWord(wParam), True
        Case SB_LINELEFT
            On Error Resume Next
            SetScrollPos hwnd, SB_HORZ, SI.nPos - 1, True
        Case SB_LINERIGHT
            On Error Resume Next
            SetScrollPos hwnd, SB_HORZ, SI.nPos + 1, True
        Case SB_PAGELEFT
            SetScrollPos hwnd, SB_HORZ, SI.nPos - SI.nPage, True
        Case SB_PAGERIGHT
            SetScrollPos hwnd, SB_HORZ, SI.nPos + SI.nPage, True
        End Select
        
        '-- Raise the HScroll event
        RaiseEvent HScroll
    End Select
End Function
