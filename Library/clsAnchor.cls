VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Const WM_WINDOWPOSCHANGED As Long = &H47
Private Const WM_DESTROY As Long = &H2
Private Const WM_NCDESTROY As Long = &H82

Private Const SPI_GETNONCLIENTMETRICS As Long = 41

Private Const LF_FACESIZE As Long = 32

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long

Dim mColCtls As Collection
Dim AttachedHwnd As Long

Public Enum eAnchorTypes
    eLeft = 1
    eTop = 2
    eRight = 4
    eBottom = 8
    eAll = 15
    eNone = 0
End Enum

Private Sub Class_Initialize()
    '-- Initialise the collection
    Set mColCtls = New Collection
End Sub

Private Sub Class_Terminate()
    '-- Make sure we unsubclass
    DetachWind
    '-- Destroy the collection
    Set mColCtls = Nothing
End Sub

Private Function ISubclass_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean) As Long
    Select Case uMsg
        Case WM_WINDOWPOSCHANGED
            '-- The window position has changed
            If mColCtls.Count > 0 Then
                '-- Make sure we don't bother to loop if there's nothing to loop through
                Dim i As Long
                Dim relRect As RECT
                Dim WndRect As RECT
                Dim CtlRect As RECT
                Dim AnchorType As eAnchorTypes
                Dim Left As Long, Top As Long, width As Long, height As Long
                
                GetClientRect hwnd, WndRect
                
                For i = 1 To mColCtls.Count
                    '-- Retrieve the control's anchor type
                    AnchorType = GetProp(mColCtls(i), "AnchorType")
                    '-- Get its original rect relative to the parent
                    relRect = GetRectFromWndProperties(mColCtls(i))
                    '-- Get its new rect
                    GetWindowRect mColCtls(i), CtlRect
                    
                    '-- Initialise our vars
                    Left = relRect.Left
                    Top = relRect.Top
                    width = CtlRect.Right - CtlRect.Left
                    height = CtlRect.Bottom - CtlRect.Top
                    
                    '-- If it's anchored right
                    If (AnchorType And eAnchorTypes.eRight) = eAnchorTypes.eRight Then
                        '-- If it's also anchored left
                        If (AnchorType And eAnchorTypes.eLeft) = eAnchorTypes.eLeft Then
                            '-- Adjust the width
                            width = (WndRect.Right - WndRect.Left) - (relRect.Right + relRect.Left)
                        Else
                            '-- Otherwise just adjust the left position
                            Left = ((WndRect.Right - WndRect.Left) - relRect.Right) - (CtlRect.Right - CtlRect.Left)
                        End If
                    End If
                    
                    '-- If it's anchored to the bottom
                    If (AnchorType And eAnchorTypes.eBottom) = eAnchorTypes.eBottom Then
                        '-- If it's also anchored to the top
                        If (AnchorType And eAnchorTypes.eTop) = eAnchorTypes.eTop Then
                            '-- Adjust the height
                            height = (WndRect.Bottom - WndRect.Top) - (relRect.Bottom + relRect.Top)
                        Else
                            '-- Otherwise we just need to adjust its top position
                            Top = ((WndRect.Bottom - WndRect.Top) - relRect.Bottom) - (CtlRect.Bottom - CtlRect.Top)
                        End If
                    End If
                    
                    '-- Change the window's position
                    SetWindowPos mColCtls(i), 0, Left, Top, width, height, 0
                    
                    '-- Get the pointer to the label object, if there is one
                    Dim pLabel As Long
                    pLabel = GetProp(mColCtls(i), "pLabel")
                    If pLabel <> 0 Then
                        '-- If we have a pointer
                        Dim mLabel As Label
                        
                        '-- Get a Label object from the pointer
                        CopyMemory mLabel, pLabel, 4&
                        
                        '-- Get the alignment data
                        Dim Align As AlignmentConstants, Side As AlignConstants
                        Align = GetProp(mColCtls(i), "LabelAlign")
                        Side = GetProp(mColCtls(i), "LabelSide")
                        
                        '-- Move the label to the correct place
                        Select Case Align
                            Case vbLeftJustify
                                Select Case Side
                                    Case vbAlignLeft
                                        mLabel.Left = (Left * Screen.TwipsPerPixelX) - mLabel.width
                                        mLabel.Top = Top * Screen.TwipsPerPixelY
                                    Case vbAlignRight
                                        mLabel.Left = ((Left + width) * Screen.TwipsPerPixelX) - mLabel.width
                                        mLabel.Top = Top * Screen.TwipsPerPixelY
                                    Case vbAlignTop
                                        mLabel.Left = Left * Screen.TwipsPerPixelX
                                        mLabel.Top = (Top * Screen.TwipsPerPixelY) - mLabel.height
                                    Case vbAlignBottom
                                        mLabel.Left = Left * Screen.TwipsPerPixelX
                                        mLabel.Top = (Top + height) * Screen.TwipsPerPixelY
                                End Select
                            Case vbRightJustify
                                Select Case Side
                                    Case vbAlignLeft
                                        mLabel.Left = (Left * Screen.TwipsPerPixelX) - mLabel.width
                                        mLabel.Top = (Top + height) * Screen.TwipsPerPixelY - mLabel.height
                                    Case vbAlignRight
                                        mLabel.Left = (Left + width) * Screen.TwipsPerPixelX
                                        mLabel.Top = (Top + height) * Screen.TwipsPerPixelY
                                    Case vbAlignTop
                                        mLabel.Left = (Left + width) * Screen.TwipsPerPixelX - mLabel.width
                                        mLabel.Top = (Top * Screen.TwipsPerPixelY) - mLabel.height
                                    Case vbAlignBottom
                                        mLabel.Left = (Left + width) * Screen.TwipsPerPixelX - mLabel.width
                                        mLabel.Top = (Top + height) * Screen.TwipsPerPixelY
                                End Select
                            Case vbCenter
                                Select Case Side
                                    Case vbAlignLeft
                                        mLabel.Left = (Left * Screen.TwipsPerPixelX) - mLabel.width
                                        mLabel.Top = (Top + height \ 2) * Screen.TwipsPerPixelY - mLabel.height \ 2
                                    Case vbAlignRight
                                        mLabel.Left = (Left + width) * Screen.TwipsPerPixelX
                                        mLabel.Top = (Top + height \ 2) * Screen.TwipsPerPixelY - mLabel.height \ 2
                                    Case vbAlignTop
                                        mLabel.Left = (Left + width \ 2) * Screen.TwipsPerPixelX - mLabel.width \ 2
                                        mLabel.Top = Top * Screen.TwipsPerPixelY - mLabel.height
                                    Case vbAlignBottom
                                        mLabel.Left = (Left + width \ 2) * Screen.TwipsPerPixelX - mLabel.width \ 2
                                        mLabel.Top = (Top + height) * Screen.TwipsPerPixelY
                                End Select
                        End Select
                        
                        '-- Clean up
                        ZeroMemory mLabel, 4&
                    End If
                Next i
            End If
        
        Case WM_DESTROY, WM_NCDESTROY
            '-- Make sure we unsubclass
            DetachWind
    End Select
End Function

Public Sub AttachWind(ByVal hwnd As Long)
    '-- Make sure we're not trying to subclass something
    '   when we're already Attached to another window
    If AttachedHwnd = 0 Then
        '-- Subclass the window
        modSubclass.SubWnd hwnd, Me, modSubclass.NewCSubclass(hwnd), "CAnchor"
        '-- Store its handle for later use
        AttachedHwnd = hwnd
    End If
End Sub

Public Sub DetachWind()
    If Not AttachedHwnd = 0 Then
        '-- Remove the handler
        modSubclass.GetCSubclass(AttachedHwnd).RemoveHandler "CAnchor"
        
        '-- Reset our handle variable
        AttachedHwnd = 0
    End If
End Sub

Public Sub AddCtlLabel(ByVal hwnd As Long, c As Object, ByVal Align As AlignmentConstants, ByVal Side As AlignConstants)
    SetProp hwnd, "pLabel", ObjPtr(c)
    SetProp hwnd, "LabelAlign", Align
    SetProp hwnd, "LabelSide", Side
End Sub

Public Sub RemoveCtlLabel(ByVal hwnd As Long)
    RemoveProp hwnd, "pLabel"
End Sub

Public Sub AddCtl(ByVal hwnd As Long, ByVal AnchorType As eAnchorTypes)
    '-- Set the control's anchor type and add it to the collection
    SetCtlAnchorType hwnd, AnchorType
    mColCtls.Add hwnd, "m" & hwnd
End Sub

Public Sub RemoveCtl(ByVal hwnd As Long)
    '-- Remove the window's properties and remove the control from the collection
    RemoveProp hwnd, "AnchorType"
    RemoveProp hwnd, "relrectLeft"
    RemoveProp hwnd, "relrectTop"
    RemoveProp hwnd, "relrectRight"
    RemoveProp hwnd, "relrectBottom"
    
    mColCtls.Remove "m" & hwnd
End Sub

Public Sub SetCtlAnchorType(ByVal hwnd As Long, ByVal AnchorType As eAnchorTypes)
    '-- Set the anchor type and relative rect position
    SetProp hwnd, "AnchorType", AnchorType
    SetRectAsWndProperties hwnd, GetWindRelPos(hwnd, AttachedHwnd)
End Sub

Public Function GetCtlAnchorType(ByVal hwnd As Long) As eAnchorTypes
    '-- Retrieve the anchor type
    GetCtlAnchorType = GetProp(hwnd, "AnchorType")
End Function

Public Sub RefreshStoredCtlPositions()
    If mColCtls.Count > 0 Then
        Dim i As Long
        
        For i = 1 To mColCtls.Count
            SetRectAsWndProperties mColCtls(i), GetWindRelPos(mColCtls(i), AttachedHwnd)
        Next i
    End If
End Sub

Private Sub SetRectAsWndProperties(ByVal hwnd As Long, r As RECT)
    SetProp hwnd, "relrectLeft", r.Left
    SetProp hwnd, "relrectTop", r.Top
    SetProp hwnd, "relrectRight", r.Right
    SetProp hwnd, "relrectBottom", r.Bottom
End Sub

Private Function GetRectFromWndProperties(ByVal hwnd As Long) As RECT
    Dim r As RECT
    SetRect _
        r, _
        GetProp(hwnd, "relrectLeft"), _
        GetProp(hwnd, "relrectTop"), _
        GetProp(hwnd, "relrectRight"), _
        GetProp(hwnd, "relrectBottom")
    GetRectFromWndProperties = r
End Function


