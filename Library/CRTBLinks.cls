VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRTBLinks"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type

Private Const WM_USER As Long = &H400

Private Const CFE_LINK As Long = &H20

Private Const EM_AUTOURLDETECT As Long = (WM_USER + 91)
Private Const EM_GETEVENTMASK As Long = (WM_USER + 59)
Private Const EM_GETTEXTRANGE As Long = (WM_USER + 75)
Private Const EM_SETEVENTMASK As Long = (WM_USER + 69)

Private Const EN_LINK As Long = &H70B

Private Const ENM_LINK As Long = &H4000000

Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_SETCURSOR As Long = &H20
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long

Implements ISubclass

Dim mHwnd As Long
Dim bDetect As Boolean

Public Event URLMouseDown(ByVal URLText As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Public Event URLMouseUp(ByVal URLText As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
Public Event URLDblClick(ByVal URLText As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)

Public Sub AttachWind(ByVal m_hwnd As Long)
    If mHwnd = 0 Then
        mHwnd = m_hwnd
        modSubclass.SubWnd GetParent(m_hwnd), Me, modSubclass.NewCSubclass(GetParent(mHwnd)), "CRTBLinks"
    End If
End Sub

Public Sub DetachWind()
    If Not mHwnd = 0 Then
        modSubclass.GetCSubclass(GetParent(mHwnd)).RemoveHandler "CRTBLinks"
        mHwnd = 0
    End If
End Sub

Public Property Get hwnd() As Long
    hwnd = mHwnd
End Property

Public Property Let AutoDetectURLs(ByVal Value As Boolean)
    bDetect = Value
    
    SendMessage mHwnd, EM_AUTOURLDETECT, Abs(Value), ByVal 0
    SendMessage mHwnd, EM_SETEVENTMASK, ByVal 0, ByVal SendMessage(mHwnd, EM_GETEVENTMASK, 0, 0) Or ENM_LINK
End Property

Public Property Get AutoDetectURLs() As Boolean
    AutoDetectURLs = bDetect
End Property

Private Sub Class_Terminate()
    DetachWind
End Sub

Private Function ISubclass_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean) As Long
    Dim uHead As NMHDR
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim sText As String
    Dim lLen As Long
    
    Select Case uMsg
    Case WM_NOTIFY
                
        'Ooo! A notify message! Something exciting must be happening...
        'Copy the notification header into our structure from the pointer
        CopyMemory uHead, ByVal lParam, Len(uHead)
        
        'Peek inside the structure
        If (uHead.hWndFrom = mHwnd) And (uHead.code = EN_LINK) Then
            'Yay! Some kind of kinky linky message.
            'Now that we know its a link message, we can copy the whole ENLINK structure
            'into our structure
            CopyMemory eLink, ByVal lParam, Len(eLink)
            
            'Set up out TEXTRANGE struct
            eText.chrg.cpMin = eLink.chrg.cpMin
            eText.chrg.cpMax = eLink.chrg.cpMax
            eText.lpstrText = Space$(1024)
            
            'Tell the RTB to fill out our TEXTRANGE with the text
            lLen = SendMessage(mHwnd, EM_GETTEXTRANGE, 0, eText)
            
            'Trim the text
            sText = Strings.Left$(eText.lpstrText, lLen)
            
            Dim x As Integer, y As Integer
            x = LoWord(eLink.lParam)
            y = HiWord(eLink.lParam)
            
            Select Case eLink.msg
                Case WM_LBUTTONDOWN:    RaiseEvent URLMouseDown(sText, vbLeftButton, eLink.wParam, x, y)
                Case WM_LBUTTONDBLCLK:  RaiseEvent URLDblClick(sText, vbLeftButton, eLink.wParam, x, y)
                Case WM_LBUTTONUP:      RaiseEvent URLMouseUp(sText, vbLeftButton, eLink.wParam, x, y)
                Case WM_RBUTTONDOWN:    RaiseEvent URLMouseDown(sText, vbRightButton, eLink.wParam, x, y)
                Case WM_RBUTTONDBLCLK:  RaiseEvent URLDblClick(sText, vbRightButton, eLink.wParam, x, y)
                Case WM_RBUTTONUP:      RaiseEvent URLMouseUp(sText, vbRightButton, eLink.wParam, x, y)
                Case WM_MBUTTONDOWN:    RaiseEvent URLMouseDown(sText, vbMiddleButton, eLink.wParam, x, y)
                Case WM_MBUTTONDBLCLK:  RaiseEvent URLDblClick(sText, vbMiddleButton, eLink.wParam, x, y)
                Case WM_MBUTTONUP:      RaiseEvent URLMouseUp(sText, vbMiddleButton, eLink.wParam, x, y)
            End Select
            
        End If
        
    End Select
End Function
