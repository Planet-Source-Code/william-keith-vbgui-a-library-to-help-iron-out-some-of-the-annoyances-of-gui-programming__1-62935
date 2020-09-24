Attribute VB_Name = "modMain"
Option Explicit

Private Const SPI_GETNONCLIENTMETRICS As Long = 41

Private Const LF_FACESIZE As Long = 32

Public Type RECT
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

Public Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

Public Function GetWindRelPos(ByVal hwndA As Long, ByVal hwndB As Long) As RECT
    Dim r As RECT
    Dim r2 As RECT
    Dim ret As RECT
    Dim ncm As NONCLIENTMETRICS
    Dim ncOffsetY As Long
    Dim ncOffsetX As Long
    
    GetWindowRect hwndA, r
    GetWindowRect hwndB, r2
    
    ncm.cbSize = Len(ncm)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, ncm, 0
    
    ncOffsetY = ncm.iBorderWidth * 5 + ncm.iCaptionHeight
    ncOffsetX = ncm.iBorderWidth * 5
    
    SetRect ret, _
        (r.Left - ncOffsetX) - r2.Left, _
        (r.Top - ncOffsetY) - r2.Top, _
        (r2.Right - ncOffsetX) - r.Right, _
        (r2.Bottom - ncOffsetX) - r.Bottom
    
    GetWindRelPos = ret
End Function

Public Function GetCaptionOffset() As Integer
    Dim ncm As NONCLIENTMETRICS
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, ncm, 0
    GetCaptionOffset = ncm.iBorderWidth * 5 + ncm.iCaptionHeight
End Function

Public Function GetBorderOffset() As Integer
    Dim ncm As NONCLIENTMETRICS
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, ncm, 0
    GetBorderOffset = ncm.iBorderWidth * 5
End Function

