VERSION 5.00
Begin VB.UserControl CTabstrip 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "CTabstrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mColTabs As Collection
Dim mSelIndex As Long
Dim bFlatStyle As Boolean

Private Const COLOR_BTNFACE As Long = 15

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, ByRef lpSize As SIZE) As Long

Private Sub UserControl_Initialize()
    Set mColTabs = New Collection
    
    Dim mTab As CTab
    
    Set mTab = New CTab
    mTab.Tag = "pask"
    mTab.Text = "pask"
    mColTabs.Add mTab
    Set mTab = Nothing
    
    Set mTab = New CTab
    mTab.Tag = "paskas"
    mTab.Text = "paskaserfg"
    mColTabs.Add mTab
    Set mTab = Nothing
    
    mSelIndex = 1
End Sub

Private Sub UserControl_Paint()
    PaintControl
End Sub

Private Sub UserControl_Terminate()
    Dim i As Long
    
    For i = mColTabs.Count To 1 Step -1
        'Set mColTabs(i) = Nothing
        mColTabs.Remove i
    Next i
    Set mColTabs = Nothing
End Sub

Private Sub RecalculateTabPlaces()
    Dim CurPos As Long
    
    
End Sub

Private Sub PaintControl()
    Dim hBackDC As Long, hBackBMP As Long
    Dim TP As CTab
    Dim i As Long
    Dim hBrush As Long
    
    hBackDC = CreateCompatibleDC(UserControl.hdc)
    hBackBMP = CreateCompatibleBitmap(UserControl.hdc, UserControl.width \ Screen.TwipsPerPixelX, UserControl.height \ Screen.TwipsPerPixelY)
    DeleteObject SelectObject(hBackDC, hBackBMP)
    
    For i = 1 To mColTabs.Count
        Set TP = mColTabs(i)
        
        If i = mSelIndex Or bFlatStyle Then
            hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
            FillRect hBackDC, TP.GetPosition, hBrush
            DeleteObject hBrush
        End If
        
        Set TP = Nothing
    Next i
    
    BitBlt UserControl.hdc, 0, 0, UserControl.width \ Screen.TwipsPerPixelX, UserControl.height \ Screen.TwipsPerPixelY, hBackDC, 0, 0, vbSrcCopy
    
    DeleteObject hBackBMP
    DeleteDC hBackDC
End Sub
