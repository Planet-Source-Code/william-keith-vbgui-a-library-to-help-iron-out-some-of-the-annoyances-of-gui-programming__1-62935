VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Dim colHandlers As Collection
Dim mHwnd As Long

Public Property Get Handlers() As Collection
    Set Handlers = colHandlers
End Property

Public Property Get hwnd() As Long
    '-- Return our hwnd
    hwnd = mHwnd
End Property

Public Sub AttachWind(ByVal hwnd As Long)
    '-- Make sure we're not already attached to a window
    If mHwnd = 0 Then
        mHwnd = hwnd
        '-- Subclass the window
        modSubclass.SubWnd hwnd, Me, Me
    End If
End Sub

Public Sub DetachWind()
    '-- If we are attached to a window...
    If Not mHwnd = 0 Then
        mHwnd = 0
        '-- Unsubclass the window
        modSubclass.UnSubWnd mHwnd
    End If
End Sub

Public Sub AddHandler(Handler As ISubclass, Optional key)
    '-- Just in case you try to add a handler with a key that is already there...
    On Error GoTo DontBother
    
    '-- Instantiate the collection
    If colHandlers Is Nothing Then Set colHandlers = New Collection
    '-- Add the handler item
    colHandlers.Add Handler, key
    
DontBother:
End Sub

Public Sub RemoveHandler(key)
    On Error GoTo DontBother
    
    '-- Make sure the collection is instantiated
    If Not colHandlers Is Nothing Then
        '-- Remove the item
        colHandlers.Remove key
        '-- If there are no handlers left then we can unsubclass
        If colHandlers.Count = 0 Then
            modSubclass.UnSubWnd mHwnd
        End If
    End If
    
DontBother:
End Sub

Private Sub Class_Terminate()
    '-- Destroy the collection
    If Not colHandlers Is Nothing Then
        Set colHandlers = Nothing
    End If
    
    '-- Detach ourselves from the window
    If Not hwnd = 0 Then DetachWind
End Sub

Private Function ISubclass_WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean) As Long
    Dim Handler As ISubclass
    Dim ret As Long
    
    '-- Call all the handlers
    '-- NOTE: If you want to return a value from the wndproc then it's advisable
    '         to have only one handler for that window because otherwise the return
    '         value will only be that of the last handler in the collection
    If Not colHandlers Is Nothing Then
        For Each Handler In colHandlers
            ret = Handler.WndProc(hwnd, uMsg, wParam, lParam, bHandled)
        Next Handler
    End If
    
    ISubclass_WndProc = ret
End Function
