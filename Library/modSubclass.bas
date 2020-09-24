Attribute VB_Name = "modSubclass"
Option Explicit

Private Const GWL_WNDPROC As Long = -4

Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim colSubclassPool As Collection

Public Function NewCSubclass(ByVal hwnd As Long) As CSubclass
    If ColContains("m" & hwnd) Then
        '-- If it already contains the item then simply return the item already there
        Set NewCSubclass = colSubclassPool("m" & hwnd)
    Else
        '-- Otherwise...
        Dim CS As CSubclass
        
        '-- Make sure the collection is instantiated
        If colSubclassPool Is Nothing Then Set colSubclassPool = New Collection
        
        '-- Create a new CSubclass
        Set CS = New CSubclass
        On Error Resume Next
        '-- Add it to the collection
        colSubclassPool.Add CS, "m" & hwnd
        Set CS = Nothing
        
        '-- Attach it to the window
        With colSubclassPool("m" & hwnd)
            .AttachWind hwnd
        End With
    End If
End Function

Public Function GetCSubclass(ByVal hwnd As Long) As CSubclass
    '-- Return the CSubclass object associated with the window
    On Error Resume Next
    Set GetCSubclass = colSubclassPool("m" & hwnd)
End Function

Private Function ColContains(key) As Boolean
    On Error GoTo DontBother
    
    '-- Try to get the object from the collection
    '-- If it's not there then there will be an error and we'll jump to DontBother
    Dim o As Object
    
    Set o = colSubclassPool(key)
    ColContains = True
    
    Exit Function
    
DontBother:
    ColContains = False
End Function

Public Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function LoWord(ByVal DWord As Long) As Integer
    If (DWord And &H8000&) = 0 Then
        LoWord = DWord And &HFFFF&
    Else
        LoWord = DWord Or &HFFFF0000
    End If
End Function

'-- Function to get the address of a function in a long
Private Function GetAddr(ByVal Addr As Long) As Long: GetAddr = Addr: End Function

Public Sub SubWnd(ByVal hwnd As Long, Handler As ISubclass, MainSubclassObj As CSubclass, Optional ByVal HandlerKey)
    Dim pOldWndProc As Long

    '-- Make sure we actually have a window here...
    If IsWindow(hwnd) <> 0 Then
        If Not (GetWindowLong(hwnd, GWL_WNDPROC) = GetAddr(AddressOf fnWndProc)) Then
            '-- If it's not already subclassed, then subclass it
            pOldWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf fnWndProc)
            SetProp hwnd, "pOldWndProc", pOldWndProc
            SetProp hwnd, "pHandler", ObjPtr(Handler)
            SetProp hwnd, "pMainHandler", ObjPtr(MainSubclassObj)
        Else
            '-- If, however, it is subclassed then...
            Dim pMainHandler As Long
            Dim oMainHandler As CSubclass
            
            '-- Get the pointer to the window's CSubclass object
            pMainHandler = GetProp(hwnd, "pMainHandler")
            If pMainHandler <> 0 Then
                '-- If it's not 0 then...
                If Not colSubclassPool Is Nothing Then
                    '-- Make sure the collection is instantiated, then get the CSubclass object
                    Set oMainHandler = GetCSubclass(hwnd)
                    If Not oMainHandler Is Nothing Then
                        '-- Make sure we have a CSubclass object, and if so, add the handler to
                        '   its handlers collection
                        oMainHandler.AddHandler Handler, HandlerKey
                    End If
                    Set oMainHandler = Nothing
                End If
            End If
        End If
    End If
End Sub

Public Sub UnSubWnd(ByVal hwnd As Long)
    '-- Make sure we have a window here
    If IsWindow(hwnd) <> 0 Then
        Dim pOldWndProc As Long
        
        pOldWndProc = GetProp(hwnd, "pOldWndProc")
        If pOldWndProc <> 0 Then
            '-- Set the proc back to the old window proc
            SetWindowLong hwnd, GWL_WNDPROC, pOldWndProc
            
            '-- Remove the properties
            RemoveProp hwnd, "pOldWndProc"
            RemoveProp hwnd, "pHandler"
            RemoveProp hwnd, "pMainHandler"
        End If
        
        '-- If the CSubclass item is in the collection...
        If ColContains("m" & hwnd) Then
            '-- Remove it from the collection
            colSubclassPool.Remove "m" & hwnd
            '-- If there are none left in the collection...
            If colSubclassPool.Count = 0 Then
                '-- Destroy the collection
                Set colSubclassPool = Nothing
            End If
        End If
    End If
End Sub

Private Function fnWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pOldWndProc As Long
    Dim pHandler As Long
    Dim oHandler As ISubclass
    Dim ret As Long
    Dim bHandled As Boolean
    
    '-- Get the old proc and pointer to the handler object
    pOldWndProc = GetProp(hwnd, "pOldWndProc")
    pHandler = GetProp(hwnd, "pHandler")
    
    '-- Copy the handler into memory and call its sub-handlers
    If pHandler <> 0 Then
        CopyMemory oHandler, pHandler, 4&
        ret = oHandler.WndProc(hwnd, uMsg, wParam, lParam, bHandled)
        ZeroMemory oHandler, 4&
    End If
    
    '-- If the handlers have chosen not to 'eat' the message then...
    If Not bHandled Then
        If pOldWndProc <> 0 Then
            '-- If we have an old window proc pointer then call the old window proc
            ret = CallWindowProc(pOldWndProc, hwnd, uMsg, wParam, lParam)
        Else
            '-- Else call the default window proc
            ret = DefWindowProc(hwnd, uMsg, wParam, lParam)
        End If
    End If
    
    fnWndProc = ret
End Function
