Attribute VB_Name = "modError"
Option Explicit

Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1
Const ERROR_BAD_USERNAME = 2202&
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Function GetError(error As Long) As String
    Dim Buffer As String
    'Create a string buffer
    Buffer = String(256, Chr(0))
    'Format the message string
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, error, LANG_NEUTRAL, Buffer, 256, ByVal 0&
    GetError = Buffer
End Function
