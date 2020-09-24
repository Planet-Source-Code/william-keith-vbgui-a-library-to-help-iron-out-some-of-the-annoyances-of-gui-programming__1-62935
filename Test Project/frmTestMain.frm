VERSION 5.00
Begin VB.MDIForm frmTestMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9240
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11715
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuNewWindow 
      Caption         =   "New Window"
   End
   Begin VB.Menu mnuotherwind 
      Caption         =   "Show Other Test Window"
   End
End
Attribute VB_Name = "frmTestMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim f As Form
    For Each f In Forms
        Unload f
    Next f
End Sub

Private Sub mnuNewWindow_Click()
    Dim f As New frmTest
    f.Show
End Sub

Private Sub mnuotherwind_Click()
    frmTest2.Show
End Sub
