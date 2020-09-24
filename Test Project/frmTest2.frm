VERSION 5.00
Begin VB.Form frmTest2 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   6990
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   3960
      Width           =   5955
   End
   Begin VB.TextBox Text2 
      Height          =   2775
      Left            =   960
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   5955
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   5955
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3540
      TabIndex        =   5
      Top             =   3660
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mAnchor As CAnchor
Dim MinMaxSize As CMinMaxSize
'-- This gives us access to the mouse events not exposed by VB - hovering and mouse wheel
'   You can see the results by looking in the Immediate window
Dim WithEvents MouseEvents As CMouseEvents
Attribute MouseEvents.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.Height = 4920
    Me.Width = 7095
    
    Set MouseEvents = New CMouseEvents
    MouseEvents.AttachWind Me.hWnd
    
    Set mAnchor = New CAnchor
    With mAnchor
        .AttachWind Me.hWnd
        .AddCtl Text1.hWnd, eAnchorTypes.eAll
        '-- This example takes the first one a little farther by adding labels for controls
        '-- As you can see, labels for controls are added in much the same way as anchored controls are
        '-- All you have to do is tell the library which side of the control the label should be on
        '   and what its alignment should be and it'll handle the rest!
        .AddCtlLabel Text1.hWnd, Label1, vbCenter, vbAlignLeft
        .AddCtl Text2.hWnd, eAnchorTypes.eAll And Not eAnchorTypes.eTop
        .AddCtlLabel Text2.hWnd, Label2, vbLeftJustify, vbAlignLeft
        .AddCtl Text3.hWnd, eAnchorTypes.eAll And Not eAnchorTypes.eTop
        .AddCtlLabel Text3.hWnd, Label3, vbCenter, vbAlignTop
    End With
    
    Set MinMaxSize = New CMinMaxSize
    MinMaxSize.AttachWind Me.hWnd
    MinMaxSize.MinHeight = 303
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mAnchor.DetachWind
    Set mAnchor = Nothing
    
    MinMaxSize.DetachWind
    Set MinMaxSize = Nothing
    
    MouseEvents.DetachWind
    Set MouseEvents = Nothing
End Sub

Private Sub MouseEvents_MouseHover()
    Debug.Print "Mouse Hover"
End Sub

Private Sub MouseEvents_MouseLeave()
    Debug.Print "Mouse Leave"
End Sub

Private Sub MouseEvents_MouseWheel(ByVal Dist As Long, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    Debug.Print "Mouse Wheel: "; Dist; x; y
End Sub
