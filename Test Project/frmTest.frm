VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   3375
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   8025
   Begin RichTextLib.RichTextBox text2 
      Height          =   2835
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   5001
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmTest.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Scrollbars"
      Height          =   315
      Left            =   6540
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2835
      IntegralHeight  =   0   'False
      Left            =   6540
      TabIndex        =   1
      Top             =   60
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   3000
      Width           =   6435
   End
   Begin VB.Menu mnupopup 
      Caption         =   "URLpopup"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenLink 
         Caption         =   "&Open link"
      End
      Begin VB.Menu mnyCopyLink 
         Caption         =   "&Copy link"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Anchors controls on form like in VB.NET
Dim mAnchor As CAnchor
'-- API Scrollbars on the form
Dim WithEvents scrl As CScrollbar
Attribute scrl.VB_VarHelpID = -1
'-- Sets the form's minimum and maximum size
Dim MinMaxSize As CMinMaxSize
'-- Handles links in the RichTextBox
Dim WithEvents RtbLinks As CRTBLinks
Attribute RtbLinks.VB_VarHelpID = -1

Dim mURLText As String

Private Sub Command1_Click()
    If Not scrl Is Nothing Then
        '-- Remove scrollbars
        scrl.DetachWind
        Set scrl = Nothing
        
        Command1.Caption = "Show Scrollbars"
    Else
        Set scrl = New CScrollbar
        
        '-- Attatch the CScrollbar class instance to the form
        scrl.AttachWind Me.hWnd
        
        '-- Show the bars
        scrl.ShowVBar True
        scrl.ShowHBar True
        
        '-- Set the bars' page sizes
        scrl.HPageSize = 10
        scrl.VPageSize = 100
        
        Command1.Caption = "Hide Scrollbars"
    End If
End Sub

Private Sub Form_Load()
    Set mAnchor = New VBGUI.CAnchor
    
    '-- Stop VB from screwing with the form size when the form first opens in the MDI container
    Me.Width = 8175
    Me.Height = 3780
    
    '-- Attach the CAnchor to the form
    mAnchor.AttachWind Me.hWnd
    '-- Anchor controls
    '-- This can be done to any control that has a hWnd
    '-- As you can see, there are different anchor types, which mimic those in .NET, and are
    '   relatively self-explanatory
    mAnchor.AddCtl Text1.hWnd, eAnchorTypes.eRight Or eAnchorTypes.eLeft Or eAnchorTypes.eBottom
    mAnchor.AddCtl text2.hWnd, eAnchorTypes.eAll
    mAnchor.AddCtl List1.hWnd, eAnchorTypes.eAll And Not eAnchorTypes.eLeft
    mAnchor.AddCtl Command1.hWnd, eAnchorTypes.eRight Or eAnchorTypes.eBottom
    
    Set MinMaxSize = New CMinMaxSize
    '-- Attach the CMinMaxSize to the form
    MinMaxSize.AttachWind Me.hWnd
    '-- Set the min size and max size
    '-- NOTE: These are in pixels, not twips
    MinMaxSize.MinHeight = Me.Height \ Screen.TwipsPerPixelY
    MinMaxSize.MinWidth = Me.Width \ Screen.TwipsPerPixelX
    
    Set RtbLinks = New CRTBLinks
    '-- Attach the CRTBLinks to the RTB and auto-detect the URLs
    RtbLinks.AttachWind text2.hWnd
    RtbLinks.AutoDetectURLs = True
    
    '-- Add a URL to the text
    text2.Text = "http://wakjah.ath.cx/"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '-- Clean up and detach all the components from the form
    
    If Not scrl Is Nothing Then
        scrl.DetachWind
        Set scrl = Nothing
    End If
    
    mAnchor.DetachWind
    Set mAnchor = Nothing
    
    MinMaxSize.DetachWind
    Set MinMaxSize = Nothing
    
    RtbLinks.DetachWind
    Set RtbLinks = Nothing
End Sub

Private Sub mnuOpenLink_Click()
    MsgBox "Open " & mURLText, , "Open URL"
End Sub

Private Sub mnyCopyLink_Click()
    '-- Put the URL in the clipboard
    Clipboard.Clear
    Clipboard.SetText mURLText
End Sub

Private Sub RtbLinks_URLMouseDown(ByVal URLText As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    '-- Store the URL text for later use
    mURLText = URLText
    If Button = vbRightButton Then
        '-- Popup the menu if the link was right-clicked
        Me.PopupMenu mnupopup, 0, x * Screen.TwipsPerPixelX + 120, y * Screen.TwipsPerPixelY + 120
    End If
End Sub

Private Sub RtbLinks_URLMouseUp(ByVal URLText As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    If Button = vbLeftButton Then
        If Shift And vbKeyControl = vbKeyControl Then
            '-- Open the link if it's left clicked while control is pressed
            mURLText = URLText
            mnuOpenLink_Click
        End If
    End If
End Sub

Private Sub scrl_VScroll()
    '-- Show the scrolling in the richtextbox
    text2.SelStart = Len(text2.Text)
    text2.SelText = vbCrLf & "VScroll: " & scrl.VPos & "    " & scrl.VMax
End Sub

Private Sub scrl_HScroll()
    '-- Show the scrolling in the richtextbox
    text2.SelStart = Len(text2.Text)
    text2.SelText = vbCrLf & "HScroll: " & scrl.HPos & "    " & scrl.HMax
End Sub

