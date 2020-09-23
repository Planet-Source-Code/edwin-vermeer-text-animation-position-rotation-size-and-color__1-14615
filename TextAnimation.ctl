VERSION 5.00
Begin VB.UserControl TextAnimation 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   PropertyPages   =   "TextAnimation.ctx":0000
   ScaleHeight     =   3795
   ScaleWidth      =   4500
   ToolboxBitmap   =   "TextAnimation.ctx":001E
   Begin VB.Timer ReDrawTimer 
      Interval        =   20
      Left            =   180
      Top             =   3120
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   1380
      ScaleHeight     =   2535
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2475
      Left            =   780
      ScaleHeight     =   2475
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox picOut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   180
      Width           =   2895
   End
End
Attribute VB_Name = "TextAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long
Private Const SRCCOPY = &HCC0020

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const LF_FACESIZE = 32
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Private Const ANTIALIASED_QUALITY = 4 ' Ensure font edges are smoothed if system is set to smooth font edges
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
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Type TextMessage
    MessageID As String
    MessageText As String
    MessageFontName As String
    MessageFontColorStart As OLE_COLOR
    MessageFontColorEnd As OLE_COLOR
    MessageFontSizeStart As Integer
    MessageFontSizeEnd As Integer
    MessageLeftStart As Integer
    MessageLeftEnd As Integer
    MessageTopStart As Integer
    MessageTopEnd As Integer
    MessageFontRotationStart As Integer
    MessageFontRotationEnd As Integer
    MessageIntervalStart As Long
    MessageIntervalCount As Long
End Type

Private m_messages() As TextMessage

Public Enum SPBorderStyle
    [None] = 0
    [Fixed Single] = 1
End Enum

Dim m_counter As Long
Const m_def_counter = 0
Dim m_counterMax As Long
Const m_def_counterMax = 600
Dim m_backcolorStart As OLE_COLOR
Const m_def_backcolorStart = 8388607  'RGB(255, 255, 127)
Dim m_backcolorEnd As OLE_COLOR
Const m_def_backcolorEnd = 16744319   'RGB(127, 127, 255)
Dim m_Border As Integer
Const m_def_Border = [None]
Dim m_Enabled As Boolean
Const m_def_Enabled = True
Dim m_Speed As Integer
Const m_def_Speed = 20

Event BeforeDraw(PictureBuffer As PictureBox)
Event AfterDraw(PictureBuffer As PictureBox)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, messages() As String)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, messages() As String)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, messages() As String)



Private Sub RedrawTimer_Timer()
On Error Resume Next
Dim RedFrom As Integer
Dim GreenFrom As Integer
Dim BlueFrom As Integer
Dim RedTo As Integer
Dim GreenTo As Integer
Dim BlueTo As Integer
Dim l As Long
Dim j As Long
Dim tLF As LOGFONT
Dim hFnt As Long
Dim hFntOld As Long
Dim lR As Long
Dim iChar As Integer

    m_counter = m_counter + 1
    If Counter > CounterMax Then Counter = 0
    l = BitBlt(picBuffer.hdc, 0, picBuffer.ScaleTop, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBackBuffer.hdc, 0, 0, SRCCOPY)
    RaiseEvent BeforeDraw(picBuffer)
    For j = 0 To MessageCount
      If m_messages(j).MessageIntervalStart <= m_counter And m_messages(j).MessageIntervalStart + m_messages(j).MessageIntervalCount > m_counter Then
        ' The text color
        RedFrom = m_messages(j).MessageFontColorStart And RGB(255, 0, 0)
        GreenFrom = (m_messages(j).MessageFontColorStart And RGB(0, 255, 0)) / 256
        BlueFrom = (m_messages(j).MessageFontColorStart And RGB(0, 0, 255)) / 65536
        RedTo = m_messages(j).MessageFontColorEnd And RGB(255, 0, 0)
        GreenTo = (m_messages(j).MessageFontColorEnd And RGB(0, 255, 0)) / 256
        BlueTo = (m_messages(j).MessageFontColorEnd And RGB(0, 0, 255)) / 65536
        picBuffer.ForeColor = RGB(RedFrom - (RedFrom - RedTo) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount, GreenFrom - (GreenFrom - GreenTo) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount, BlueFrom - (BlueFrom - BlueTo) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount)
        ' The text size
        tLF.lfHeight = MulDiv((m_messages(j).MessageFontSizeStart - (m_messages(j).MessageFontSizeStart - m_messages(j).MessageFontSizeEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount), (GetDeviceCaps(picBuffer.hdc, LOGPIXELSY)), 72)
        ' The rotation of the font
        tLF.lfEscapement = m_messages(j).MessageFontRotationStart - (m_messages(j).MessageFontRotationStart - m_messages(j).MessageFontRotationEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount
        
        ' The text font
        For iChar = 1 To Len(m_messages(j).MessageFontName)
            tLF.lfFaceName(iChar - 1) = CByte(Asc(Mid$(m_messages(j).MessageFontName, iChar, 1)))
        Next iChar
        ' Other font properties (for now default)
        tLF.lfItalic = picBuffer.Font.Italic
        If (picBuffer.Font.Bold) Then
            tLF.lfWeight = FW_BOLD
        Else
            tLF.lfWeight = FW_NORMAL
        End If
        tLF.lfUnderline = picBuffer.Font.Underline
        tLF.lfStrikeOut = picBuffer.Font.Strikethrough
        tLF.lfCharSet = picBuffer.Font.Charset
        tLF.lfQuality = ANTIALIASED_QUALITY
        ' Print the text at the right location
        hFnt = CreateFontIndirect(tLF)
        If (hFnt <> 0) Then
          hFntOld = SelectObject(picBuffer.hdc, hFnt)
          lR = TextOut(picBuffer.hdc, m_messages(j).MessageLeftStart - (m_messages(j).MessageLeftStart - m_messages(j).MessageLeftEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount, m_messages(j).MessageTopStart - (m_messages(j).MessageTopStart - m_messages(j).MessageTopEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount, m_messages(j).MessageText, lstrlen(m_messages(j).MessageText))
          SelectObject picBuffer.hdc, hFntOld
          DeleteObject hFnt
        End If
      End If
    Next j
    RaiseEvent AfterDraw(picBuffer)
    l = BitBlt(picOut.hdc, 0, picOut.ScaleTop, picOut.ScaleWidth, picOut.ScaleHeight, picBuffer.hdc, 0, 0, SRCCOPY)
    picOut.Refresh

End Sub


Private Sub picOut_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim messages() As String
  BuildAray x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, messages
  RaiseEvent MouseDown(Button, Shift, x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, messages)
End Sub

Private Sub picOut_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim messages() As String
'  BuildAray x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, messages
'  RaiseEvent MouseMove(Button, Shift, x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, messages)
End Sub

Private Sub picOut_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim messages() As String
  BuildAray x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, messages
  RaiseEvent MouseUp(Button, Shift, x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, messages)
End Sub

Private Sub BuildAray(x As Single, y As Single, messages() As String)
Dim j As Integer
Dim t As Integer

  For j = 0 To MessageCount
    If m_messages(j).MessageIntervalStart <= m_counter And m_messages(j).MessageIntervalStart + m_messages(j).MessageIntervalCount > m_counter Then
      t = m_messages(j).MessageLeftStart - (m_messages(j).MessageLeftStart - m_messages(j).MessageLeftEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount
      If x >= t And x <= t + MessageWidth(j) Then
        t = m_messages(j).MessageTopStart - (m_messages(j).MessageTopStart - m_messages(j).MessageTopEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount
        If y >= t And y <= t + MessageHeight(j) Then
          On Error Resume Next
          ReDim Preserve messages(UBound(messages) + 1) As String
          If Err.Number = 9 Then ReDim Preserve messages(0)
          On Error GoTo 0
          messages(UBound(messages)) = MessageID(j)
        End If
      End If
    End If
  Next j
  
End Sub


Private Function GradientBackground(picBox As PictureBox)
Dim RedFrom As Integer
Dim GreenFrom As Integer
Dim BlueFrom As Integer
Dim RedTo As Integer
Dim GreenTo As Integer
Dim BlueTo As Integer
Dim rgncnt As Integer
Dim iheight As Long

  RedFrom = m_backcolorStart And RGB(255, 0, 0)
  GreenFrom = (m_backcolorStart And RGB(0, 255, 0)) / 256
  BlueFrom = (m_backcolorStart And RGB(0, 0, 255)) / 65536
  RedTo = m_backcolorEnd And RGB(255, 0, 0)
  GreenTo = (m_backcolorEnd And RGB(0, 255, 0)) / 256
  BlueTo = (m_backcolorEnd And RGB(0, 0, 255)) / 65536
    
  For rgncnt = 1 To 256
    picBox.Line (-1, rgncnt * (Int(picBox.ScaleHeight / 256) + 1))-(picBox.ScaleWidth, rgncnt * (Int(picBox.ScaleHeight / 256) + 1) + (Int(picBox.ScaleHeight / 256) + 1)), RGB(RedTo - ((RedTo - RedFrom) * (rgncnt / 256)), GreenTo - ((GreenTo - GreenFrom) * (rgncnt / 256)), BlueTo - ((BlueTo - BlueFrom) * (rgncnt / 256))), BF
  Next rgncnt
    
End Function



'---------------------------------------------------------------------------
' Usercontrol events
'---------------------------------------------------------------------------

Private Sub UserControl_Initialize()
Dim iLine As Integer
Dim x As Variant
    
    UserControl.ScaleMode = vbPixels
    
    picBuffer.ScaleMode = vbPixels
    picBuffer.ForeColor = vbWhite
    picBuffer.BackColor = vbBlack
    picBuffer.AutoRedraw = True
    picBuffer.Visible = False
    ReDrawTimer.Enabled = True
    
    On Error Resume Next
    x = UserControl.Parent
    If Err.Number = 398 Then
      ' We are in design mode
      AddMessage "design1", "Edwin", "Arial", vbBlue, vbYellow, 24, 24, 100, 100, 100, 100, 0, 360, , 0, 300
      AddMessage "design2", "Vermeer", "Arial", vbBlue, vbGreen, 1, 100, 100, 0, 0, 170, 0, 0, , 0, 300
      AddMessage "design3", "Edwin", "Arial", vbYellow, vbBlue, 24, 24, 100, 100, 100, 100, 0, 360, , 300, 300
      AddMessage "design4", "Vermeer", "Arial", vbGreen, vbBlue, 100, 1, 0, 100, 170, 0, 0, 0, , 300, 300
      AddMessage "design5", "If you speak dutch, then please visit my homepage at www.beursmonitor.com", "Brush Script MT", vbGreen, vbWhite, 32, 32, 300, -900, 0, 0, 0, 0, , 0, 600
    End If
    
End Sub


Private Sub UserControl_Show()
    
    GradientBackground picBackBuffer

End Sub



Private Sub UserControl_Resize()
    
    picBackBuffer.Left = 0
    picBackBuffer.Top = 0
    picBackBuffer.Height = UserControl.ScaleHeight
    picBackBuffer.Width = UserControl.ScaleWidth
    
    picOut.Left = 0
    picOut.Top = 0
    picOut.Height = UserControl.ScaleHeight
    picOut.Width = UserControl.ScaleWidth

    picBuffer.Left = 0
    picBuffer.Top = 0
    picBuffer.Height = UserControl.ScaleHeight
    picBuffer.Width = UserControl.ScaleWidth
      
    GradientBackground picBackBuffer
    
End Sub



'---------------------------------------------------------------------------
' Executing Methods
'---------------------------------------------------------------------------

Public Sub AddMessage( _
       ByVal MessageID As String, _
       Optional ByVal MessageText As String, _
       Optional ByVal MessageFontName As String, _
       Optional ByVal MessageFontColorStart As OLE_COLOR, _
       Optional ByVal MessageFontColorEnd As OLE_COLOR, _
       Optional ByVal MessageFontSizeStart As Integer, _
       Optional ByVal MessageFontSizeEnd As Integer, _
       Optional ByVal MessageLeftStart As Integer, _
       Optional ByVal MessageLeftEnd As Integer, _
       Optional ByVal MessageTopStart As Integer, _
       Optional ByVal MessageTopEnd As Integer, _
       Optional ByVal MessageFontRotationStart As Integer, _
       Optional ByVal MessageFontRotationEnd As Integer, _
       Optional ByVal BeforeMessageID As Variant, _
       Optional ByVal MessageIntervalStart As Long = 0, _
       Optional ByVal MessageIntervalCount As Long = 0 _
       )

Dim iM As Long
Dim i As Long


   If IsMissing(MessageText) Then MessageText = "Edwin Vermeer"
   If IsMissing(MessageFontName) Then MessageFontName = "Ariel"
   If IsMissing(MessageFontColorStart) Then MessageFontColorStart = vbBlue
   If IsMissing(MessageFontColorEnd) Then MessageFontColorEnd = vbWhite
   If IsMissing(MessageFontSizeStart) Then MessageFontSizeStart = 8
   If IsMissing(MessageFontSizeEnd) Then MessageFontSizeEnd = 16
   If IsMissing(MessageLeftStart) Then MessageLeftStart = picBuffer.ScaleWidth
   If IsMissing(MessageLeftEnd) Then MessageLeftEnd = 0
   If IsMissing(MessageTopStart) Then MessageTopStart = picBuffer.ScaleHeight
   If IsMissing(MessageTopEnd) Then MessageTopEnd = 0
   If IsMissing(MessageFontRotationStart) Then MessageFontRotationStart = 0
   If IsMissing(MessageFontRotationEnd) Then MessageFontRotationEnd = 0
   If IsMissing(MessageIntervalStart) Then MessageIntervalStart = 0
   If IsMissing(MessageIntervalCount) Then MessageIntervalCount = CounterMax
   
   ReDim Preserve m_messages(0 To MessageCount + 1) As TextMessage
   If Not (IsMissing(BeforeMessageID)) Then
      iM = MessageIndex(BeforeMessageID)
      If (iM > -1) Then ' insert
         For i = MessageCount To iM + 1 Step -1
            LSet m_messages(i) = m_messages(i - 1)
         Next i
      End If
    Else
      iM = MessageCount
    End If
    With m_messages(iM)
       .MessageID = MessageID
       .MessageText = MessageText
       .MessageFontName = MessageFontName
       .MessageFontColorStart = MessageFontColorStart
       .MessageFontColorEnd = MessageFontColorEnd
       .MessageFontSizeStart = MessageFontSizeStart
       .MessageFontSizeEnd = MessageFontSizeEnd
       .MessageLeftStart = MessageLeftStart
       .MessageLeftEnd = MessageLeftEnd
       .MessageTopStart = MessageTopStart
       .MessageTopEnd = MessageTopEnd
       .MessageFontRotationStart = MessageFontRotationStart * 10
       .MessageFontRotationEnd = MessageFontRotationEnd * 10
       .MessageIntervalStart = MessageIntervalStart
       .MessageIntervalCount = MessageIntervalCount
    End With

End Sub


Public Sub RemoveMessage(ByVal MessageID As Variant)
Dim iM As Integer
Dim i As Long
   
   iM = MessageIndex(MessageID)
   If (iM > -1) Then
      If MessageCount > 0 Then
         For i = iM To MessageCount - 1
             LSet m_messages(i) = m_messages(i + 1)
         Next i
         ReDim Preserve m_messages(0 To MessageCount - 1) As TextMessage
      End If
   End If
   
End Sub

Public Sub RemoveAllMessages()
  ReDim m_messages(0) As TextMessage
End Sub

'---------------------------------------------------------------------------
' Getting and Setting the properties
'---------------------------------------------------------------------------
Private Sub UserControl_InitProperties()

    m_backcolorStart = m_def_backcolorStart
    m_backcolorEnd = m_def_backcolorEnd
    m_Border = m_def_Border
    m_Enabled = m_def_Enabled
    m_counter = m_def_counter
    m_counterMax = m_def_counterMax
    m_Speed = m_def_Speed
    
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_backcolorStart = PropBag.ReadProperty("BackColorStart", m_def_backcolorStart)
    m_backcolorEnd = PropBag.ReadProperty("BackColorEnd", m_def_backcolorEnd)
    m_Border = PropBag.ReadProperty("Border", m_def_Border)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_counter = PropBag.ReadProperty("Counter", m_def_counter)
    m_counterMax = PropBag.ReadProperty("CounterMax", m_def_counterMax)
    m_Speed = PropBag.ReadProperty("Speed", m_def_Speed)

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_backcolorStart, m_def_backcolorStart)
    Call PropBag.WriteProperty("BackColor", m_backcolorEnd, m_def_backcolorEnd)
    Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Counter", m_counter, m_def_counter)
    Call PropBag.WriteProperty("CounterMax", m_counterMax, m_def_counterMax)
    Call PropBag.WriteProperty("Speed", m_Speed, m_def_Speed)

End Sub


' .Counter
Public Property Get Counter() As Long
Attribute Counter.VB_ProcData.VB_Invoke_Property = "General"
    Counter = m_counter
End Property
Public Property Let Counter(ByVal New_Counter As Long)
    m_counter = New_Counter
    PropertyChanged "Counter"
End Property


' .CounterMax
Public Property Get CounterMax() As Long
Attribute CounterMax.VB_ProcData.VB_Invoke_Property = "General"
    CounterMax = m_counterMax
End Property
Public Property Let CounterMax(ByVal New_CounterMax As Long)
    m_counterMax = New_CounterMax
    PropertyChanged "CounterMax"
End Property


' .BackColorStart
Public Property Get BackColorStart() As OLE_COLOR
    BackColorStart = m_backcolorStart
End Property
Public Property Let BackColorStart(ByVal New_BackColorStart As OLE_COLOR)
    m_backcolorStart = New_BackColorStart
    PropertyChanged "BackColorStart"
    UserControl_Resize
End Property


' .BackColorEnd
Public Property Get BackColorEnd() As OLE_COLOR
    BackColorEnd = m_backcolorEnd
End Property
Public Property Let BackColorEnd(ByVal New_BackColorEnd As OLE_COLOR)
    m_backcolorEnd = New_BackColorEnd
    PropertyChanged "BackColorEnd"
    UserControl_Resize
End Property


' . Border
Public Property Get Border() As SPBorderStyle
    Border = m_Border
End Property
Public Property Let Border(ByVal New_Border As SPBorderStyle)
    m_Border = New_Border
    PropertyChanged "Border"
    UserControl.BorderStyle = m_Border
End Property


' .Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    ReDrawTimer = m_Enabled
End Property


' .Speed
Public Property Get Speed() As Long
Attribute Speed.VB_ProcData.VB_Invoke_Property = "General"
    Speed = m_Speed
End Property
Public Property Let Speed(ByVal New_Speed As Long)
    m_Speed = New_Speed
    PropertyChanged "Speed"
    ReDrawTimer.Interval = m_Speed
End Property


' .MessageCount
Public Property Get MessageCount() As Long
On Error Resume Next
    MessageCount = UBound(m_messages)
    If Err.Number <> 0 Then MessageCount = -1
End Property


' .MessageIndex(MessageID)
Public Property Get MessageIndex(ByVal MessageID As Variant) As Integer
Dim iM As Integer
Dim iIndex As Integer
    
    iIndex = -1
    If (IsNumeric(MessageID)) Then
        iIndex = CInt(MessageID)
    Else
        If MessageCount > 0 Then
           For iM = 0 To MessageCount
              If (m_messages(iM).MessageID = MessageID) Then
                  iIndex = iM
                  Exit For
              End If
           Next iM
        Else
           MessageIndex = -1
        End If
    End If
    If (iIndex > -1) And (iIndex <= MessageCount) Then
        MessageIndex = iIndex
    Else
        MessageIndex = -1
    End If
    
End Property


' .MessageID(MessageIndex)
Public Property Get MessageID(ByVal iMessage As Long) As String
   If (iMessage > -1) And (iMessage <= MessageCount) Then
      MessageID = m_messages(iMessage).MessageID
   End If
End Property


' .MessageWidth(MessageID)
Public Property Get MessageWidth(ByVal MessageID As Variant) As Integer
Dim j As Integer
Dim w As Integer

    j = MessageIndex(MessageID)
    If j < 0 Then
      MessageWidth = 0
    Else
      picBuffer.FontName = m_messages(j).MessageFontName
      w = (m_messages(j).MessageFontSizeStart - (m_messages(j).MessageFontSizeStart - m_messages(j).MessageFontSizeEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount) - 2
      If w < 1 Then w = 1
      picBuffer.FontSize = w
      MessageWidth = picBuffer.TextWidth(m_messages(j).MessageText)
    End If
    
End Property

' .MessageHeight(MessageID)
Public Property Get MessageHeight(ByVal MessageID As Variant) As Integer
Dim j As Integer
    
    j = MessageIndex(MessageID)
    picBuffer.FontName = m_messages(j).MessageFontName
    picBuffer.FontSize = (m_messages(j).MessageFontSizeStart - (m_messages(j).MessageFontSizeStart - m_messages(j).MessageFontSizeEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount)
    MessageHeight = picBuffer.TextHeight(m_messages(j).MessageText)
    
End Property


' .MessageText (MessageID)
Public Property Get MessageText(ByVal MessageID As Variant) As String
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageText = m_messages(j).MessageText
End Property
Public Property Let MessageText(ByVal MessageID As Variant, ByVal New_MessageText As String)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageText = New_MessageText
End Property


' .MessageFontName (MessageID)
Public Property Get MessageFontName(ByVal MessageID As Variant) As String
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontName = m_messages(j).MessageFontName
End Property
Public Property Let MessageFontName(ByVal MessageID As Variant, ByVal New_MessageFontName As String)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontName = New_MessageFontName
End Property


' .MessageFontColorStart (MessageID)
Public Property Get MessageFontColorStart(ByVal MessageID As Variant) As OLE_COLOR
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontColorStart = m_messages(j).MessageFontColorStart
End Property
Public Property Let MessageFontColorStart(ByVal MessageID As Variant, ByVal New_MessageFontColorStart As OLE_COLOR)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontColorStart = New_MessageFontColorStart
End Property


' .MessageFontColorEnd (MessageID)
Public Property Get MessageFontColorEnd(ByVal MessageID As Variant) As OLE_COLOR
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontColorEnd = m_messages(j).MessageFontColorEnd
End Property
Public Property Let MessageFontColorEnd(ByVal MessageID As Variant, ByVal New_MessageFontColorEnd As OLE_COLOR)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontColorEnd = New_MessageFontColorEnd
End Property


' .MessageFontSizeStart (MessageID)
Public Property Get MessageFontSizeStart(ByVal MessageID As Variant) As Integer
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontSizeStart = m_messages(j).MessageFontSizeStart
End Property
Public Property Let MessageFontSizeStart(ByVal MessageID As Variant, ByVal New_MessageFontSizeStart As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontSizeStart = New_MessageFontSizeStart
End Property


' .MessageFontSizeEnd (MessageID)
Public Property Get MessageFontSizeEnd(ByVal MessageID As Variant) As Integer
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontSizeEnd = m_messages(j).MessageFontSizeEnd
End Property
Public Property Let MessageFontSizeEnd(ByVal MessageID As Variant, ByVal New_MessageFontSizeEnd As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontSizeEnd = New_MessageFontSizeEnd
End Property


' .MessageLeftStart (MessageID)
Public Property Get MessageLeftStart(ByVal MessageID As Variant) As Integer
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageLeftStart = m_messages(j).MessageLeftStart
End Property
Public Property Let MessageLeftStart(ByVal MessageID As Variant, ByVal New_MessageLeftStart As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageLeftStart = New_MessageLeftStart
End Property


' .MessageLeftEnd (MessageID)
Public Property Get MessageLeftEnd(ByVal MessageID As Variant) As Integer
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageLeftEnd = m_messages(j).MessageLeftEnd
End Property
Public Property Let MessageLeftEnd(ByVal MessageID As Variant, ByVal New_MessageLeftEnd As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageLeftEnd = New_MessageLeftEnd
End Property


' .MessageTopStart (MessageID)
Public Property Get MessageTopStart(ByVal MessageID As Variant) As Integer
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageTopStart = m_messages(j).MessageTopStart
End Property
Public Property Let MessageTopStart(ByVal MessageID As Variant, ByVal New_MessageTopStart As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageTopStart = New_MessageTopStart
End Property


' .MessageTopEnd (MessageID)
Public Property Get MessageTopEnd(ByVal MessageID As Variant) As Integer
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageTopEnd = m_messages(j).MessageTopEnd
End Property
Public Property Let MessageTopEnd(ByVal MessageID As Variant, ByVal New_MessageTopEnd As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageTopEnd = New_MessageTopEnd
End Property


' .MessageFontRotationStart (MessageID)
Public Property Get MessageFontRotationStart(ByVal MessageID As Variant) As Integer
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontRotationStart = m_messages(j).MessageFontRotationStart
End Property
Public Property Let MessageFontRotationStart(ByVal MessageID As Variant, ByVal New_MessageFontRotationStart As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontRotationStart = New_MessageFontRotationStart
End Property


' .MessageFontRotationEnd (MessageID)
Public Property Get MessageFontRotationEnd(ByVal MessageID As Variant) As Integer
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontRotationEnd = m_messages(j).MessageFontRotationEnd
End Property
Public Property Let MessageFontRotationEnd(ByVal MessageID As Variant, ByVal New_MessageFontRotationEnd As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontRotationEnd = New_MessageFontRotationEnd
End Property


' .MessageIntervalStart (MessageID)
Public Property Get MessageIntervalStart(ByVal MessageID As Variant) As Long
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageIntervalStart = m_messages(j).MessageIntervalStart
End Property
Public Property Let MessageIntervalStart(ByVal MessageID As Variant, ByVal New_MessageIntervalStart As Long)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageIntervalStart = New_MessageIntervalStart
End Property


' .MessageIntervalCount (MessageID)
Public Property Get MessageIntervalCount(ByVal MessageID As Variant) As Long
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageIntervalCount = m_messages(j).MessageIntervalCount
End Property
Public Property Let MessageIntervalCount(ByVal MessageID As Variant, ByVal New_MessageIntervalCount As Long)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageIntervalCount = New_MessageIntervalCount
End Property


