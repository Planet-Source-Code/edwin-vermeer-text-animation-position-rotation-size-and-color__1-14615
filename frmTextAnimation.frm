VERSION 5.00
Begin VB.Form frmTextAnimation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text animation demo"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "frmTextAnimation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox MessageClick 
      Height          =   5655
      Left            =   4260
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   300
      Width           =   3315
   End
   Begin VB.TextBox MouseDown 
      Height          =   1935
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4020
      Width           =   4095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3300
      Width           =   1275
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3300
      Width           =   1275
   End
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   3255
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5741
      Counter         =   165
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Clicked on message :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Mouse down / up :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
End
Attribute VB_Name = "frmTextAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOk_Click()
  Unload Me
End Sub


Private Sub cmdClear_Click()
  TextAnimation1.RemoveAllMessages
End Sub


Private Sub cmdReset_Click()
  InitializeMessages
End Sub


Private Sub Form_Load()
  InitializeMessages
End Sub


Private Sub TextAnimation1_BeforeDraw(PictureBuffer As PictureBox)
' This will be printed behind the text animation
' All message properties, methods and events will not aply to this text.
  PictureBuffer.CurrentX = (PictureBuffer.ScaleWidth / 2) - 95
  PictureBuffer.CurrentY = 50
  PictureBuffer.FontName = "Arial"
  PictureBuffer.ForeColor = vbWhite
  PictureBuffer.FontSize = 32
  PictureBuffer.Print "Edwin"

End Sub

Private Sub TextAnimation1_AfterDraw(PictureBuffer As PictureBox)
' This will be printed in front of the text animation
' All message properties, methods and events will not aply to this text.
  PictureBuffer.CurrentX = (PictureBuffer.ScaleWidth / 2) - 50
  PictureBuffer.CurrentY = 66
  PictureBuffer.FontName = "Arial"
  PictureBuffer.ForeColor = vbRed
  PictureBuffer.FontSize = 32
  PictureBuffer.Print "Vermeer"

End Sub


Private Sub TextAnimation1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, messages() As String)
' Just do the same as mouse down
  TextAnimation1_MouseDown Button, Shift, x, y, messages

End Sub


Private Sub TextAnimation1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, messages() As String)
' I disabled the mousemove event because this is slowing down the animation considerably as soon as you move the mouse.
' If you want to enable mousemove, then just convert the two comment lines to statements (around line 164)
End Sub


Private Sub TextAnimation1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, messages() As String)
Dim i As Integer
Dim m As Integer
Dim c As String

  ' Display the normal mouse event parameters
  MouseDown = "Button :  " & Button & vbCrLf & _
              "Shift :   " & Shift & vbCrLf & _
              "X :       " & x & vbCrLf & _
              "Y :       " & y & vbCrLf
  
  ' How many messages where unther the mouse cursor?
  On Error Resume Next
  m = UBound(messages)
  If Err.Number = 9 Then
    ' No messages under the mouse cursor
    On Error GoTo 0
    MouseDown = MouseDown & "No messages here"
    MessageClick = ""
  Else
    ' Display all the message ID's that are under the mouse cursor
    On Error GoTo 0
    MouseDown = MouseDown & "Messages here : " & m + 1 & " out of " & TextAnimation1.MessageCount & vbCrLf
    For i = 0 To m
      MouseDown = MouseDown & "Message " & i & " : " & messages(i) & vbCrLf
    Next i
    ' The last message in the passed array will be the topmost. (the one you actually clicked on)
    ' Display all properties of this message.
    c = messages(UBound(messages))
    MessageClick = "MessageFontColorEnd : " & TextAnimation1.MessageFontColorEnd(c) & vbCrLf & _
                   "MessageFontColorStart : " & TextAnimation1.MessageFontColorStart(c) & vbCrLf & _
                   "MessageFontName : " & TextAnimation1.MessageFontName(c) & vbCrLf & _
                   "MessageFontRotationEnd : " & TextAnimation1.MessageFontRotationEnd(c) & vbCrLf & _
                   "MessageFontRotationStart : " & TextAnimation1.MessageFontRotationStart(c) & vbCrLf & _
                   "MessageFontSizeEnd : " & TextAnimation1.MessageFontSizeEnd(c) & vbCrLf & _
                   "MessageFontSizeStart : " & TextAnimation1.MessageFontSizeStart(c) & vbCrLf & _
                   "MessageHeight : " & TextAnimation1.MessageHeight(c) & vbCrLf & _
                   "MessageID : " & TextAnimation1.MessageID(0) & vbCrLf & _
                   "MessageIndex : " & TextAnimation1.MessageIndex(c) & vbCrLf & _
                   "MessageIntervalCount : " & TextAnimation1.MessageIntervalCount(c) & vbCrLf & _
                   "MessageIntervalStart : " & TextAnimation1.MessageIntervalStart(c) & vbCrLf & _
                   "MessageLeftEnd : " & TextAnimation1.MessageLeftEnd(c) & vbCrLf & _
                   "MessageLeftStart : " & TextAnimation1.MessageLeftStart(c) & vbCrLf & _
                   "MessageText : " & TextAnimation1.MessageText(c) & vbCrLf & _
                   "MessageTopEnd : " & TextAnimation1.MessageTopEnd(c) & vbCrLf & _
                   "MessageTopStart : " & TextAnimation1.MessageTopStart(c) & vbCrLf & _
                   "MessageWidth : " & TextAnimation1.MessageWidth(c) & vbCrLf
    ' Change the color of the clicked message
    ' Exept for MessageHeight, MessageID, MessageIndex, and MessageWidth you can also change all the other message properties
    TextAnimation1.MessageFontColorStart(c) = vbRed
    TextAnimation1.MessageFontColorEnd(c) = vbRed
  End If
  
End Sub


Public Sub InitializeMessages()
Dim i As Integer
Dim j As Integer
Dim s As Integer

  TextAnimation1.RemoveAllMessages
  
  ' Animate the 4 messages
  TextAnimation1.AddMessage "test00", "This is test message 1", "Arial", vbBlue, vbWhite, 16, 16, 200, 0, 0, 0, 0, 0, , 0, 300
  TextAnimation1.AddMessage "test01", "This is test message 2", "Arial", vbBlue, vbRed, 16, 16, 200, 0, 200, 0, 0, 0, , 0, 300
  TextAnimation1.AddMessage "test02", "Turn !", "Arial", vbBlue, vbYellow, 24, 24, 100, 100, 100, 100, 0, 360, , 0, 300
  TextAnimation1.AddMessage "test03", "Zoom", "Arial", vbBlue, vbGreen, 1, 100, 100, 0, 0, 170, 0, 0, , 0, 300
  
  ' Followed by the referce animation
  TextAnimation1.AddMessage "test04", "This is test message 1", "Arial", vbWhite, vbBlue, 16, 16, 0, 200, 0, 0, 0, 0, , 300, 300
  TextAnimation1.AddMessage "test05", "This is test message 2", "Arial", vbRed, vbBlue, 16, 16, 0, 200, 0, 200, 0, 0, , 300, 300
  TextAnimation1.AddMessage "test06", "Turn !", "Arial", vbYellow, vbBlue, 24, 24, 100, 100, 100, 100, 0, 360, , 300, 300
  TextAnimation1.AddMessage "test07", "Zoom", "Arial", vbGreen, vbBlue, 100, 1, 0, 100, 170, 0, 0, 0, , 300, 300
  
  ' And another animation cyclus for those 4 messages
  TextAnimation1.AddMessage "test08", "This is test message 1", "Arial", vbBlue, vbWhite, 16, 16, 200, 0, 0, 0, 0, 0, , 600, 300
  TextAnimation1.AddMessage "test09", "This is test message 2", "Arial", vbBlue, vbRed, 16, 16, 200, 0, 200, 0, 0, 0, , 600, 300
  TextAnimation1.AddMessage "test10", "Turn !", "Arial", vbBlue, vbYellow, 24, 24, 100, 100, 100, 100, 0, 360, , 600, 300
  TextAnimation1.AddMessage "test11", "Zoom", "Arial", vbBlue, vbGreen, 1, 100, 100, 0, 0, 170, 0, 0, , 600, 300
  
  ' Followed by the referce animation
  TextAnimation1.AddMessage "test12", "This is test message 1", "Arial", vbWhite, vbBlue, 16, 16, 0, 200, 0, 0, 0, 0, , 900, 300
  TextAnimation1.AddMessage "test13", "This is test message 2", "Arial", vbRed, vbBlue, 16, 16, 0, 200, 0, 200, 0, 0, , 900, 300
  TextAnimation1.AddMessage "test14", "Turn !", "Arial", vbYellow, vbBlue, 24, 24, 100, 100, 100, 100, 0, 360, , 900, 300
  TextAnimation1.AddMessage "test15", "Zoom", "Arial", vbGreen, vbBlue, 100, 1, 0, 100, 170, 0, 0, 0, , 900, 300
  
  ' Just put in 14 * 6 = 84 different messages
  For i = 0 To 13
    s = 0
    For j = 0 To 5
      TextAnimation1.AddMessage "text" & Trim(Str(i)) & "-" & Trim(Str(j)), Left(Trim(Str(i)) & "-" & Trim(Str(j)) & "ABCDEFGHIJKLMNOPQRSTUVWXYZ", i + j + 3), "Arial", vbGreen, vbBlue, 12, 12, 300, -200, i * 16, i * 16, 0, 0, , s, 500
      s = s + TextAnimation1.MessageWidth("text" & Trim(Str(i)) & "-" & Trim(Str(j)))
    Next j
  Next i
  
  ' Set the general animation properties
  TextAnimation1.Counter = 0
  TextAnimation1.CounterMax = 1200
  TextAnimation1.Speed = 20  'This is equal to a refresh rate of 50 frames per second. Making this number smaller will only slow down your application. If you want more speed, then make the all the MessageIntervalCount parameters smaller.
  TextAnimation1.BackColorStart = &H7FFFFF
  TextAnimation1.BackColorEnd = &HFF7F7F
  TextAnimation1.Border = None
  
End Sub

