Attribute VB_Name = "Module"
Option Explicit

Public Type Variables
  X              As Double
  Y              As Double
  Height         As Double
  Width          As Double
  GravityX       As Double
  GravityY       As Double
  clip           As Long
  Neutral_Key    As Long
  gamestarted    As Boolean
  key            As String
  GravityReverse As Boolean
  Speed          As Double
  GravPwr        As Double
End Type
Public P1 As Variables
Public Misc As Variables
Public Fuel As Variables
Public pause As Variables
Public fuelspot As Variables

Public level As Integer

Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long

Declare Function Beep Lib "kernel32" _
(ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" _
(ByVal vkey As Long) As Long

Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long) As Long

Declare Function LineTo Lib "gdi32" _
(ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long) As Long

Declare Function SetPixelV Lib "gdi32" _
(ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal crColor As Long) As Long

Sub load()
  
  Form1.Player1.Picture = LoadPicture(App.Path & "\Lunar Ship.bmp")
  MakeMaskFor Form1.Player1
  With P1
    .GravityY = 0
    .Height = Form1.Player1.Height / 2
    .Width = Form1.Player1.Width / 6
    .X = 10
    .Y = 30
  End With

  Form1.Height = 8775
  Form1.Width = 10185


  For Misc.Neutral_Key = 0 To 255 Step 1
    If GetAsyncKeyState(Misc.Neutral_Key) Then
    End If
  Next Misc.Neutral_Key
  
  Fuel.X = 1024
  Form1.FuelBar.Max = Fuel.X
    
  level = 1
  Select_Level

  Misc.gamestarted = False
  
End Sub
 
Sub Main()
P1.Speed = Form1.GravScrl.Value * 0.01

If GetAsyncKeyState(vbKeyP) Then
  Misc.gamestarted = False
  Form1.GravScrl.Visible = True
End If
  

If GetAsyncKeyState(vbKeySpace) Then
  Misc.gamestarted = True
  Form1.GravScrl.Visible = False
End If



Form1.FuelBar.Value = Fuel.X
If Misc.gamestarted = True Then
  With P1
    Select Case Fuel.X
    Case Is > 1
      If GetAsyncKeyState(vbKeyUp) Then
        Fuel.X = Fuel.X - 1
        If .GravityY > -P1.Speed * 100 Then
          .GravityY = .GravityY - P1.Speed
        Else
          .GravityY = -P1.Speed * 100
        End If
      End If
    
      Select Case True
        Case GetAsyncKeyState(vbKeyRight)
          Fuel.X = Fuel.X - 1
          If .GravityX < P1.Speed * 100 Then
            .GravityX = .GravityX + P1.Speed
          Else
            .GravityX = P1.Speed * 100
          End If
        Case GetAsyncKeyState(vbKeyLeft)
          Fuel.X = Fuel.X - 1
          If .GravityX > -P1.Speed * 100 Then
            .GravityX = .GravityX - P1.Speed
          Else
            .GravityX = -P1.Speed * 100
          End If
      End Select
      Case Else
        Fuel.X = 0
    End Select
    
    Select Case False
      Case GetAsyncKeyState(vbKeyUp) And Fuel.X > 1
        If .GravityY < P1.Speed * 100 Then
          .GravityY = .GravityY + P1.Speed
        Else
          .GravityY = P1.Speed * 100
        End If
    End Select
    
    Select Case False
      Case GetAsyncKeyState(vbKeyRight And vbKeyLeft) And Fuel.X > 1
        Select Case .GravityX
          Case Is > 0
            .GravityX = .GravityX - P1.Speed * 0.1
            If .GravityX < 0 Then
              .GravityX = 0
            End If
          Case Is < 0
            .GravityX = .GravityX + P1.Speed * 0.1
            If .GravityX > 0 Then
              .GravityX = 0
            End If
        End Select
    End Select
   

    If Misc.GravityReverse = True Then
      .Y = .Y - .GravityY
      .X = .X + .GravityX
    Else
      .Y = .Y + .GravityY
      .X = .X + .GravityX
    End If
    
    If .Y + P1.Height < Form1.ScaleTop Then
      .Y = Form1.ScaleTop + Form1.ScaleHeight
    End If
    If .Y > Form1.ScaleTop + Form1.ScaleHeight Then
      .Y = Form1.ScaleTop - P1.Height
    End If
    
    If .X + .Width < Form1.ScaleLeft Then
      .X = Form1.ScaleLeft + Form1.ScaleWidth
    End If
    If .X > Form1.ScaleLeft + Form1.ScaleWidth Then
      .X = Form1.ScaleLeft - .Width
    End If
    
      If Form1.Point(P1.X, P1.Y + .Height / 2) = vbBlue Or Form1.Point(P1.X + .Width / 2, P1.Y + .Height / 2) = vbBlue Or Form1.Point(P1.X + P1.Width / 2, P1.Y) = vbBlue Or Form1.Point(P1.X + .Width / 2, P1.Y + .Height) = vbBlue Or Form1.Point(P1.X + .Width, P1.Y + .Height / 2) = vbBlue Then
        Died
      End If
      
      If Form1.Point(P1.X, P1.Y + .Height / 2) = RGB(0, 255, 255) Or Form1.Point(P1.X + .Width / 2, P1.Y + .Height / 2) = RGB(0, 255, 255) Or Form1.Point(P1.X + P1.Width / 2, P1.Y) = RGB(0, 255, 255) Or Form1.Point(P1.X + .Width, P1.Y + .Height / 2) = RGB(0, 255, 255) Then
        Died
      ElseIf Form1.Point(P1.X + P1.Width / 2, P1.Y + P1.Height) = RGB(0, 255, 255) Then
        If GetAsyncKeyState(vbKeySpace) Then
          P1.GravityY = -P1.Speed
          If Fuel.X < 1024 Then
            Fuel.X = Fuel.X + 1
          End If
        Else
          Died
        End If
      End If
      
      If Form1.Point(P1.X, P1.Y + .Height / 2) = vbGreen Or Form1.Point(P1.X + .Width / 2, P1.Y + .Height / 2) = vbGreen Or Form1.Point(P1.X + P1.Width / 2, P1.Y) = vbGreen Or Form1.Point(P1.X + .Width, P1.Y + .Height / 2) = vbGreen Then
          Died
        Else
        If Form1.Point(P1.X + P1.Width / 2, P1.Y + P1.Height) = vbGreen Then
          If GetAsyncKeyState(vbKeySpace) Then
            Win
          Else
            Died
          End If
        End If
      End If
      
    End With
  End If
    
  Game_Text_and_Blt
End Sub

Sub Win()
  MsgBox "Winner"
  If Misc.GravityReverse = True Then
  MsgBox "Gravity Is Back To Normal"
  End If
  P1.X = 10
  P1.Y = 30
  P1.GravityX = 0
  P1.GravityY = 0
  Fuel.X = Form1.FuelBar.Max
  Misc.GravityReverse = False
  level = level + 1
  Select_Level
  Form1.GravScrl.Visible = True
  Misc.gamestarted = False
End Sub

Sub Died()
  MsgBox "you died"
  P1.X = 10
  P1.Y = 30
  P1.GravityX = 0
  P1.GravityY = 0
  Fuel.X = Form1.FuelBar.Max
  Form1.GravScrl.Visible = True
  Misc.gamestarted = False
End Sub

Sub Select_Level()
  Select Case level
    Case 1
      Form1.Picture = LoadPicture(App.Path & "\Level.bmp")
    Case 2
      Form1.Picture = LoadPicture(App.Path & "\Level2.bmp")
    Case 3
      Misc.GravityReverse = True
      MsgBox "Gravity Has Been Reversed"
      Form1.Picture = LoadPicture(App.Path & "\Level3.bmp")
    Case 4
      Form1.Picture = LoadPicture(App.Path & "\Level4.bmp")
    Case 5
      Form1.Picture = LoadPicture(App.Path & "\Level5.bmp")
    Case 6
      Form1.Picture = LoadPicture(App.Path & "\Level6.bmp")
    Case 7
      Form1.Picture = LoadPicture(App.Path & "\Level7.bmp")
    Case 8
      Misc.GravityReverse = True
      MsgBox "Gravity Has Been Reversed"
      Form1.Picture = LoadPicture(App.Path & "\Level8.bmp")
    Case Else
      MsgBox "You Beat The Game!!!"
      level = 1
      Select_Level
  End Select
End Sub

Sub Game_Text_and_Blt()
  Form1.Caption = "Level:" & level & " of 8" & "  " & "Fuel Left:" & Fuel.X & "  " & "Speed:" & P1.Speed * 10 & "  " & Misc.key
  If Misc.gamestarted = False Then
    Misc.key = "Press Space"
  Else
    Misc.key = ""
  End If
  
  Form1.AutoRedraw = True
  Form1.Cls
  Form1.Print
  Form1.Print
  With P1
    BitBlt Form1.hdc, .X, .Y, .Width, .Height, Form1.Player1.hdc, .Width * Misc.clip, .Height, vbSrcAnd
    BitBlt Form1.hdc, .X, .Y, .Width, .Height, Form1.Player1.hdc, .Width * Misc.clip, 0, vbSrcPaint
    
    If GetAsyncKeyState(vbKeySpace) Then
      BitBlt Form1.hdc, .X, .Y, .Width, .Height, Form1.Player1.hdc, .Width * 5, .Height, vbSrcAnd
      BitBlt Form1.hdc, .X, .Y, .Width, .Height, Form1.Player1.hdc, .Width * 5, 0, vbSrcPaint
    End If
    
    If GetAsyncKeyState(vbKeyLeft) Or GetAsyncKeyState(vbKeyUp) Or GetAsyncKeyState(vbKeyRight) Then
      BitBlt Form1.hdc, .X, .Y, .Width, .Height, Form1.Player1.hdc, .Width * .clip, .Height, vbSrcAnd
      BitBlt Form1.hdc, .X, .Y, .Width, .Height, Form1.Player1.hdc, .Width * .clip, 0, vbSrcPaint
    End If
  End With
  Form1.Refresh
  Form1.AutoRedraw = False
End Sub

Public Function MakeMaskFor(p As PictureBox)

  Dim X As Double
  Dim Y As Double
  Dim Color As Double
  Dim NoColor  As Double
  Dim h As Double
  Dim w As Double
  
  h = p.Height
  w = p.Width
  
  p.Height = p.Height * 2
  NoColor = GetPixel(p.hdc, 0, 0)
    
  For Y = 0 To (p.Height)
    For X = 0 To (p.Width)
      Color = GetPixel(p.hdc, X, Y)
        If Color <> NoColor Then
        
          SetPixelV _
          p.hdc, _
          X, Y + h, _
          vbBlack
        Else
          SetPixelV _
          p.hdc, _
          X, Y, _
          vbBlack
          
          SetPixelV _
          p.hdc, _
          X, Y + h, _
          vbWhite
          
        End If
      Next X
  Next Y
End Function
