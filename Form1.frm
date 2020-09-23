VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   480
   ClientTop       =   855
   ClientWidth     =   10065
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   671
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar GravScrl 
      CausesValidation=   0   'False
      Height          =   180
      Left            =   0
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   0
      Value           =   10
      Width           =   1800
   End
   Begin VB.HScrollBar FuelBar 
      CausesValidation=   0   'False
      Height          =   180
      Left            =   0
      Max             =   1024
      TabIndex        =   1
      Top             =   8085
      Value           =   1024
      Width           =   10065
   End
   Begin VB.PictureBox Player1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Game_Timer
  GameSpeed As Double
  LastTick As Double
  
  GameSpeed2 As Double
  LastTick2 As Double
  
  GameSpeed3 As Double
  LastTick3 As Double
End Type
Private t As Game_Timer
Dim X As Integer
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Form_Activate()
  P1.Speed = 0.1
  load
  t.GameSpeed = 15
  t.GameSpeed2 = 30
  t.GameSpeed3 = 120
  Do
    DoEvents
      If timeGetTime - t.LastTick >= t.GameSpeed Then
          If GetAsyncKeyState(vbKeyEscape) Then End
          If GetAsyncKeyState(vbKeyBack) Then
            Form1.Hide
            Form2.Show
          End If
          t.LastTick = timeGetTime()
          Main
          Form1.Width = 10185
          Form1.Height = 8775
          Form1.Top = Screen.Height / 2 - Form1.Height / 2
          Form1.Left = Screen.Width / 2 - Form1.Width / 2
      End If
      
      If timeGetTime - t.LastTick2 >= t.GameSpeed2 Then
        t.LastTick2 = timeGetTime()
        If GetAsyncKeyState(vbKeyLeft) Or GetAsyncKeyState(vbKeyUp) Or GetAsyncKeyState(vbKeyRight) Then
          P1.clip = P1.clip + 1
          If P1.clip > 4 Then
            P1.clip = 3
          End If
          Else
          P1.clip = 3
        End If
      End If
      
      If timeGetTime - t.LastTick3 >= t.GameSpeed3 Then
        t.LastTick3 = timeGetTime()
        If P1.GravityX < 0 Then
          Misc.clip = Misc.clip + 1
          If Misc.clip > 2 Then
            Misc.clip = 0
          End If
        ElseIf P1.GravityX > 0 Then
          Misc.clip = Misc.clip - 1
          If Misc.clip < 0 Then
            Misc.clip = 2
          End If
        ElseIf P1.GravityX = 0 Then
          Misc.clip = Misc.clip
        End If
      End If
  Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub GravScrl_Scroll()
  P1.Speed = Form1.GravScrl.Value * 0.01
  Game_Text_and_Blt

End Sub



