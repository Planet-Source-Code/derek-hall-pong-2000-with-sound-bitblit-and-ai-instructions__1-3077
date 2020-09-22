VERSION 5.00
Begin VB.Form Pong 
   Caption         =   "Pong"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   180
      Top             =   6840
   End
   Begin VB.PictureBox PlayField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   7215
      Left            =   840
      ScaleHeight     =   477
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   0
      Top             =   60
      Width           =   4395
      Begin VB.PictureBox RightShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   480
         Left            =   840
         Picture         =   "Pong.frx":0000
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.PictureBox LeftShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000001&
         Height          =   480
         Left            =   480
         Picture         =   "Pong.frx":0342
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox WallTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "Pong.frx":0984
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Bats 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   540
         Picture         =   "Pong.frx":15C6
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox Balls 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         Picture         =   "Pong.frx":1F0A
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "Pong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Copyright Derek Hall 18/8/1999
' derek.hall@virgin.net
' Your rights
' You may re-distribute this code
' You may not charge for it
' You may alter the code for your own use
' You may learn from it, I hope....
' as for the graphics (By me) you can use them in whatever you want.


' Instructions
' Best viewed in 1024x768 mode to read text.
' Best Played in 800x600 for speed without internet connection.
' try with connection and whatch speed..
' Try to rescale the form.
'Hint 1
'Using a timer is not the best way to get speed... But I have here for simple instructions

'Hint 2
'For more speed
' Use a loop with a doevents init at the maximum loop point
' so you can get the processor hold ups to a minimum.

'if doevents is at the minimum point in the loop then there
'will be more for the processor to do each time it gets to doevents
' so that is why games have jurky movements.
'give the processor all the time you can when multi tasking

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 39 Then
    Bat(Human).Speed = ConstBatSpeed ' while key down make human player speed to constant speed
  ElseIf KeyCode = 37 Then
   Bat(Human).Speed = -(ConstBatSpeed) ' while key down make human player speed to minus the constant speed
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Bat(Human).Speed = 0 ' set this so that the bat does not move while keys are up
End Sub

Private Sub Form_Load()
  SetUPData  'goto sub
  Reset 'goto sub
End Sub

Private Sub printScore()
  'putscore in titlebar
  Me.Caption = "Pong   " & player(Computer).Name & ":" & player(Computer).Score & "     " & player(Human).Name & ":" & player(Human).Score & "  " & "Try to rescale form"
End Sub


Private Sub Form_Resize()
  'if form window state is max or min then reset to normal
  If Pong.WindowState = 1 Or Pong.WindowState = 2 Then Pong.WindowState = 0
  
  'Calculate the blocks so the screen can be resized b 32 pixel blocks
  If Pong.Width < (Tilesize * 15) * 5 Then Pong.Width = (Tilesize * 15) * 7
  If Pong.Height < (Tilesize * 15) * 5 Then Pong.Height = (Tilesize * 15) * 9

  'set map.DivideByX
  Map.DivideByX = Int(Pong.Width / (32 * 15)) 'Divide form size by 32 pixels, Remember it is in Twips.
  Pong.Width = (32 * 15) * Map.DivideByX  ' now make the form width resize to the nearest 32 pixels
  'set map.DivideByY
  Map.DivideByY = Int(Pong.Height / (32 * 15)) 'Divide form size by 32 pixels, Remember it is in Twips.
  Pong.Height = (32 * 15) * Map.DivideByY  ' now make the form width resize to the nearest 32 pixels
  
  PlayField.Top = 2 'set the picturebox Called Playfield Top
  PlayField.Left = 2 'set the picturebox Called Playfield Left
  PlayField.Width = (Pong.Width / 15) - 10 'set the picturebox Called Playfield Width
  PlayField.Height = (Pong.Height / 15) - 28 'set the picturebox Called Playfield Height
  
  Reset ' goto sub
End Sub
Sub Reset()
  Bat(Human).PositionY = PlayField.Height - 28
  Bat(Computer).PositionX = Int(PlayField.ScaleWidth / 2) - (Bats.Width / 2)
  Bat(Human).PositionX = Bat(Computer).PositionX
  
  'Set the balls X at random start in the middle of the playfield
  Ball.PositionX = ((200 * Rnd) - 100) + Int(PlayField.ScaleWidth / 2) - (Balls.Width / 2)
  'Set the balls Y in the middle of the playfield
  Ball.PositionY = Int(PlayField.ScaleHeight / 2) - (Balls.Height / 2)
  
  printScore ' goto sub
End Sub
Private Sub Timer1_Timer()
  Drawmap 'Goto Sub
End Sub

Private Sub Drawmap()
  Dim i As Integer
  PlayField = LoadPicture ' Clear the PlayField
  For i = 0 To Map.DivideByY
    'Draw left side wall
    BitBlt PlayField.hDC, 0, i * Tilesize, Tilesize, Tilesize, WallTile.hDC, 0, 0, SRCCOPY
    'Draw left side wall Shadow
    BitBlt PlayField.hDC, Tilesize, i * Tilesize, 16, Tilesize, LeftShadow.hDC, 0, 0, SRCCOPY
    'Draw right side wall Shadow
    BitBlt PlayField.hDC, PlayField.Width - (Tilesize + 12), i * Tilesize, 8, 32, RightShadow.hDC, 0, 0, SRCCOPY
     'Draw left side wall
    BitBlt PlayField.hDC, PlayField.Width - (Tilesize + 5), i * Tilesize, Tilesize, Tilesize, WallTile.hDC, 0, 0, SRCCOPY
  Next i

  
  GetBatPositions 'goto sub and calculate bat positions
  
  'OK now we draw the bats shadows using a mask but not using SRCPAINT to fill it in
  BitBlt PlayField.hDC, Bat(Computer).PositionX + 4, Bat(Computer).PositionY + 4, 48, 16, Bats.hDC, 0, 0, SRCAND
  BitBlt PlayField.hDC, Bat(Human).PositionX + 4, Bat(Human).PositionY + 4, 48, 16, Bats.hDC, 0, 0, SRCAND
    
  'OK now we draw the bats no masks needed as they are Rectangles
  BitBlt PlayField.hDC, Bat(Computer).PositionX, Bat(Computer).PositionY, 48, 16, Bats.hDC, 0, 0, SRCCOPY
  BitBlt PlayField.hDC, Bat(Human).PositionX, Bat(Human).PositionY, 48, 16, Bats.hDC, 0, 0, SRCCOPY
  
  GetBallPositions 'goto sub and calculate ball positions
  
  'Then the draw ball shadow using a mask but not using SRCPAINT to fill it in
  BitBlt PlayField.hDC, Ball.PositionX + Ball.Height, Ball.PositionY + Ball.Height, 16, 16, Balls.hDC, 0, 0, SRCAND 'mask
  'Then Mask out the ball
  BitBlt PlayField.hDC, Ball.PositionX, Ball.PositionY, 16, 16, Balls.hDC, 0, 0, SRCAND 'mask
  'then blit it using SRCPAINT
  BitBlt PlayField.hDC, Ball.PositionX, Ball.PositionY, 16, 16, Balls.hDC, 16, 0, SRCPAINT 'onto mask
   
  ' only check if you hit the bat if you are in range of it,
  ' this will save you processor time if you do not check all the other code
  If Ball.PositionY < (Bat(Computer).PositionY + Bats.ScaleHeight) Or Ball.PositionY > (Bat(Human).PositionY - (Bats.ScaleHeight)) Then CheckHitABat
   
  PlayField.Refresh  'refresh draws to the Playfield.image so you can see it, make sure auto redraw is true
End Sub
Sub GetBallPositions()
  
  'first add to new position
  Ball.PositionX = Ball.PositionX + Ball.SpeedX
  Ball.PositionY = Ball.PositionY + Ball.SpeedY
  
  'what direction is the ball going, true= south, false=north
  ' this is for the ball shadow so that when it goes the other direction it still falls or goes up in the air
  If Ball.Height > 16 Or Ball.Height < 1 Then Ball.GoingUpOrDown = Not Ball.GoingUpOrDown
  
  If Ball.GoingUpOrDown Then
    Ball.Height = Ball.Height + 1  'shadow Offset
  Else
    Ball.Height = Ball.Height - 1 'shadow Offset
  End If
  
  'Is ball out of play, beyond the walls?
  'Left wall
  If Ball.PositionX > PlayField.ScaleWidth - 56 Then
    HitWallLeft           ' goto sub
    s_Playsound "HitWall" ' play sound
  End If
  'Right wall
  If Ball.PositionX < 32 Then
    HitWallRight          ' goto sub
    s_Playsound "HitWall" ' play sound
  End If
End Sub

Sub CheckHitABat()
  ' first find what end of the field the ball is in,
  'so we only calculate the half the calculations, and do it for the correct end.
  If Ball.PositionY > (Bat(Human).PositionY - 16) Then 'players end
    
    'Check to see if Human the ball or not
    If ((Ball.PositionX > (Bat(Human).PositionX - 12)) And Ball.PositionX < (Bat(Human).PositionX + 48)) Then
      HitBatSpeedChangeUp   ' goto sub
      s_Playsound "HitBat"  ' play sound
    Else
      'Did it go past the bat
      If Ball.PositionY > (Bat(Human).PositionY) Then  'computer wins a point
        
        player(Computer).Score = player(Computer).Score + 1
         Reset      ' goto sub
         printScore ' goto sub
      End If
    End If
  Else ' computer's end
    'Check to see if Computer hit the ball or not
    If ((Ball.PositionX > (Bat(Computer).PositionX - 12)) And Ball.PositionX < (Bat(Computer).PositionX + 48)) Then
      HitBatSpeedChangeDown ' goto sub
      s_Playsound "HitBat"  ' play sound
    Else
      'Did it go past the bat
      If Ball.PositionY < Bat(Computer).PositionY Then
        'computer wins a point
        player(Human).Score = player(Human).Score + 1
        Reset ' goto sub
        printScore ' goto sub
      End If
    End If
  End If
End Sub

Sub GetBatPositions()
  
  If (Bat(Computer).PositionX + 24) > Ball.PositionX Then
    'if Computer bat is greater than balls then go left
    Bat(Computer).PositionX = (Bat(Computer).PositionX) - Bat(Computer).Speed
  Else
    'if Computer bat is less than balls then go Right
    If (Bat(Computer).PositionX + 24) < Ball.PositionX Then Bat(Computer).PositionX = (Bat(Computer).PositionX) + Bat(Computer).Speed
  End If
  'Make sure Computer bat is not further than the wall
  If Bat(Computer).PositionX > (PlayField.Width - 88) Then Bat(Computer).PositionX = (PlayField.Width - 88)
  If Bat(Computer).PositionX < 32 Then Bat(Computer).PositionX = 32
  
  'move Human Player bat
  Bat(Human).PositionX = (Bat(Human).PositionX + Bat(Human).Speed)
  
  'Make sure Players bats are not further than the wall
  If Bat(Human).PositionX > (PlayField.Width - 88) Then Bat(Human).PositionX = (PlayField.Width - 88)
  If Bat(Human).PositionX < 32 Then Bat(Human).PositionX = 32
End Sub


