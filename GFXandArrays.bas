Attribute VB_Name = "GFXandArrays"
' Graphics functions and constants used in the example.
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6    ' Masks
Public Const SRCPAINT = &HEE0086  ' onto masks
Public Const SRCCOPY = &HCC0020   ' backgrounds

'*****************************************************
Public Const Computer = 1
Public Const Human = 2
Public Const ConstBatSpeed = 7
Public Const Tilesize = 32

Private Type PlayfieldData
   DivideByX As Integer   ' the amount of times you can divide the Playfield.Width by 32
   DivideByY As Integer   ' the amount of times you can divide the Playfield.Height by 32
End Type

Private Type BallData
  SpeedX As Integer           ' to hold Speed
  SpeedY As Integer           ' to hold + & - Values for direction & Position
  PositionX As Integer        ' to hold + & - Values for Position
  PositionY As Integer        ' to hold + & - Values for Position
  Height As Integer           ' height offset for ball shadow
  GoingUpOrDown As Boolean    ' to hold + & - Values for direction  True=down
  GoingLeftOrRight As Boolean ' to hold + & - Values for direction  True=Left
  MinSpeedX As Integer        'minimum X speed the ball can travel
  MinSpeedY As Integer        'minimum Y speed the ball can travel
End Type

Private Type BatData
  Speed As Integer           ' can not be more than 10 pixels per move
  PositionX As Integer    ' to hold + & - Values for Position
  PositionY As Integer    ' to hold + & - Values for Position 5 below top and 5 above bottom of playfield
End Type

Private Type PlayerData
  Name As String          'hold player name
  Score As Long           'hold score
End Type

Public Ball As BallData             ' use like  Ball.SpeedX
Public Bat(1 To 2) As BatData       ' use like  Bat(2).PositionY
Public player(1 To 2) As PlayerData ' use like  Player(1).Name
Public Map As PlayfieldData         ' use like  Map.DivideByY

Public Sub SetUPData()

  player(Computer).Name = "Computer"
  player(Human).Name = "Human"
  player(Computer).Score = 0
  player(Human).Score = 0
  Bat(Computer).PositionY = 2
  Bat(Human).PositionY = 10
  Bat(Computer).Speed = ConstBatSpeed 'Computer
  Bat(Human).Speed = 0 'Human do not set until key down, see form keydown event
  Ball.Height = 5 ' height offset for ball shadow
  Ball.GoingLeftOrRight = True
  HitBatSpeedChangeDown
  Ball.MinSpeedX = 8
  Ball.MinSpeedY = 5
End Sub

Private Sub GetARandomBallXDirection()
  Do 'get a random value to take you left or right
    Randomize
    Ball.SpeedX = (30 * Rnd) - 15
  Loop Until Ball.SpeedX > Ball.MinSpeedX Or Ball.SpeedX < -(Ball.MinSpeedX)
End Sub
Public Sub HitBatSpeedChangeDown()
  GetARandomBallXDirection 'goto sub
  
  Randomize 'Get a positiveY direction Value
  Ball.SpeedY = ((6 * Rnd) + 1) + Ball.MinSpeedY
End Sub
Public Sub HitBatSpeedChangeUp()
  GetARandomBallXDirection 'goto sub
  
  Randomize 'Get a negative direction Value
  Ball.SpeedY = -(((6 * Rnd) + 1) + Ball.MinSpeedY)
End Sub
Public Sub HitWallLeft()
  'change from a negative value direction to a positive direction using this calculation
  Ball.SpeedX = Ball.SpeedX + (-(Ball.SpeedX * 2))
End Sub
Public Sub HitWallRight()
  'change from a positive value direction to a negative direction using this calculation
  Ball.SpeedX = (Ball.SpeedX - ((Ball.SpeedX * 2) - 2))
End Sub
