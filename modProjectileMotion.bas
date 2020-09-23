Attribute VB_Name = "modProjectileMotion"
Option Explicit
Global Gravity          As Double   'simulation gravity
Global Angle            As Double   'the angle of the trajectory
Global Init_Velocity    As Double   'the initial velocity
Global vx               As Double   'the x component of the velocity vectory

Global TrajectoryColor  As Long     'the color of the path of the trajectory

Global Const PI = 3.14159265        'PI
Private Const TIME_STEP = 0.1

Public Function RadianToDegree(d As Variant) As Double
'Convert a given angle (d), to radians
    RadianToDegree = d * (PI / 180)
End Function

Public Sub drawTrajectory(p As PictureBox, x0 As Integer, y0 As Integer, v0 As Double, g As Double, a As Double)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'p is the picture box to draw the trajectory
'x0 is the initial x value
'y0 is the initial y value
'v0 is the initial velocity (meters per second)
'g is acceleration due to gravity (meters per second per second)
'a is the angle, measured in radians
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim curTime     As Double
Dim tCounter    'THIS CAN BE REMOVED IF YOU ARE NOT USING THE MSFLEXGRID
Dim curX        As Double   'the current x value
Dim curY        As Double   'the current y value
Dim prevX       As Integer  'the previous x value
Dim prevY       As Integer  'the previous y value

prevX = x0
curX = x0
prevY = y0
curY = y0

curTime = 0
    Do
        p.Line (prevX, prevY)-(CInt(curX), CInt(curY)), TrajectoryColor
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''the code between these brackets can be removed
'''''''''to make the function more general it only needed
'''''''''to display data
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Round(tCounter, 1) = 1 Then
            'addDataToGrid curTime, (y0 - curY), (y0 - curY), curX
            addDataToGrid curTime, _
                          getVelocity(vx, (v0 * Sin(a) - (g * curTime))), _
                          Round(getHeight(y0 - 287, v0, a, g, curTime), 2), _
                          Round(getDistance(v0, a, curTime, x0), 2)
            tCounter = 0
        End If
        tCounter = tCounter + TIME_STEP
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        curTime = curTime + TIME_STEP
        

        prevX = CInt(curX)
        prevY = CInt(curY)

        'plot the path parametricly with respect to time
        curX = Init_Velocity * Cos(Angle) * curTime + x0
        curY = -1 * ((Init_Velocity * Sin(Angle)) * curTime - ((0.5) * Gravity * curTime ^ 2) - y0) 'It should actually be '+ y0' at the end but the coordinate system in the picture box is upsidedown
    Loop Until prevY > y0
End Sub

Public Function getHeight(ByVal y0 As Double, ByVal v0 As Double, ByVal a As Double, ByVal g As Double, ByVal t As Double)

getHeight = (v0 * Sin(a)) * t - (0.5 * g * t ^ 2) + y0
End Function

Public Function getMaxHeight(y0 As Double, v0 As Double, a As Double, g As Double)
'y0 = initial y position
'v0 = initial velocity
'a = the angle measured in radians
'g = the gravity
Dim t As Double

'first get the time at which the maximum height is reached (when the vertical velocity = 0)
t = (v0 * Sin(a)) / g

'use the time to get the maximum height
getMaxHeight = (v0 * Sin(a)) * t - (0.5 * g * t ^ 2) + y0
End Function

Public Function getDistance(ByVal v0 As Double, ByVal an As Double, ByVal t As Double, ByVal x0 As Double)
'v0 = initial velocity
'an = the angle measured in radians
't  = the time
'x0 = initial x position

getDistance = (v0 * Cos(an)) * t + x0
End Function

Public Function getMaxDistance(y0 As Double, x0 As Double, v0 As Double, an As Double, g As Double)
'y0 = initial y position
'x0 = initial x position
'v0 = initial velocity
'a  = the angle measured in radians
'g  = the gravity

Dim t       As Double

t = getTime(y0, x0, v0, an, g)

'return the max distance
getMaxDistance = (v0 * Cos(an)) * t + x0
End Function

Public Function getTime(y0 As Double, x0 As Double, v0 As Double, an As Double, g As Double)
'============================================
'Returns the amount of time it takes for one arc
'(from an initial height, back down to the initial height
'============================================
'y0 = initial y position
'x0 = initial x position
'v0 = initial velocity
'a  = angle measured in radians
'g  = gravity

Dim a, b, c     As Double

a = 0.5 * g
b = (v0 * Sin(an)) * -1
c = y0

'in order to obtain the time at which the max distance is reached, we have to solve the
'equation using the quadratic formula. Solving the quadratic formula yeilds 2 answers,
'one being at t=0, and the other is the time when the height equals the initial height

getTime = ((-1 * b) + Sqr(b ^ 2 - (4 * a * c))) / (2 * a)
End Function

Function getVelocity(vx As Double, vy As Double)

getVelocity = Sqr(vx ^ 2 + vy ^ 2)
End Function

''''''''''''''''''''''''''''''''''''''''''''
'THIS IS A PROJECT SPECIFIC SUB ROUTINE
'IT IS ONLY NEEDED IF YOU INTEND TO DISPLAY
'RESULTS TO A MSFLEXGRID
''''''''''''''''''''''''''''''''''''''''''''
Public Sub addDataToGrid(ByVal t As Double, ByVal v As Double, ByVal h As Double, ByVal d As Double)
With frmData.grid
    .TextMatrix(.Rows - 1, 0) = Round(t, 2)
    .TextMatrix(.Rows - 1, 1) = Round(v, 2)
    .TextMatrix(.Rows - 1, 2) = Round(h, 2)
    .TextMatrix(.Rows - 1, 3) = Round(d, 2)
    .Rows = .Rows + 1
End With

End Sub
