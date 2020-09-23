VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projectile Motion"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   0  'User
   ScaleWidth      =   539.683
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   14
      Top             =   0
      Width           =   975
      Begin VB.ComboBox cboWidth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   0
      Width           =   1575
      Begin VB.ComboBox cboColor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmMain.frx":0004
         Left            =   120
         List            =   "frmMain.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Gravity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   1335
      Begin VB.TextBox txtGravity 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "9.8"
         Top             =   240
         Width           =   700
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "m/s²"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   10
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Angle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   1095
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "45"
         Top             =   240
         Width           =   700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Initial velocity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      Begin VB.TextBox txtIniVelocity 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "20"
         Top             =   240
         Width           =   700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "m/s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   4
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.PictureBox canvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   598
      ScaleMode       =   0  'User
      ScaleWidth      =   680.835
      TabIndex        =   0
      Top             =   720
      Width           =   9000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub canvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'used for debugging
'MsgBox "X= " & X
'MsgBox "Y= " & y
'canvas.Refresh
End Sub

Private Sub cboColor_Click()
Select Case cboColor.ListIndex
    Case 0:
        TrajectoryColor = vbBlack
    Case 1:
        TrajectoryColor = vbRed
    Case 2:
        TrajectoryColor = vbGreen
    Case 3:
        TrajectoryColor = vbBlue
End Select
End Sub

Private Sub cboWidth_Click()
    canvas.DrawWidth = Val(cboWidth.ListIndex + 1)
End Sub

Private Sub cmdClear_Click()
canvas.Cls
End Sub

Private Sub cmdLaunch_Click()
Dim x As Integer
Dim y As Integer
x = 0
y = 287

IniData
'canvas.AutoRedraw = False
drawTrajectory canvas, x, y, Init_Velocity, Gravity, Angle
'canvas.AutoRedraw = True
frmData.grid.Rows = frmData.grid.Rows - 1
frmData.lblMaxHeight = CStr(Round(getMaxHeight(0, Init_Velocity, Angle, Gravity), 2)) & " m"
frmData.lblTotalDistance = CStr(Round(getMaxDistance(0, 0, Init_Velocity, Angle, Gravity), 2)) & " m"
End Sub

Private Sub Form_Load()
Dim i As Integer
canvas.Line (0, 0)-(0, 0)
canvas.Refresh

For i = 1 To 10
    cboWidth.AddItem i
Next i

cboWidth.ListIndex = 0
cboColor.ListIndex = 0
End Sub

Private Sub Form_Resize()

canvas.Height = Me.Width
canvas.Width = Me.Height - canvas.Top
End Sub

Public Sub IniData()
'Initialize the variables
    Gravity = Val(frmMain.txtGravity.Text)
    
    'convert the angle to radian before it is used in calculations
    Angle = RadianToDegree(Val(frmMain.txtAngle.Text))
    
    Init_Velocity = Val(frmMain.txtIniVelocity.Text)
    
    'the x-component of the velocity vector stays constant throughout the launch
    vx = Init_Velocity * Cos(Angle)
        
    IniDataForm
End Sub

Public Sub IniDataForm()
Dim i As Integer
With frmData
    .Show
    .Left = Me.Left + Me.Width
    .Top = Me.Top
    .grid.Rows = 2
    .grid.FormatString = "Time|Velocity|Height|Distance"
    For i = 0 To 3
        .grid.ColWidth(i) = 1000
    Next i
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
