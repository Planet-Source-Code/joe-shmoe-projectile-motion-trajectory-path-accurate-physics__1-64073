VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmData 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Data"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483630
      AllowBigSelection=   0   'False
      GridLines       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      FormatString    =   "Time|Velocity|Height|Distance"
   End
   Begin VB.Label lblTotalDistance 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   2325
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Distance:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2325
      Width           =   1080
   End
   Begin VB.Label lblMaxHeight 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Maximum Height:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
