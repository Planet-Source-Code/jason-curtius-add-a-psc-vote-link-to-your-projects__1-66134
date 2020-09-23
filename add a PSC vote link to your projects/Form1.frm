VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Add a Vote Link"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   1215
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "If you wish to comment or vote for this code"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3075
   End
   Begin Project1.vLink vLink1 
      Height          =   435
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   767
      Caption         =   "Please click here."
      ForeColor       =   16711680
      mess            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   192
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'optional failsafe return from hover color back to link color
vLink1.ForeColor = &HFF0000
'                     ^change this to your link forecolor
End Sub
