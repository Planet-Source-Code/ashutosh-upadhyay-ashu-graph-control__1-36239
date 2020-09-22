VERSION 5.00
Object = "*\AProject1.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6804
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7704
   LinkTopic       =   "Form1"
   ScaleHeight     =   6804
   ScaleWidth      =   7704
   StartUpPosition =   3  'Windows Default
   Begin GraphControl.AshuGraphControl AshuGraphControl1 
      Height          =   4572
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   6852
      _ExtentX        =   12086
      _ExtentY        =   8065
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleHeight     =   381
      ScaleWidth      =   571
   End
   Begin VB.CommandButton Command3 
      Caption         =   "print as bitmap"
      Height          =   972
      Left            =   5160
      TabIndex        =   2
      Top             =   5160
      Width           =   2292
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw My Graph"
      Height          =   852
      Left            =   2640
      TabIndex        =   1
      Top             =   5160
      Width           =   2292
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Me"
      Height          =   852
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   2052
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call AshuGraphControl1.PrintSettings(10, 10, "ashu", "X --->", "Y ^ | | |", 1, 200, 9, -1)
''' zoom factor will only work if your pinter support zooming printing
''''paper size = 9 means  A4 , if = 10 means A4 small, see msdn help on "printer" in vb editor
''' printquality = - 1 means draft(poor)

AshuGraphControl1.PrintMe

End Sub

Private Sub Command2_Click()
Dim p As Single
Dim q As Single
Dim i As Single
AshuGraphControl1.SetColor RGB(0, 200, 0), 1
For i = 1 To 100 Step 0.05
q = (5 * i)
'''' don't give negative data
p = (100 * CSng(Cos(i)) + 100)
Call AshuGraphControl1.AddData(q, p, 1)
Next i
AshuGraphControl1.Invalidate
End Sub



Private Sub Command3_Click()
AshuGraphControl1.PrintOnlyShownPortionAsBitmap
End Sub

Private Sub Form_Load()
AshuGraphControl1.InitializeMe
AshuGraphControl1.SetColor RGB(255, 0, 0), 0
AshuGraphControl1.SineCurveFill (0)

End Sub
