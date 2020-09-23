VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Triple Triad"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Blue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   4
      Left            =   5880
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   18
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox Blue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   3
      Left            =   4440
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   17
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox Blue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   2
      Left            =   3000
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   16
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox Blue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   1560
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   15
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   8
      Left            =   4440
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   14
      Top             =   5280
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   7
      Left            =   3000
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   6
      Left            =   1560
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   5
      Left            =   4440
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   4
      Left            =   3000
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   3
      Left            =   1560
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   2
      Left            =   4440
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   3000
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Red 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   4
      Left            =   5880
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Red 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   3
      Left            =   4440
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Red 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   2
      Left            =   3000
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Red 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   1560
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Blue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox Red 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Card 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   1560
      ScaleHeight     =   1575
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label BlueMarker 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   5880
      TabIndex        =   24
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label RedMarker 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label BlueScore 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   5880
      TabIndex        =   22
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label RedScore 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Turn As eColor
Private RedSelIndex As Integer
Private BlueSelIndex As Integer

Private Sub Blue_Click(Index As Integer)
If Turn = cRed Then Exit Sub
If Blue(Index).Tag = "" Then Exit Sub
For X = 0 To Blue.UBound
Blue(X).Top = 472
Next
Blue(Index).Top = 464
BlueSelIndex = Index
End Sub

Private Sub Card_Click(Index As Integer)
If Card(Index).Tag <> "" Then Exit Sub
If Turn = cRed Then
    If RedSelIndex = -1 Then Exit Sub
Else
    If BlueSelIndex = -1 Then Exit Sub
End If
If Turn = cRed Then
    CheckRules Index, FillCard(RedSelIndex, Red)
Else
    CheckRules Index, FillCard(BlueSelIndex, Blue)
End If
End Sub

Private Sub CheckRules(Spot As Integer, CardInfo As tCard)
DrawCard Card(Spot), CardInfo.Color, CardInfo.Top, CardInfo.Left, CardInfo.Bottom, CardInfo.Right
If CardInfo.Color = cRed Then
    ClearCard Red(RedSelIndex)
    Red(RedSelIndex).Top = 8
    RedSelIndex = -1
Else
    ClearCard Blue(BlueSelIndex)
    Blue(BlueSelIndex).Top = 472
    BlueSelIndex = -1
End If
Dim TopCard As tCard, LeftCard As tCard, BottomCard As tCard, RightCard As tCard
TopCard = FillCard(Spot - 3, Card)
LeftCard = FillCard(Spot - 1, Card)
BottomCard = FillCard(Spot + 3, Card)
RightCard = FillCard(Spot + 1, Card)
    If TopCard.IsThere Then 'top
        If CardInfo.Top > TopCard.Bottom Then ChangeColor Spot - 3, Turn
    End If
    If LeftCard.IsThere Then 'left
        If CardInfo.Left > LeftCard.Right Then ChangeColor Spot - 1, Turn
    End If
    If BottomCard.IsThere Then 'bottom
        If CardInfo.Bottom > BottomCard.Top Then ChangeColor Spot + 3, Turn
    End If
    If RightCard.IsThere Then 'right
        If CardInfo.Right > RightCard.Left Then ChangeColor Spot + 1, Turn
    End If
    If TopCard.IsThere And LeftCard.IsThere Then 'top-left
        If CardInfo.Top = TopCard.Bottom And CardInfo.Left = LeftCard.Right Then ChangeColor Spot - 3, Turn: ChangeColor Spot - 1, Turn
        If (CardInfo.Top + TopCard.Bottom) = (CardInfo.Left + LeftCard.Right) Then ChangeColor Spot - 3, Turn: ChangeColor Spot - 1, Turn
    End If
    If LeftCard.IsThere And BottomCard.IsThere Then 'bottom-left
        If CardInfo.Left = LeftCard.Right And CardInfo.Bottom = BottomCard.Top Then ChangeColor Spot - 1, Turn: ChangeColor Spot + 3, Turn
        If (CardInfo.Left + LeftCard.Right) = (CardInfo.Bottom + BottomCard.Top) Then ChangeColor Spot - 1, Turn: ChangeColor Spot + 3, Turn
    End If
    If BottomCard.IsThere And RightCard.IsThere Then 'botton-right
        If CardInfo.Bottom = BottomCard.Top And CardInfo.Right = RightCard.Left Then ChangeColor Spot + 1, Turn: ChangeColor Spot + 3, Turn
        If (CardInfo.Bottom + BottomCard.Top) = (CardInfo.Right + RightCard.Left) Then ChangeColor Spot + 1, Turn: ChangeColor Spot + 3, Turn
    End If
    If TopCard.IsThere And RightCard.IsThere Then 'top-right
        If CardInfo.Top = TopCard.Bottom And CardInfo.Right = RightCard.Left Then ChangeColor Spot + 1, Turn: ChangeColor Spot - 3, Turn
        If (CardInfo.Top + TopCard.Bottom) = (CardInfo.Right + RightCard.Left) Then ChangeColor Spot + 1, Turn: ChangeColor Spot - 3, Turn
    End If
    If TopCard.IsThere And BottomCard.IsThere Then 'top-bottom
        If CardInfo.Top = TopCard.Bottom And CardInfo.Bottom = BottomCard.Top Then ChangeColor Spot + 3, Turn: ChangeColor Spot - 3, Turn
        If (CardInfo.Top + TopCard.Bottom) = (CardInfo.Bottom + BottomCard.Top) Then ChangeColor Spot + 3, Turn: ChangeColor Spot - 3, Turn
    End If
    If LeftCard.IsThere And RightCard.IsThere Then 'left-right
        If CardInfo.Left = LeftCard.Right And CardInfo.Right = RightCard.Left Then ChangeColor Spot + 1, Turn: ChangeColor Spot - 1, Turn
        If (CardInfo.Left + LeftCard.Right) = (CardInfo.Right + RightCard.Left) Then ChangeColor Spot + 1, Turn: ChangeColor Spot - 1, Turn
    End If
CheckScores
ChangeTurn IIf(CardInfo.Color = cRed, cBlue, cRed)
End Sub

Private Sub ChangeColor(Spot As Integer, Color As eColor)
Dim TmpTCard As tCard
TmpTCard = FillCard(Spot, Card)
DrawCard Card(Spot), Color, TmpTCard.Top, TmpTCard.Left, TmpTCard.Bottom, TmpTCard.Right
If Color = cBlue Then
    BlueScore.Caption = Val(BlueScore.Caption + 1)
    RedScore.Caption = Val(RedScore.Caption - 1)
Else
    BlueScore.Caption = Val(BlueScore.Caption - 1)
    RedScore.Caption = Val(RedScore.Caption + 1)
End If
'CheckScores
End Sub

Private Sub Form_Load()
NewGame
End Sub

Private Sub DrawCard(Picbox As PictureBox, Color As eColor, TopVal As Integer, LeftVal As Integer, BottomVal As Integer, RightVal As Integer)
ClearCard Picbox
Dim Top As String * 1, Left As String * 1, Bottom As String * 1, Right As String * 1
Top = IIf(TopVal = 10, "A", TopVal)
Left = IIf(LeftVal = 10, "A", LeftVal)
Bottom = IIf(BottomVal = 10, "A", BottomVal)
Right = IIf(RightVal = 10, "A", RightVal)
Picbox.BackColor = IIf(Color = cRed, vbRed, vbBlue)
Picbox.Tag = Color & "|" & TopVal & "|" & LeftVal & "|" & BottomVal & "|" & RightVal
'top
Picbox.CurrentX = (Picbox.ScaleWidth - Picbox.TextWidth(Top)) / 2
Picbox.CurrentY = 0
Picbox.Print Top
'left
Picbox.CurrentX = 0
Picbox.CurrentY = (Picbox.ScaleHeight - Picbox.TextHeight(Left)) / 2
Picbox.Print Left
'bottom
Picbox.CurrentX = (Picbox.ScaleWidth - Picbox.TextWidth(Bottom)) / 2
Picbox.CurrentY = Picbox.ScaleHeight - Picbox.TextHeight(Bottom)
Picbox.Print Bottom
'right
Picbox.CurrentX = Picbox.ScaleWidth - Picbox.TextWidth(Right)
Picbox.CurrentY = (Picbox.ScaleHeight - Picbox.TextHeight(Right)) / 2
Picbox.Print Right
End Sub

Private Sub ClearCard(Picbox As PictureBox)
Picbox.Tag = ""
Picbox.BackColor = vbWhite
Picbox.Cls
End Sub

Private Sub Red_Click(Index As Integer)
If Turn = cBlue Then Exit Sub
If Red(Index).Tag = "" Then Exit Sub
For X = 0 To Red.UBound
Red(X).Top = 8
Next
Red(Index).Top = 16
RedSelIndex = Index
End Sub

Private Sub CheckScores()
Dim X As Integer, RedBCount As Integer, BlueBCount As Integer, RedCount As Integer, BlueCount As Integer, Blah As tCard
'red cards
For X = 0 To 4
Blah = FillCard(X, Red)
If Blah.IsThere Then RedCount = RedCount + 1
Next X
'blue cards
For X = 0 To 4
Blah = FillCard(X, Blue)
If Blah.IsThere Then BlueCount = BlueCount + 1
Next X
'board cards
For X = 0 To 8
Blah = FillCard(X, Card)
If Blah.Color = cBlue And Blah.IsThere Then
    BlueBCount = BlueBCount + 1
ElseIf Blah.Color = cRed And Blah.IsThere Then
    RedBCount = RedBCount + 1
End If
Next X
RedScore.Caption = CStr(RedCount + RedBCount)
BlueScore.Caption = CStr(BlueCount + BlueBCount)
If RedBCount + BlueBCount = 9 Then
    If BlueScore.Caption > RedScore.Caption Then
        WinningColor = "Blue"
    ElseIf RedScore.Caption > BlueScore.Caption Then
        WinningColor = "Red"
    ElseIf RedScore.Caption = BlueScore.Caption Then
        WinningColor = "No one"
    Else
        MsgBox "wtf? lol red:" & RedScore.Caption & " blue:" & BlueScore.Caption & " boardcount:" & CStr(RedBCount + BlueBCount)
    End If
    MsgBox "Good Game! " & WinningColor & " wins!": NewGame
End If
End Sub

Private Sub NewGame()
'give red random cards
For X = 0 To Red.UBound
Randomize Timer
DrawCard Red(X), cRed, Int(Rnd * 10) + 1, Int(Rnd * 10) + 1, Int(Rnd * 10) + 1, Int(Rnd * 10) + 1
Next X
'give blue random cards
For X = 0 To Blue.UBound
Randomize Timer
DrawCard Blue(X), cBlue, Int(Rnd * 10) + 1, Int(Rnd * 10) + 1, Int(Rnd * 10) + 1, Int(Rnd * 10) + 1
Next X
'reset board
For X = 0 To 8
ClearCard Card(X)
Next X
BlueScore.Caption = "5"
RedScore.Caption = "5"
Randomize Timer
ChangeTurn Int(Rnd * 2)
RedSelIndex = -1
BlueSelIndex = -1
End Sub

Private Sub ChangeTurn(Color As eColor)
Turn = Color
If Turn = cBlue Then BlueMarker.Caption = "*": RedMarker.Caption = "" Else RedMarker.Caption = "*": BlueMarker.Caption = ""
End Sub
