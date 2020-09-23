Attribute VB_Name = "Module1"
Public Enum eColor
cRed = 0
cBlue = 1
End Enum

Public Type tCard
Color As eColor
Top As Integer
Left As Integer
Bottom As Integer
Right As Integer
IsThere As Boolean
End Type

Public Function FillCard(Index As Integer, Box) As tCard
On Error GoTo NotThere
'Color & "|" & TopVal & "|" & LeftVal & "|" & BottomVal & "|" & RightVal
Dim TmpTCard As tCard
X = Split(Box(Index).Tag, "|")
TmpTCard.Color = X(0)
TmpTCard.Top = X(1)
TmpTCard.Left = X(2)
TmpTCard.Bottom = X(3)
TmpTCard.Right = X(4)
TmpTCard.IsThere = True
FillCard = TmpTCard
Exit Function
NotThere:
TmpTCard.IsThere = False
End Function
