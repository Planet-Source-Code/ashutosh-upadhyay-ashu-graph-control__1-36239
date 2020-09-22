VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGraphData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const MAX_POINTS_PER_GRAPH = 70000
Dim X(MAX_POINTS_PER_GRAPH) As Single
Dim Y(MAX_POINTS_PER_GRAPH) As Single
Public index As Long
Public Total_Data As Long
Public color As Long



Public Function CGraphData()
Total_Data = 0
index = -1
color = RGB(255, 0, 0)

End Function

Public Sub AddData(ByVal xx As Single, ByVal yy As Single)
If Total_Data > MAX_POINTS_PER_GRAPH Then Exit Sub
X(Total_Data) = xx
Y(Total_Data) = yy
Total_Data = Total_Data + 1

End Sub

Public Sub DeleteAll()
Total_Data = 0
index = -1

End Sub
Public Function GetPoint(ByRef xx As Single, ByRef yy As Single, ByRef iindex As Long) As Boolean
If (iindex < 0) Or (iindex > Total_Data) Then
iindex = -1
GetPoint = False
Exit Function
End If
xx = X(iindex)
yy = Y(iindex)
index = iindex
GetPoint = True
End Function
Public Function SetPoint(ByVal xx As Single, ByVal yy As Single, ByVal iindex) As Boolean
If (iindex < 0) Or (iindex > Total_Data) Then
SetPoint = False
Exit Function
End If
X(iindex) = xx
Y(iindex) = yy
index = iindex

End Function
