VERSION 5.00
Begin VB.UserControl AshuGraphControl 
   ClientHeight    =   3516
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4596
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   Begin VB.HScrollBar HScroll1 
      Height          =   132
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2652
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2172
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   132
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00404040&
      ForeColor       =   &H00FF0000&
      Height          =   2772
      Left            =   0
      Picture         =   "UserControl1.ctx":0000
      ScaleHeight     =   227
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   0
      Top             =   0
      Width           =   3972
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H80000001&
         Height          =   216
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   516
      End
   End
End
Attribute VB_Name = "AshuGraphControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_XUnitToPixels = 1
Const m_def_YUnitToPixels = 1
Const m_def_XGap = 50
Const m_def_YGap = 50
Const m_def_XMinPosition = 0
Const m_def_YMinPosition = 0
'Property Variables:
Dim m_FontColor As OLE_COLOR
Dim m_XUnitToPixels As Single
Dim m_YUnitToPixels As Single
Dim m_XGap As Single
Dim m_YGap As Single
Dim m_XMinPosition As Single
Dim m_YMinPosition As Single
Dim m_ShowGrids As Boolean
Dim m_GridColor As OLE_COLOR
'''''
''''''Variable of class CGraphData to store graph points
Const MAX_NO_OF_GRAPHS = 5
Dim m_Graph(MAX_NO_OF_GRAPHS) As New CGraphData
'''''following values are used for scrolling graph as well as deciding begining point of graph
Dim m_XInitialMargin As Single
Dim m_YInitialMargin As Single

'''''following variables holds maximum and minimum values of x & y among all graphs
'''''these are used in scrolling graph  as well as printing
Dim p_MaxX As Single
Dim p_MinX As Single
Dim p_MaxY As Single
Dim p_MinY As Single
'''''following variables are used in printing graphs
Dim p_PageHeight As Long '''size in pixels
Dim p_PageWidth As Long
Dim p_NumCols As Long    '''' no of pages in row and column
Dim p_NumRows As Long    ''' total no of pages will be printed = p_NumCols * p_NumRows
Dim p_XScaleHeight As Long
Dim p_YScaleWidth As Long
Dim p_XPageMargin As Long    '''left & top margins on paper
Dim p_YPagemargin As Long
Dim p_XGraphSize As Long  ''''max. size in pixels on paper
Dim p_YGraphSize As Long
Dim p_XInitialMargin As Single  ''' to decide begining point of graph and scale on current paper
Dim p_YInitialMargin As Single
Dim p_zoomfactor As Long
Dim p_Header As String
Dim p_XLabel As String
Dim p_YLabel As String
Dim p_PrintPageNo As Integer ' 0 means don't print  1 means print at top   2 means print as bottom
''''''''''variable to tell whether track mouse pointer or not
Dim m_TrackPointer As Boolean
'''''End of list of variables


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
If (New_Appearance = 0) Or (New_Appearance = 1) Then
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    If (New_BorderStyle = 0) Or (New_BorderStyle = 1) Then
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    Picture1.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Picture1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Dim tmp As Long
    Set Picture1.Font = New_Font
     PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Picture1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Picture1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = Picture1.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
'Public Property Get FontColor() As OLE_COLOR
'     FontColor = m_FontColor
'End Property

'Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
'     m_FontColor = New_FontColor
'    PropertyChanged "FontColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,BackColor
Public Property Get TooltipBkColor() As OLE_COLOR
Attribute TooltipBkColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    TooltipBkColor = Label1.BackColor
End Property

Public Property Let TooltipBkColor(ByVal New_TooltipBkColor As OLE_COLOR)
    Label1.BackColor() = New_TooltipBkColor
    PropertyChanged "TooltipBkColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get TooltipForeColor() As OLE_COLOR
Attribute TooltipForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TooltipForeColor = Label1.ForeColor
End Property

Public Property Let TooltipForeColor(ByVal New_TooltipForeColor As OLE_COLOR)
    Label1.ForeColor() = New_TooltipForeColor
    PropertyChanged "TooltipForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get XUnitToPixels() As Single
    XUnitToPixels = m_XUnitToPixels
End Property

Public Property Let XUnitToPixels(ByVal New_XUnitToPixels As Single)
    m_XUnitToPixels = New_XUnitToPixels
    
    PropertyChanged "XUnitToPixels"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,1
Public Property Get YUnitToPixels() As Single
    YUnitToPixels = m_YUnitToPixels
End Property

Public Property Let YUnitToPixels(ByVal New_YUnitToPixels As Single)
    m_YUnitToPixels = New_YUnitToPixels
    
    PropertyChanged "YUnitToPixels"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get XGap() As Single
    XGap = m_XGap
End Property

Public Property Let XGap(ByVal New_XGap As Single)
    m_XGap = New_XGap
    HScroll1.LargeChange = CInt(Me.XGap * Me.XUnitToPixels)
    PropertyChanged "XGap"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get YGap() As Single
    YGap = m_YGap
End Property

Public Property Let YGap(ByVal New_YGap As Single)
    m_YGap = New_YGap
    VScroll1.LargeChange = CInt(Me.YGap * Me.YUnitToPixels)
    PropertyChanged "YGap"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get XMinPosition() As Single
    XMinPosition = m_XMinPosition
End Property

Public Property Let XMinPosition(ByVal New_XMinPosition As Single)
    m_XMinPosition = New_XMinPosition
    p_MinX = Me.XMinPosition
    If p_MaxX < p_MinX Then p_MaxX = p_MinX
    PropertyChanged "XMinPosition"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get YMinPosition() As Single
    YMinPosition = m_YMinPosition
End Property

Public Property Let YMinPosition(ByVal New_YMinPosition As Single)
    m_YMinPosition = New_YMinPosition
    
    p_MinY = Me.YMinPosition
    If p_MaxY < p_MinY Then p_MaxY = p_MinY
    PropertyChanged "YMinPosition"
End Property
Public Property Get XScaleHeight() As Long
           XScaleHeight = p_XScaleHeight
End Property
Public Property Let XScaleHeight(ByVal New_XScaleHeight As Long)
        p_XScaleHeight = New_XScaleHeight
        PropertyChanged "XScaleHeight"
End Property
Public Property Get YScaleWidth() As Long
         YScaleWidth = p_YScaleWidth
End Property
Public Property Let YScaleWidth(ByVal new_yscalewidth As Long)
        p_YScaleWidth = new_yscalewidth
        PropertyChanged "YScaleWidth"
End Property
Public Property Get TrackMousePointer() As Boolean
     TrackMousePointer = m_TrackPointer
End Property
Public Property Let TrackMousePointer(ByVal new_val As Boolean)
        m_TrackPointer = new_val
        PropertyChanged "TrackMousePointer"
End Property
Public Property Get ShowGrids() As Boolean
        ShowGrids = m_ShowGrids
End Property
Public Property Let ShowGrids(ByVal new_val As Boolean)
           m_ShowGrids = new_val
           PropertyChanged "ShowGrids"
End Property
Public Property Get GridColor() As OLE_COLOR
         GridColor = m_GridColor
End Property
Public Property Let GridColor(ByVal new_val As OLE_COLOR)
             m_GridColor = new_val
             PropertyChanged "GridColor"
End Property
Private Sub HScroll1_Change()
m_XInitialMargin = HScroll1.Value
Label1.Visible = False
UserControl_Paint
End Sub

Private Sub HScroll1_Scroll()
m_XInitialMargin = HScroll1.Value
Label1.Visible = False
UserControl_Paint
End Sub

Private Sub Picture1_LostFocus()
Label1.Visible = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call drawGraph(Picture1.hDC, p_YScaleWidth + 1, p_XScaleHeight + 1, Picture1.Height - p_XScaleHeight - 6, Picture1.Width - p_YScaleWidth - 6, i)
Dim xx As Single
Dim yy As Single
Dim XFirstPoint As Single
Dim YFirstpoint As Single
Dim str1 As String
Dim str2 As String

If TrackMousePointer Then
     If (X > (p_YScaleWidth - 1)) And (Y > (p_XScaleHeight - 1)) And (Y < (Picture1.Height)) And (X < (Picture1.Width)) Then
          Label1.Visible = False
          XFirstPoint = Me.XMinPosition + m_XInitialMargin
          YFirstpoint = Me.YMinPosition + m_YInitialMargin
          If (Me.XUnitToPixels <> 0) And (Me.YUnitToPixels <> 0) Then
              xx = CSng(CSng(X - p_YScaleWidth + 0) / Me.XUnitToPixels + XFirstPoint)
              yy = CSng(CSng(Picture1.Height - 5 - Y) / Me.YUnitToPixels + YFirstpoint)
          End If
          str1 = Format(xx, "##0.00")
          str2 = Format(yy, "##0.00")
          Label1.Caption = "( " & str1 & ", " & str2 & ")"
          If (Picture1.Width - X - 16 - Label1.Width) > 0 Then
                   Label1.Left = X + 16
          Else
                   Label1.Left = X - Label1.Width - 1
          End If
          If (Picture1.Height - Y - 16 - Label1.Height - 5) > 0 Then
                   Label1.Top = Y + 16
          Else
                   Label1.Top = Y - Label1.Height - 1
          End If
          Label1.Visible = True
                   
     Else
          Label1.Visible = False
     End If
End If
End Sub

Private Sub UserControl_ExitFocus()
Label1.Visible = False
End Sub

Private Sub UserControl_Initialize()
      If m_XUnitToPixels = 0 Then m_XUnitToPixels = 1
      If m_YUnitToPixels = 0 Then m_YUnitToPixels = 1
      If m_XGap = 0 Then m_XGap = m_def_XGap
      If m_YGap = 0 Then m_YGap = m_def_YGap
      If m_XMinPosition = 0 Then m_XMinPosition = m_def_XMinPosition
      If m_YMinPosition = 0 Then m_YMinPosition = m_def_YMinPosition
      Picture1.Left = 0
      Picture1.Top = 0
      HScroll1.Left = 0
      VScroll1.Top = 0
      HScroll1.Top = Me.ScaleHeight - HScroll1.Height
      VScroll1.Left = Me.ScaleWidth - VScroll1.Width
      HScroll1.Width = Me.ScaleWidth - VScroll1.Width
      VScroll1.Height = Me.ScaleHeight - HScroll1.Height
      Picture1.Width = Me.ScaleWidth - VScroll1.Width
      Picture1.Height = Me.ScaleHeight - HScroll1.Height
      
      initializeGraph
      initializeScrolls
      Label1.Visible = False
      
     
      
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
     m_FontColor = RGB(0, 0, 255)
    m_XUnitToPixels = m_def_XUnitToPixels
    m_YUnitToPixels = m_def_YUnitToPixels
    m_XGap = m_def_XGap
    m_YGap = m_def_YGap
    m_XMinPosition = m_def_XMinPosition
    m_YMinPosition = m_def_YMinPosition
    p_XScaleHeight = 20
    p_YScaleWidth = 40
    m_TrackPointer = True
    m_ShowGrids = True
    
    m_GridColor = QBColor(14)
End Sub

Private Sub UserControl_LostFocus()
Label1.Visible = False
End Sub

Private Sub UserControl_Paint()
Dim i As Long
Picture1.Cls
'Picture1.Refresh
Call drawHScale(Picture1.hdc, p_YScaleWidth, 0, p_XScaleHeight, Picture1.Width - p_YScaleWidth - 5)
Call drawVScale(Picture1.hdc, 0, p_XScaleHeight, Picture1.Height - p_XScaleHeight - 5, p_YScaleWidth)
If Me.ShowGrids Then
Call ShowHGrid(Picture1.hdc, p_YScaleWidth, p_XScaleHeight, Picture1.Height - p_XScaleHeight - 6, Picture1.Width - p_YScaleWidth - 6)
Call ShowVGrid(Picture1.hdc, p_YScaleWidth, p_XScaleHeight, Picture1.Height - p_XScaleHeight - 6, Picture1.Width - p_YScaleWidth - 6)
End If
For i = 0 To MAX_NO_OF_GRAPHS
Call drawGraph(Picture1.hdc, p_YScaleWidth, p_XScaleHeight, Picture1.Height - p_XScaleHeight - 6, Picture1.Width - p_YScaleWidth - 6, i)
Next i

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", &HE0E0E0)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Picture1.ForeColor = PropBag.ReadProperty("ForeColor", &HFF0000)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 2880)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 3840)
     m_FontColor = PropBag.ReadProperty("FontColor", RGB(0, 0, 255))
    Label1.BackColor = PropBag.ReadProperty("TooltipBkColor", &H80000005)
    Label1.ForeColor = PropBag.ReadProperty("TooltipForeColor", &H80000001)
    m_XUnitToPixels = PropBag.ReadProperty("XUnitToPixels", m_def_XUnitToPixels)
    m_YUnitToPixels = PropBag.ReadProperty("YUnitToPixels", m_def_YUnitToPixels)
    m_XGap = PropBag.ReadProperty("XGap", m_def_XGap)
    m_YGap = PropBag.ReadProperty("YGap", m_def_YGap)
    m_XMinPosition = PropBag.ReadProperty("XMinPosition", m_def_XMinPosition)
    m_YMinPosition = PropBag.ReadProperty("YMinPosition", m_def_YMinPosition)
    p_XScaleHeight = PropBag.ReadProperty("XScaleHeight", 20)
    p_YScaleWidth = PropBag.ReadProperty("YScaleWidth", 40)
    m_TrackPointer = PropBag.ReadProperty("TrackMousePointer", True)
    m_ShowGrids = PropBag.ReadProperty("ShowGrids", True)
    m_GridColor = PropBag.ReadProperty("GridColor", QBColor(14))
End Sub

Private Sub UserControl_Resize()
      HScroll1.Top = Me.ScaleHeight - HScroll1.Height
      VScroll1.Left = Me.ScaleWidth - VScroll1.Width
      HScroll1.Width = Me.ScaleWidth - VScroll1.Width
      VScroll1.Height = Me.ScaleHeight - HScroll1.Height
      Picture1.Width = Me.ScaleWidth - VScroll1.Width
      Picture1.Height = Me.ScaleHeight - HScroll1.Height
End Sub

Private Sub UserControl_Terminate()
Dim i As Long
For i = 0 To ((i < MAX_NO_OF_GRAPHS) Or (i = MAX_NO_OF_GRAPHS))
Set m_Graph(i) = Nothing
Next i
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Picture1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Picture1.ForeColor, &HFF0000)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 2880)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 3840)
    Call PropBag.WriteProperty("FontColor", m_FontColor, RGB(0, 0, 255))
    Call PropBag.WriteProperty("TooltipBkColor", Label1.BackColor, &H80000005)
    Call PropBag.WriteProperty("TooltipForeColor", Label1.ForeColor, &H80000001)
    Call PropBag.WriteProperty("XUnitToPixels", m_XUnitToPixels, m_def_XUnitToPixels)
    Call PropBag.WriteProperty("YUnitToPixels", m_YUnitToPixels, m_def_YUnitToPixels)
    Call PropBag.WriteProperty("XGap", m_XGap, m_def_XGap)
    Call PropBag.WriteProperty("YGap", m_YGap, m_def_YGap)
    Call PropBag.WriteProperty("XMinPosition", m_XMinPosition, m_def_XMinPosition)
    Call PropBag.WriteProperty("YMinPosition", m_YMinPosition, m_def_YMinPosition)
    Call PropBag.WriteProperty("XScaleHeight", p_XScaleHeight, 20)
    Call PropBag.WriteProperty("YScaleWidth", p_YScaleWidth, 40)
    Call PropBag.WriteProperty("TrackMousePointer", m_TrackPointer, True)
    Call PropBag.WriteProperty("ShowGrids", m_ShowGrids, True)
    Call PropBag.WriteProperty("GridColor", m_GridColor, QBColor(14))
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''funtions for initializing
Public Sub SetColor(color As Long, GraphNo As Long)
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
m_Graph(GraphNo).color = color


End Sub
Private Sub initializeGraph()
Dim i As Long
Dim tmp As Long
m_XInitialMargin = 0#
m_YInitialMargin = 0#
p_MinX = Me.XMinPosition
p_MinY = Me.YMinPosition
p_MaxX = p_MinX
p_MaxY = p_MinY
p_NumCols = 1
p_NumRows = 1
p_PageHeight = 1
p_PageWidth = 1
p_XGraphSize = 1
p_XInitialMargin = 0#
p_XPageMargin = 0
p_XScaleHeight = 20
p_YGraphSize = 1
p_YInitialMargin = 0#
p_YPagemargin = 0
p_YScaleWidth = 40
p_zoomfactor = 100
 p_Header = ""
      p_XLabel = ""
      p_YLabel = ""
      p_PrintPageNo = 0
End Sub
Private Sub initializeScrolls()
HScroll1.Min = 0
HScroll1.SmallChange = 1
HScroll1.LargeChange = CInt(Me.XGap * Me.XUnitToPixels)
HScroll1.Max = CInt((p_MaxX - p_MinX + 1) * Me.XUnitToPixels) - Picture1.Width
HScroll1.Value = CInt((Me.XMinPosition - p_MinX) * Me.XUnitToPixels)
VScroll1.Min = 0
VScroll1.SmallChange = 1
VScroll1.LargeChange = CInt(Me.YGap * Me.YUnitToPixels)
VScroll1.Max = CInt((p_MaxY - p_MinY + 1) * Me.YUnitToPixels)
VScroll1.Value = VScroll1.Max - CInt((Me.YMinPosition - p_MinY) * Me.YUnitToPixels)

End Sub
Public Sub InitializeMe()
Dim i As Long
For i = 0 To MAX_NO_OF_GRAPHS
m_Graph(i).index = -1
m_Graph(i).Total_Data = 0
Next i
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''Function for data manipulation
Public Sub AddData(xx As Single, yy As Single, GraphNo As Long)
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
m_Graph(GraphNo).AddData xx, yy
If p_MaxX < xx Then p_MaxX = xx
If p_MaxY < yy Then p_MaxY = yy
If p_MinX > xx Then p_MinX = xx
If p_MinY > yy Then p_MinY = yy
HScroll1.Max = CInt((p_MaxX - p_MinX + 1) * Me.XUnitToPixels) ''- Picture1.Width
VScroll1.Max = CInt((p_MaxY - p_MinY + 1) * Me.YUnitToPixels) ''- Picture1.Height
VScroll1.Value = VScroll1.Max - CInt((Me.YMinPosition - p_MinY) * Me.YUnitToPixels)
HScroll1.Value = CInt((Me.XMinPosition - p_MinX) * Me.XUnitToPixels)

End Sub
Public Function SetData(xx As Single, yy As Single, index As Long, GraphNo As Long) As Boolean
SetData = False
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Function
If (m_Graph(GraphNo).SetPoint(xx, yy, index)) Then
If p_MaxX < xx Then p_MaxX = xx
If p_MaxY < yy Then p_MaxY = yy
If p_MinX > xx Then p_MinX = xx
If p_MinY > yy Then p_MinY = yy
HScroll1.Max = CInt((p_MaxX - p_MinX + 1) * Me.XUnitToPixels) ''- Picture1.Width
VScroll1.Max = CInt((p_MaxY - p_MinY + 1) * Me.YUnitToPixels)
VScroll1.Value = VScroll1.Max - CInt((Me.YMinPosition - p_MinY) * Me.YUnitToPixels)
HScroll1.Value = CInt((Me.XMinPosition - p_MinX) * Me.XUnitToPixels)
SetData = True
End If
End Function
Public Function GetData(ByRef xx As Single, ByRef yy As Single, ByRef index As Long, GraphNo As Long) As Boolean
GetData = False
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Function
If (m_Graph(GraphNo).GetPoint(xx, yy, index)) Then GetData = True

End Function
Public Sub RemoveGraph(GraphNo As Long)
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
m_Graph(GraphNo).index = -1
m_Graph(GraphNo).Total_Data = 0

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''Private Functions for drawing graph and scales
Private Sub drawGraph(ByRef hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long, GraphNo As Long)
Dim i As Long
Dim xx As Single
Dim yy As Single
Dim XFirstPoint As Single
Dim YFirstpoint As Single
Dim Xdistance As Single
Dim Ydistance As Single
Dim tmp As Long

If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub
If (m_Graph(GraphNo).Total_Data = 0) Or (m_Graph(GraphNo).Total_Data < 0) Then Exit Sub

XFirstPoint = Me.XMinPosition + m_XInitialMargin
YFirstpoint = Me.YMinPosition + m_YInitialMargin
Xdistance = 0#
Ydistance = 0#
For i = 0 To m_Graph(GraphNo).Total_Data
If (m_Graph(GraphNo).GetPoint(xx, yy, i)) Then
    If ((xx > XFirstPoint) Or (xx = XFirstPoint)) And ((yy > YFirstpoint) Or (yy = YFirstpoint)) Then
          Xdistance = (xx - XFirstPoint) * Me.XUnitToPixels
          Ydistance = (yy - YFirstpoint) * Me.YUnitToPixels
          If (Xdistance < wWidth) And (Ydistance < hHeight) Then
                 tmp = SetPixelV(hdc, CLng(Xdistance) + Xdisplacement, hHeight + Ydisplacement - CLng(Ydistance), m_Graph(GraphNo).color)
                 'Picture1.PSet (CLng(Xdistance) + Xdisplacement, hHeight + Ydisplacement - CLng(Ydistance)), RGB(255, 0, 0)
          End If
    End If
End If
Next i


End Sub
Private Sub drawHScale(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI
Dim oldPen As Long
Dim newPen As Long
Dim newFont As Long
Dim oldFont As Long
Dim tmp As Long
Dim tmp1 As Long
Dim oldColor As Long
Dim str As String
Dim FirstPoint As Single
Dim distance As Single
Dim i As Single
''''''create objects of pen,font,etc.
'newPen = CreatePen(0, 0, Me.ForeColor)
'oldPen = SelectObject(hDC, newPen)
''''calculate height of font
tmp = -1 * MulDiv(Me.Font.Size, GetDeviceCaps(hdc, 90), 72)
tmp = Abs(tmp)
'newFont = CreateFont(tmp, 0, 0, 0, Me.Font.Weight, Me.Font.Italic, Me.Font.Underline, Me.Font.Strikethrough, Me.Font.Charset, 0, 0, 0, 0, Me.Font.Name)
'olfont = SelectObject(hDC, newFont)
'oldColor = SetTextColor(hDC, Me.FontColor)
'''''''draw boundary lines
tmp1 = MoveToEx(hdc, Xdisplacement, Ydisplacement + hHeight, ptAPI)
tmp1 = LineTo(hdc, Xdisplacement + wWidth, Ydisplacement + hHeight)
tmp1 = LineTo(hdc, Xdisplacement + wWidth, Ydisplacement + tmp)

FirstPoint = Me.XMinPosition + m_XInitialMargin
distance = 0#
i = FirstPoint

Do
   distance = (i - FirstPoint) * Me.XUnitToPixels
   If distance > CSng(wWidth) Then Exit Do
   str = Format(i, "#0.00")
   tmp1 = TextOut(hdc, Xdisplacement + CLng(distance), Ydisplacement, str, Len(str))
   tmp1 = MoveToEx(hdc, Xdisplacement + CLng(distance), Ydisplacement + tmp, ptAPI)
   tmp1 = LineTo(hdc, Xdisplacement + CLng(distance), Ydisplacement + hHeight)
   i = i + Me.XGap
   
Loop While distance < wWidth

'SelectObject hDC, oldFont
'SelectObject hDC, oldPen
'DeleteObject newFont
'DeleteObject newPen
'oldFont = 0
'oldPen = 0
End Sub



Private Sub drawVScale(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI
Dim oldPen As Long
Dim newPen As Long
Dim newFont As Long
Dim oldFont As Long
Dim tmp As Long
Dim oldColor As Long
Dim str As String
Dim FirstPoint As Single
Dim distance As Single
Dim i As Single
Dim lowerpoint As Boolean
''''''create objects of pen,font,etc.
'newPen = CreatePen(0, 0, Me.ForeColor)
'oldPen = SelectObject(hDC, newPen)
''''calculate height of font
tmp = -1 * MulDiv(Me.Font.Size, GetDeviceCaps(hdc, 90), 72)
tmp = Abs(tmp)
'newFont = CreateFont(tmp, 0, 0, 0, Me.Font.Weight, Me.Font.Italic, Me.Font.Underline, Me.Font.Strikethrough, Me.Font.Charset, 0, 0, 0, 0, Me.Font.Name)
'olfont = SelectObject(hDC, newFont)
'oldColor = SetTextColor(hDC, Me.FontColor)
'''''''draw boundary lines
Call MoveToEx(hdc, Xdisplacement + wWidth, Ydisplacement, ptAPI)
Call LineTo(hdc, Xdisplacement + wWidth, Ydisplacement + hHeight)
Call MoveToEx(hdc, Xdisplacement + wWidth, Ydisplacement, ptAPI)
Call LineTo(hdc, Xdisplacement + 2 * temp, Ydisplacement)

FirstPoint = Me.YMinPosition + m_YInitialMargin
distance = 0#
i = FirstPoint
lowerpoint = True
Do
   distance = (i - FirstPoint) * Me.YUnitToPixels
   If distance > CSng(hHeight) Then Exit Do
   str = Format(i, "#0.00")
   Call MoveToEx(hdc, Xdisplacement + 2 * tmp, hHeight - CLng(distance) + Ydisplacement, ptAPI)
   Call LineTo(hdc, Xdisplacement + wWidth, hHeight - CLng(distance) + Ydisplacement)
   If lowerpoint = True Then
   Call TextOut(hdc, Xdisplacement, hHeight - tmp - 2 + Ydisplacement, str, Len(str))
   lowerpoint = False
   Else
   Call TextOut(hdc, Xdisplacement, hHeight - CLng(distance) + Ydisplacement, str, Len(str))
   End If
   
   i = i + Me.YGap
   '''Call TextOut(hDC, 0, 175, "aaa", 3)
Loop While distance < hHeight

'SelectObject hDC, oldFont
'SelectObject hDC, oldPen
'DeleteObject newFont
'DeleteObject newPen

End Sub

Public Sub SineCurveFill(GraphNo As Long)
Dim i As Single
Dim p As Single
Dim Q As Single
If (GraphNo < 0) Or (GraphNo > MAX_NO_OF_GRAPHS) Then Exit Sub

For i = 0 To 100 Step 0.01
p = 100 + 100 * CSng(Sin(CDbl(i)))
Q = 10 * i
Call AddData(Q, p, GraphNo)
Next i
End Sub

Private Sub VScroll1_Change()
m_YInitialMargin = VScroll1.Max - VScroll1.Value
Label1.Visible = False
UserControl_Paint
End Sub

Private Sub VScroll1_Scroll()
m_YInitialMargin = VScroll1.Max - VScroll1.Value
Label1.Visible = False
UserControl_Paint
End Sub
Private Sub ShowHGrid(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI
Dim str As String
Dim distance As Single
Dim i As Single
Dim tmp As Integer
Dim tmp1 As Long


distance = 0#
i = Me.XGap
tmp = Picture1.DrawStyle
tmp1 = Picture1.ForeColor
Picture1.DrawStyle = vbDot
Picture1.ForeColor = Me.GridColor
Do
   
   distance = (i) * Me.XUnitToPixels
   If distance > CSng(wWidth) Then Exit Do
   
   Call MoveToEx(hdc, Xdisplacement + CLng(distance), Ydisplacement, ptAPI)
   Call LineTo(hdc, Xdisplacement + CLng(distance), Ydisplacement + hHeight)
   
   i = i + Me.XGap
   
Loop While distance < wWidth
Picture1.DrawStyle = tmp
Picture1.ForeColor = tmp1
End Sub
Private Sub ShowVGrid(hdc As Long, Xdisplacement As Long, Ydisplacement As Long, hHeight As Long, wWidth As Long)
Dim ptAPI As POINTAPI
Dim str As String
Dim distance As Single
Dim i As Single
Dim tmp As Integer
Dim tmp1 As Long


distance = 0#
i = 0
tmp = Picture1.DrawStyle
tmp1 = Picture1.ForeColor
Picture1.DrawStyle = vbDot
Picture1.ForeColor = Me.GridColor
Do
   
   distance = (i) * Me.YUnitToPixels
   If distance > CSng(hHeight) Then Exit Do
   
   Call MoveToEx(hdc, Xdisplacement + 1, hHeight - CLng(distance) + Ydisplacement, ptAPI)
   Call LineTo(hdc, Xdisplacement + wWidth, hHeight - CLng(distance) + Ydisplacement)
   
   i = i + Me.XGap
   
Loop While distance < hHeight
Picture1.DrawStyle = tmp
Picture1.ForeColor = tmp1
End Sub
Public Sub PrintMe()
Dim dDrawHeight As Long
Dim dDrawWidth As Long
Dim xextra As Long
Dim yextra As Long
Dim negyextra As Long
Dim Totalpages As Long
Dim i As Long
Dim j As Long
Dim str1 As String
Dim p_PScaleWidth As Long
Dim curRow As Long
Dim curCol As Long
Dim g_XGap As Single
Dim g_YGap As Single
''''' first set all parameters
'' set mode landscape
'PrintOrient (2)
''' set font and forecolor
Printer.Font = Me.Font
Printer.ForeColor = Me.ForeColor
'''set zoom factor
Printer.Zoom = p_zoomfactor
Printer.ScaleMode = 3 '''pixel

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' OnBeginPrinting
Printer.Orientation = 2
'''calculate page width and height
p_PageWidth = GetDeviceCaps(Printer.hdc, 8)
p_PageHeight = GetDeviceCaps(Printer.hdc, 10)
If (p_PageHeight = 0) Or (p_PageWidth = 0) Then
   MsgBox "Error in Printing, Can't Print"
   Exit Sub
End If
  ''' calculate only dimensions of graph without calculating others factors
dDrawWidth = CLng((p_MaxX - p_MinX + 1) * m_XUnitToPixels)
dDrawHeight = CLng((p_MaxY - p_MinY + 1) * m_YUnitToPixels)

''''now calculate  no of pages
p_NumRows = dDrawHeight / p_PageHeight + 1
p_NumCols = dDrawWidth / p_PageWidth + 1
''' now take all in effect
p_PScaleWidth = 100 ''5 * Printer.Font.Size
If p_PScaleWidth < p_YScaleWidth Then p_PScaleWidth = p_YScaleWidth
If p_PScaleWidth < p_XScaleHeight Then p_PScaleWidth = p_XScaleHeight
xextra = p_PScaleWidth + p_XPageMargin
yextra = p_PScaleWidth + p_YPagemargin
If p_Header <> "" Then yextra = yextra + 5 * Printer.Font.Size
If p_XLabel <> "" Then yextra = yextra + 5 * Printer.Font.Size
If p_PrintPageNo <> 0 Then yextra = yextra + 5 * Printer.Font.Size
If p_YLabel <> "" Then xextra = xextra + 5 * Printer.Font.Size

dDrawHeight = dDrawHeight + p_NumRows * p_NumCols * (yextra)
dDrawWidth = dDrawWidth + p_NumCols * p_NumRows * xextra

''''now recalculate  no of pages
If (dDrawHeight Mod p_PageHeight) > 0 Then
p_NumRows = dDrawHeight / p_PageHeight + 1
Else
p_NumRows = dDrawHeight / p_PageHeight
End If
If (dDrawWidth Mod p_PageWidth) > 0 Then
p_NumCols = dDrawWidth / p_PageWidth + 1
Else
 p_NumCols = dDrawWidth / p_PageWidth
 End If
 
''''calculate the graph size on each page
p_XGraphSize = p_PageWidth - xextra  ''' size of x axis
p_YGraphSize = p_PageHeight - yextra    ''' size of y axis

p_XInitialMargin = m_XInitialMargin  ''' save the values
p_YInitialMargin = m_YInitialMargin
Totalpages = p_NumRows * p_NumCols

'''''''''end of OnBeginPrinting
g_XGap = m_XGap
g_YGap = m_YGap
m_XGap = 2 * m_XGap
m_YGap = 2 * m_YGap
''''now start printing
For i = 1 To (Totalpages)

Printer.Print ""
Printer.Font = Me.Font
Printer.ForeColor = Me.ForeColor

xextra = p_XPageMargin
yextra = p_YPagemargin
negyextra = 0

'''' draw header,labels and pageno
If p_PrintPageNo = 1 Then
       Call TextOut(Printer.hdc, CLng(p_PageWidth / 2), yextra, (str(i)), Len((str(i))))
       yextra = yextra + 5 * Printer.Font.Size
ElseIf p_PrintPageNo = 2 Then
       negyextra = negyextra + 5 * Printer.Font.Size
       Call TextOut(Printer.hdc, CLng(p_PageWidth / 2), p_PageHeight - negyextra, str(i), Len(str(i)))
End If

If p_Header <> "" Then
      Call TextOut(Printer.hdc, CLng(Abs((p_PageWidth - Len(p_Header)) / 2)), yextra, p_Header, Len(p_Header))
      yextra = yextra + 5 * Printer.Font.Size
End If
If p_XLabel <> "" Then
       negyextra = negyextra + 5 * Printer.Font.Size
      Call TextOut(Printer.hdc, CLng(Abs(p_PageWidth - Len(p_XLabel) - xextra) / 2 + xextra), p_PageHeight - negyextra, p_XLabel, Len(p_XLabel))
      
End If
If p_YLabel <> "" Then
     For j = 1 To Len(p_YLabel)
     str1 = CStr(Mid(p_YLabel, j, j))
     Call TextOut(Printer.hdc, CLng(xextra), yextra + j * 5 * Printer.Font.Size, str1, Len(str1))
     Next j
     xextra = xextra + 5 * Printer.Font.Size
End If
''''''Now draw scales
Call drawHScale(Printer.hdc, p_PScaleWidth + xextra, p_PageHeight - negyextra - p_PScaleWidth, p_PScaleWidth, p_PageWidth - xextra - p_PScaleWidth)
Call drawVScale(Printer.hdc, xextra, yextra, p_PageHeight - negyextra - p_PScaleWidth - yextra, p_PScaleWidth)
If Me.ShowGrids Then
Call ShowHGrid(Printer.hdc, p_PScaleWidth + xextra, yextra, p_PageHeight - p_PScaleWidth - negyextra - yextra, p_PageWidth - xextra - p_PScaleWidth)
Call ShowVGrid(Printer.hdc, xextra + p_PScaleWidth, yextra, p_PageHeight - p_PScaleWidth - negyextra - yextra, p_PageWidth - xextra - p_PScaleWidth)
End If
For j = 0 To MAX_NO_OF_GRAPHS
Call drawGraph(Printer.hdc, p_PScaleWidth + xextra, yextra, p_PageHeight - p_PScaleWidth - negyextra - yextra, p_PageWidth - xextra - p_PScaleWidth, j)
Next j

If (i Mod p_NumCols) > 0 Then
curRow = i / p_NumCols + 1
Else
curRow = i / p_NumCols
End If
curCol = ((i - 1) Mod p_NumCols) + 1
m_XInitialMargin = CSng((curCol - 1) * p_XGraphSize)
m_YInitialMargin = CSng((curRow - 1) * p_YGraphSize)
Printer.NewPage
Next i

Printer.EndDoc
Printer.EndDoc
'''rexstore the values
m_XInitialMargin = p_XInitialMargin
m_YInitialMargin = p_YInitialMargin
m_XGap = g_XGap
m_YGap = g_YGap
''''in the last set mode portrait
PrintOrient (1)
End Sub
Private Sub PrintOrient(mode As Integer)
    Dim Orient As OrientStructure
    Dim Ret As Integer
    Dim X As Integer
    Printer.Print ""
    Orient.Orientation = mode
    X = Escape(Printer.hdc, 30, Len(Orient), Orient, 0&)
    On Error Resume Next
    Ret = AbortDoc(Printer.hdc)
    On Error Resume Next
    Printer.EndDoc

End Sub
Public Sub PrintSettings(ByVal LeftMarginInPixels As Long, ByVal TopMarginInPixels As Long, ByVal Header As String, ByVal XAxisLabel As String, ByVal Yaxislabel As String, ByVal PrintPageNo As Integer, ByVal zoomFactor As Long, ByVal PaperSize As Long, ByVal printQuality As Long)
p_YPagemargin = LeftMarginInPixels  ''' since we are printing in landscape
p_XPageMargin = TopMarginInPixels
p_zoomfactor = zoomFactor
p_Header = Header
p_XLabel = XAxisLabel
p_YLabel = Yaxislabel
p_PrintPageNo = PrintPageNo  '' 0 means  no print  1 means print at top  2 means print at bottom
Printer.PaperSize = PaperSize
Printer.printQuality = printQuality

End Sub

Public Sub Invalidate()
UserControl_Paint
End Sub
Public Sub PrintOnlyShownPortionAsBitmap()
''''''''this code is stolen from msdn
Const SRCCOPY = &HCC0020
      Const NEWFRAME = 1
      Const PIXEL = 3

MousePointer = 11
Picture1.Picture = Picture1.Image
Printer.ScaleMode = PIXEL
 Printer.Print ""
 hMemoryDC% = CreateCompatibleDC(Picture1.hdc)
 hOldBitMap% = SelectObject(hMemoryDC%, Picture1.Picture)
 
 ApiError% = StretchBlt(Printer.hdc, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight, hMemoryDC%, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, SRCCOPY)
 hOldBitMap% = SelectObject(hMemoryDC%, hOldBitMap%)
 ApiError% = DeleteDC(hMemoryDC%)
 Result% = Escape(Printer.hdc, NEWFRAME, 0, 0&, 0&)
 Printer.EndDoc
 MousePointer = 1
End Sub
