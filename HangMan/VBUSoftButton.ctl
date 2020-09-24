VERSION 5.00
Begin VB.UserControl VBUSoftButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   555
   FillStyle       =   0  'Solid
   ScaleHeight     =   420
   ScaleWidth      =   555
   Begin VB.Image imgPicture 
      Height          =   375
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "VBUSoftButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_AlignCaption = 0
Const m_def_BorderStyle = 1
Const m_def_UseMask = True
Const m_def_MaskColor = 0
Const m_def_DropDownVisible = False
Const m_def_Picture = ""
Const m_def_Caption = ""
'Property Variables:
Dim m_AlignCaption As Integer
Dim m_BorderStyle As scBorderStyle
Dim m_UseMask As Boolean
Dim m_MaskColor As OLE_COLOR
Dim m_DropDownVisible As Boolean
Dim m_Picture As New StdPicture
Dim m_Caption As String
'Event Declarations:
Event DropDown()
Event Enter()
Attribute Enter.VB_Description = "The mouse has just entered the boundaries of the control."
Event Leave()
Attribute Leave.VB_Description = "The mouse has just left the boundaries of the control."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

'Private Variables
Private mbMouseOver As Boolean
Private mbButtonDown As Boolean
Private mbDDDown As Boolean
Private WithEvents LeaveTimer As objTimer
Attribute LeaveTimer.VB_VarHelpID = -1
Private clsPaint As New PaintEffects
Private mlButtonWidth As Long
Private mlButtonHeight As Long
Private mlButtonLeft As Long
Private mlButtonTop As Long
'Enums
Enum scBorderStyle
    None = 0
    [3D] = 1
End Enum

Enum scAlignCaption
    [Below Picture] = 0
    [Beside Picture] = 1
End Enum

Private Sub Leave()
    mbMouseOver = False
    
    Set HoverTimer = Nothing
    Set LeaveTimer = Nothing
    DrawControl
    
    RaiseEvent Leave
    
End Sub
Private Function UnderMouse() As Boolean
    Dim ptMouse As POINTAPI

    GetCursorPos ptMouse
    If WindowFromPoint(ptMouse.x, ptMouse.y) = UserControl.hWnd Then
        UnderMouse = True
    Else
        UnderMouse = False
    End If

End Function


Private Sub LeaveTimer_Timer()
    'Check and see if we're still over the area
    If Not UnderMouse Then Leave

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    DrawControl
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    DrawControl
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    If Not mbDDDown Then
        RaiseEvent Click
    End If
    mbDDDown = False
    DrawControl
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sSoundFile As String
    
    If (x < (mlButtonLeft + mlButtonWidth - 8) Or x > (mlButtonLeft + mlButtonWidth) Or y > mlButtonHeight) Or m_DropDownVisible = False Then 'x <= UserControl.ScaleWidth - 8) Then
        mbButtonDown = True
        mbDDDown = False
        DrawControl
    Else
        mbDDDown = True
        mbButtonDown = False
        DrawControl
        RaiseEvent DropDown
    End If
    'sSoundFile = GetSetting("HKEY_CURRENT_USER\AppEvents\Schemes\Apps", "Office97", "Office97-ToolbarClick", "")
    'sSoundFile = GetRegStr(".Current", "AppEvents\Schemes\Apps\Office97\Office97-ToolbarClick")
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If mbButtonDown And Button = 0 Then
        mbButtonDown = False
    End If
    If mbDDDown And Button = 0 Then
        mbDDDown = False
    End If
    If mbMouseOver Then
        If Not UnderMouse Then
            Leave
        End If
    Else
        If UnderMouse Then
            mbMouseOver = True
            RaiseEvent Enter
            DrawControl
            'Set up the Hover Timer
            'Set HoverTimer = New objTimer
            'HoverTimer.Interval = 500
            'HoverTimer.Enabled = True
            'Set up the Leave Timer
            Set LeaveTimer = New objTimer
            LeaveTimer.Interval = 50
            LeaveTimer.Enabled = True
        End If
    End If
        
    
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mbButtonDown = False
'    mbDDDown = False
'    DrawControl
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Text displayed on the control."
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawControl
End Property
Private Sub DrawControl()
    
    Dim fWidth As Single
    Dim fHeight As Single
    Dim lPicLeft As Long
    Dim lPicTop As Long
    Dim lPicWidth As Long, lPicHeight As Long
    Dim ucWidth As Long
    Dim ucHeight As Long
    Dim ddWidth As Long
    Dim ddHeight As Long
    
    UserControl.Cls
    
    UserControl.ScaleMode = 3
    If m_DropDownVisible Then
        If m_Caption > "" And m_Picture.Width > 1 Then
            ddWidth = 8
            mlButtonWidth = imgPicture.Width + ddWidth + 4
            mlButtonHeight = imgPicture.Height + 3
            hhHeight = UserControl.ScaleHeight
            ucWidth = (imgPicture.Width + 2) + ddWidth - 1
            ucHeight = imgPicture.Height + 3
        Else
            ddWidth = 8
            mlButtonWidth = UserControl.ScaleWidth - 2
            mlButtonHeight = UserControl.ScaleHeight - 2
            hhHeight = UserControl.ScaleHeight
            ucWidth = UserControl.ScaleWidth - ddWidth - 1
            ucHeight = UserControl.ScaleHeight
        End If
    Else
        If m_Caption > "" And m_Picture.Width > 1 Then
            mlButtonWidth = imgPicture.Width + 2
            mlButtonHeight = imgPicture.Height + 3
            ucWidth = imgPicture.Width + 2
            ucHeight = imgPicture.Height + 3
        Else
            mlButtonWidth = UserControl.ScaleWidth - 2
            mlButtonHeight = UserControl.ScaleHeight - 2
            ucWidth = UserControl.ScaleWidth
            ucHeight = UserControl.ScaleHeight
        End If
    End If
    If m_Caption > "" And m_Picture.Width > 1 Then
        If m_AlignCaption = 0 Then
            mlButtonLeft = (UserControl.ScaleWidth / 2) - (mlButtonWidth / 2)
            mlButtonTop = 0
        Else
            mlButtonLeft = 0
            mlButtonTop = (UserControl.ScaleHeight / 2) - (mlButtonHeight / 2)
        End If
    Else
        mlButtonLeft = 0
        mlButtonTop = 0
    End If
    If mbMouseOver And m_BorderStyle = 1 Then
        If mbButtonDown Then
            DrawButtonDown ucWidth, ucHeight
        Else
            DrawButtonUp ucWidth, ucHeight
        End If
        If m_DropDownVisible Then
            If mbDDDown Then
                DrawDDDown ucWidth, ucHeight, ddWidth
            Else
                DrawDDUp ucWidth, ucHeight, ddWidth
            End If
        End If
    End If
    
    'Draw dropdown arrow if needed
    If m_DropDownVisible Then
        lPicTop = (mlButtonHeight / 2) '- 2
        UserControl.Line ((mlButtonLeft + mlButtonWidth - 8) + 1, lPicTop)-((mlButtonLeft + mlButtonWidth - 8) + 7, lPicTop), vbBlack
        UserControl.Line ((mlButtonLeft + mlButtonWidth - 8) + 2, lPicTop + 1)-((mlButtonLeft + mlButtonWidth - 8) + 6, lPicTop + 1), vbBlack
        UserControl.Line ((mlButtonLeft + mlButtonWidth - 8) + 3, lPicTop + 2)-((mlButtonLeft + mlButtonWidth - 8) + 5, lPicTop + 2), vbBlack
        UserControl.Line ((mlButtonLeft + mlButtonWidth - 8) + 4, lPicTop + 3)-((mlButtonLeft + mlButtonWidth - 8) + 4, lPicTop + 3), vbBlack
    End If
    
    If m_DropDownVisible Then
        ddWidth = 8
        hhHeight = UserControl.ScaleHeight
        ucWidth = UserControl.ScaleWidth - ddWidth - 1
        ucHeight = UserControl.ScaleHeight
    Else
        ucWidth = UserControl.ScaleWidth
        ucHeight = UserControl.ScaleHeight
    End If
    If m_Picture.Width > 1 Then
        'UserControl.ScaleMode = 1
        'Set imgPicture.Picture = m_Picture
        If m_Caption > "" And m_AlignCaption = 1 Then
            lPicLeft = 2
        Else
            If m_Caption > "" Then
                lPicLeft = mlButtonLeft + 2 '(ucWidth / 2) - (mlButtonWidth / 2) '(imgPicture.Width / 2)
            Else
                If m_DropDownVisible = True Then
                    lPicLeft = ((mlButtonWidth - 8) / 2) - (imgPicture.Width / 2)
                Else
                    lPicLeft = (mlButtonWidth / 2) - (imgPicture.Width / 2)
                End If
            End If
        End If
        If m_Caption > "" And m_AlignCaption = 0 Then
            lPicTop = 3
        Else
            If m_Caption > "" Then
                lPicTop = mlButtonTop + 3 '(ucHeight / 2) - (mlButtonHeight / 2) '(imgPicture.Height / 2)
            Else
                lPicTop = (mlButtonHeight / 2) - (imgPicture.Height / 2)
            End If
        End If
        lPicWidth = imgPicture.Width
        lPicHeight = imgPicture.Height
        'UserControl.PaintPicture m_Picture, lPicLeft, lPicTop
        If UserControl.Enabled Then
            If m_UseMask Then
                clsPaint.PaintTransparentStdPic UserControl.hdc, lPicLeft, lPicTop, lPicWidth, lPicHeight, m_Picture, 0, 0, m_MaskColor
            Else
                clsPaint.PaintNormalStdPic UserControl.hdc, lPicLeft, lPicTop, lPicWidth, lPicHeight, m_Picture, 0, 0
            End If
        Else
            clsPaint.PaintDisabledStdPic UserControl.hdc, lPicLeft, lPicTop, lPicWidth, lPicHeight, m_Picture, 0, 0, m_MaskColor
        End If
        If m_AlignCaption = 0 Then
            fWidth = UserControl.TextWidth(m_Caption)
            fHeight = UserControl.TextHeight(m_Caption)
            
            UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (fWidth / 2)
            UserControl.CurrentY = lPicHeight + 4
            UserControl.Print m_Caption
        Else
            fHeight = UserControl.TextHeight(m_Caption)
            UserControl.CurrentX = lPicWidth + 4
            If m_DropDownVisible Then
                UserControl.CurrentX = UserControl.CurrentX + 8
            End If
            UserControl.CurrentY = (UserControl.ScaleHeight / 2) - (fHeight / 2)
            UserControl.Print m_Caption
        End If
    Else
        'UserControl.ScaleMode = 1
        fWidth = UserControl.TextWidth(m_Caption)
        fHeight = UserControl.TextHeight(m_Caption)
        
        UserControl.CurrentX = (ucWidth / 2) - (fWidth / 2)
        UserControl.CurrentY = (ucHeight / 2) - (fHeight / 2)
    
        UserControl.Print m_Caption
    End If

End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    m_Caption = m_def_Caption
    'Set imgPicture.Picture = Nothing
    'imgPicture.Width = 0
    'imgPicture.Height = 0
    m_DropDownVisible = m_def_DropDownVisible
    m_MaskColor = m_def_MaskColor
    m_UseMask = m_def_UseMask
    m_BorderStyle = m_def_BorderStyle
    m_AlignCaption = m_def_AlignCaption
End Sub

Private Sub UserControl_Paint()
    DrawControl
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_DropDownVisible = PropBag.ReadProperty("DropDownVisible", m_def_DropDownVisible)
    m_MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)
    m_UseMask = PropBag.ReadProperty("UseMask", m_def_UseMask)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_AlignCaption = PropBag.ReadProperty("AlignCaption", m_def_AlignCaption)
    Set imgPicture.Picture = m_Picture
End Sub

Private Sub UserControl_Resize()
    DrawControl
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("DropDownVisible", m_DropDownVisible, m_def_DropDownVisible)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
    Call PropBag.WriteProperty("UseMask", m_UseMask, m_def_UseMask)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("AlignCaption", m_AlignCaption, m_def_AlignCaption)
End Sub

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Picture to display on the control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    'imgPicture.Width = 0
    'imgPicture.Height = 0
    Set m_Picture = Nothing
    Set m_Picture = New_Picture
    If m_Picture.Width > 1 Then
        'UserControl.ScaleMode = 1
        Set imgPicture.Picture = m_Picture
    End If
    DrawControl
    PropertyChanged "Picture"
End Property

Public Property Get DropDownVisible() As Boolean
    DropDownVisible = m_DropDownVisible
End Property

Public Property Let DropDownVisible(ByVal New_DropDownVisible As Boolean)
    m_DropDownVisible = New_DropDownVisible
    DrawControl
    PropertyChanged "DropDownVisible"
End Property
'
'Public Function ShowDropDown() As Variant
'
'End Function
'
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor
End Property
Private Sub DrawButtonDown(ucWidth As Long, ucHeight As Long)
    Dim lLeft As Long, lTop As Long
    Dim lWidth As Long, lHeight As Long
    
    lLeft = mlButtonLeft + 1
    lTop = mlButtonTop + 1
    lHeight = mlButtonHeight + 1
    If m_DropDownVisible Then
        lWidth = mlButtonWidth - 8 - 2
    Else
        lWidth = mlButtonWidth + 1
    End If
    UserControl.Line (lLeft, lTop)-(lLeft + lWidth, lTop), vb3DShadow
    UserControl.Line (lLeft, lTop)-(lLeft, lTop + lHeight), vb3DShadow
    UserControl.Line (lLeft + lWidth - 1, lTop)-(lLeft + lWidth - 1, lTop + lHeight), vbWhite
    UserControl.Line (lLeft, lTop + lHeight - 1)-(lLeft + lWidth, lTop + lHeight - 1), vbWhite
    
End Sub
Private Sub DrawButtonUp(ucWidth As Long, ucHeight As Long)
    Dim lLeft As Long, lTop As Long
    Dim lWidth As Long, lHeight As Long
    
    lLeft = mlButtonLeft + 1
    lTop = mlButtonTop + 1
    lHeight = mlButtonHeight + 1
    If m_DropDownVisible Then
        lWidth = mlButtonWidth - 8 - 2
    Else
        lWidth = mlButtonWidth + 1
    End If
    UserControl.Line (lLeft, lTop)-(lLeft + lWidth, lTop), vbWhite
    UserControl.Line (lLeft, lTop)-(lLeft, lTop + lHeight), vbWhite
    UserControl.Line (lLeft + lWidth - 1, lTop)-(lLeft + lWidth - 1, lTop + lHeight), vb3DShadow
    UserControl.Line (lLeft, lTop + lHeight - 1)-(lLeft + lWidth, lTop + lHeight - 1), vb3DShadow
    

End Sub
Private Sub DrawDDDown(ucWidth As Long, ucHeight As Long, ddWidth As Long)
    Dim lLeft As Long, lTop As Long
    Dim lWidth As Long, lHeight As Long
    
    lLeft = mlButtonLeft + mlButtonWidth - 9
    lTop = mlButtonTop + 1
    lHeight = mlButtonHeight + 1
    lWidth = ddWidth + 2 'mlButtonWidth + 1
    UserControl.Line (lLeft, lTop)-(lLeft + lWidth, lTop), vb3DShadow
    UserControl.Line (lLeft, lTop)-(lLeft, lTop + lHeight), vb3DShadow
    UserControl.Line (lLeft + lWidth - 1, lTop)-(lLeft + lWidth - 1, lTop + lHeight), vbWhite
    UserControl.Line (lLeft, lTop + lHeight - 1)-(lLeft + lWidth, lTop + lHeight - 1), vbWhite

End Sub
Private Sub DrawDDUp(ucWidth As Long, ucHeight As Long, ddWidth As Long)
    Dim lLeft As Long, lTop As Long
    Dim lWidth As Long, lHeight As Long
    
    lLeft = mlButtonLeft + mlButtonWidth - 9
    lTop = mlButtonTop + 1
    lHeight = mlButtonHeight + 1
    lWidth = ddWidth + 2 'mlButtonWidth + 1
    UserControl.Line (lLeft, lTop)-(lLeft + lWidth, lTop), vbWhite
    UserControl.Line (lLeft, lTop)-(lLeft, lTop + lHeight), vbWhite
    UserControl.Line (lLeft + lWidth - 1, lTop)-(lLeft + lWidth - 1, lTop + lHeight), vb3DShadow
    UserControl.Line (lLeft, lTop + lHeight - 1)-(lLeft + lWidth, lTop + lHeight - 1), vb3DShadow

End Sub
Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    DrawControl
    PropertyChanged "MaskColor"
End Property

Public Property Get UseMask() As Boolean
    UseMask = m_UseMask
End Property

Public Property Let UseMask(ByVal New_UseMask As Boolean)
    m_UseMask = New_UseMask
    DrawControl
    PropertyChanged "UseMask"
End Property

Public Property Get BorderStyle() As scBorderStyle
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As scBorderStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get AlignCaption() As Integer
    AlignCaption = m_AlignCaption
End Property

Public Property Let AlignCaption(ByVal New_AlignCaption As Integer)
    m_AlignCaption = New_AlignCaption
    PropertyChanged "AlignCaption"
End Property

