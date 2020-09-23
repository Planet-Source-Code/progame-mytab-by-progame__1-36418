VERSION 5.00
Begin VB.UserControl myTab 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   MouseIcon       =   "myTab.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   2055
   ScaleWidth      =   4395
   Begin VB.CommandButton cmdFocus 
      Caption         =   "Command1"
      Height          =   195
      Left            =   -500
      TabIndex        =   0
      Top             =   -200
      Width           =   510
   End
   Begin VB.Image imgBar 
      Height          =   225
      Left            =   135
      Picture         =   "myTab.ctx":0152
      Stretch         =   -1  'True
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "myTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long

'Default Property Values:
Const m_def_BorderLine = True
Const m_def_HoverActiveTabTextColor = &HFF8080
Const m_def_Style = 1
Const m_def_Align = 1
Const m_def_TabHeight = 200
Const m_def_TabWidthMax = 0
Const m_def_TabTextColor = vbBlack
Const m_def_TabBackColor = &H8000000A
Const m_def_ActiveTabTextColor = vbBlack
Const m_def_ActiveTabBackColor = vbWhite
Const m_def_HoverTabTextColor = &H8000&
'Property Variables:
Dim m_BorderLine As Boolean
Dim m_ImageList As Object
Dim m_HoverActiveTabTextColor As OLE_COLOR
Dim m_Style As Variant
Dim m_Align                 As typeAlign
Dim m_TabHeight             As Integer
Dim m_TabWidthMax           As Integer
Dim m_TabTextColor          As OLE_COLOR
Dim m_TabBackColor          As OLE_COLOR
Dim m_ActiveTabTextColor    As OLE_COLOR
Dim m_ActiveTabBackColor    As OLE_COLOR
Dim m_HoverTabTextColor     As OLE_COLOR

'Const IMAGESCALE   As Single = 15       '*Í¼Ïñ¿í¶È»»³ÉtwipsµÄ×ª»»

Public Enum typeStyle
    tyTabOnBottom = 0
    tyTabOnTop = 1
End Enum

Public Enum typeAlign
    tyLeft = 0
    tyMiddle = 1
    tyRight = 2
End Enum

Private m_Tabs()            As clsTab

Private WithEvents m_Tab    As clsTab
Attribute m_Tab.VB_VarHelpID = -1

Private m_TabCount      As Long

Private m_LeftTab       As Long         '*×î×óµÄTAB

'*ÔÚMouseUpÊ±¼ÇÂ¼Êó±êÎ»ÖÃ£¬¸øClickºÍDblClickÊ¹ÓÃ
Private m_button        As Integer
Private m_shift         As Integer
Private m_X             As Single
Private m_Y             As Single

'**********************************
'*¶¨ÒåÊÂ¼þ
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event tabChange(key As String)
Public Event SelectChang(previous As String, current As String)
Public Event Hover(key As String)
Public Event Click(key As String)
Public Event DblClick(key As String)
'
Public Property Get Tabs(ByVal index) As clsTab
Attribute Tabs.VB_MemberFlags = "400"
'*¶ÔtabµÄÒýÓÃ
Dim key         As String
Dim mIndex      As Long
    key = CStr(index)
    If IsNumeric(key) Then
        mIndex = CLng(key)
    Else
        mIndex = Key2Index(key)
    End If
    If mIndex > 0 And mIndex <= m_TabCount Then
        Set Tabs = m_Tabs(mIndex)
        Set m_Tab = Tabs
    Else
        Set Tabs = Nothing
    End If
End Property

Public Function TabCount() As Long
'*·µ»ØTABÊý
    TabCount = m_TabCount
End Function


'**************************************************************
'*Ãû³Æ£ºAddTab
'*¹¦ÄÜ£ºÌí¼ÓÒ»¸öTAB
'*´«Èë²ÎÊý£º
'*      key         --key
'*      caption     --caption
'*      width       --tabwidth,if not set,autofit
'*      tooltiptext --tooltiptext where hover this tab
'*      pretabindex --add current tab behind this tab
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame
'*ÈÕÆÚ£º2002-6-26 12:34:50
'**************************************************************
Public Function AddTab(key As String, _
                       Caption As String, _
                       Optional Width As Single = -1, _
                       Optional Image As String = "", _
                       Optional ToolTipText As String = "", _
                       Optional preTabIndex As Long = -1) _
    As Boolean
    
Dim lCur            As Long
Dim i               As Long
    
    AddTab = False

    On Error GoTo 0
    
    '*keyÖµ²»ÄÜÎª¿Õ£¬²»ÄÜÖØ¸´
    If key = "" Then
        Err.Raise vbObjectError + 513, , "Invalid key ---  Null!"
        Exit Function
    End If
    If ExistsTab(key) Then
        Err.Raise vbObjectError + 513, , "Invalid key ---  Repeat!"
        Exit Function
    End If
    '*keyÖµ²»ÄÜÎªÊýÖµ
    If IsNumeric(key) Then
        Err.Raise vbObjectError + 513, , "Invalid key ---  Numeric!"
        Exit Function
    End If
    '*ÖØÐÂ¶¨ÒåÊý×é
    ReDim Preserve m_Tabs(1 To m_TabCount + 1)
    Set m_Tabs(m_TabCount + 1) = New clsTab
    
    m_TabCount = m_TabCount + 1
    
    lCur = m_TabCount                        '*µ±Ç°TABµÄÎ»ÖÃ
    
    If preTabIndex <> -1 Then       '*´ÓÖÐ¼ä²åÈë£¬·ñÔòÄ¬ÈÏ²åÈëµ½×îºó
        If preTabIndex > m_TabCount - 1 Or preTabIndex < 1 Then
            AddTab = False
            Exit Function
        End If
        '*½«ËùÓÐÖ®ºóµÄtabÖ®indexºóÒÆ
        For i = m_TabCount To preTabIndex + 2 Step -1
            With m_Tabs(i)
                .index = i
                .key = m_Tabs(i - 1).key
                .Caption = m_Tabs(i - 1).Caption
                .Image = m_Tabs(i - 1).Image
                .Width = m_Tabs(i - 1).Width
                .ToolTipText = m_Tabs(i - 1).ToolTipText
                .Active = m_Tabs(i - 1).Active
                .Hover = m_Tabs(i - 1).Hover
            End With
        Next i
        lCur = preTabIndex + 1          '*µ±Ç°TABµÄÎ»ÖÃ
    End If
    
    '*Ð´Èë´ËTABµÄÐÅÏ¢
    With m_Tabs(lCur)
        .index = lCur
        .key = key
        .Caption = Caption
        .Image = Image
        '*Èç¹ûÊÇ´«µÝµÄÊÇimagelistµÄindex£¬½«Ö®×ª»»µ½key
        If IsNumeric(.Image) And Not (m_ImageList Is Nothing) Then
            .Image = m_ImageList.ListImages(CInt(.Image)).key
        End If
        .Width = Width
        '*µ÷Õû¿í¶È
        AdjustWidth (lCur)
        
        .ToolTipText = ToolTipText
    End With
    
    SelectTab = key
    
    '*È·±£´ËTAB¿É¼û
    MakeTabVisible key
    
    DrawTabs
    
    
End Function

Private Sub MakeTabVisible(key As String)
'*ÉèÖÃ´ËTAB¿É¼û
    Select Case TabVisible(key)
        Case -1             '*×ó²»¿É¼û
            m_LeftTab = Key2Index(key)
        Case 0              '*¿É¼û
            '*²»×ö´¦Àí
        Case 1              '*ÓÒ²»¿É¼û
            '*Ôö¼Ó m_lefttabÖ±ÖÁ¿É¼û
            Do While TabVisible(key) = 1
                m_LeftTab = m_LeftTab + 1
            Loop
    End Select
End Sub

Private Sub AdjustWidth(index As Long)
'*µ÷Õû×ÔÊÊÓ¦¿í¶ÈµÄTABµÄ¿í¶È
Dim tWidth          As Single
    With m_Tabs(index)
        tWidth = .Width
        If .Width = -1 Then
            tWidth = UserControl.TextWidth(.Caption) + 150
            '*Èç¹ûÒªÍ¼Ïñ£¬¼Ó¿í
            If (Not m_ImageList Is Nothing) And (m_Tabs(index).Image <> "") Then
                tWidth = tWidth + TabHeight
            End If
            '*Èç¹ûÓÐ×î´ó¿í¶ÈÏÞ¶¨ÇÒ³¬³ö£¬ÔòÉèÎª×î´ó¿í¶È
            If TabWidthMax > 0 And tWidth > TabWidthMax Then
                tWidth = TabWidthMax
            End If
        End If
        
        .Width = tWidth
    End With
End Sub

'**************************************************************
'*Ãû³Æ£ºRemoveTab
'*¹¦ÄÜ£ºÒÆ³ýÒ»¸öTAB
'*´«Èë²ÎÊý£º
'*      index         --index or key
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame
'*ÈÕÆÚ£º2002-6-26 12:39:22
'**************************************************************
Public Function RemoveTab(ByVal index) As Boolean
'*ÒÆ³ýTAB£¬È»ºóËùÓÐÖ®ºóµÄTABÏòÇ°ÒÆ
Dim mIndex          As Long
Dim key             As String
Dim i               As Long
    RemoveTab = False
    key = CStr(index)
    If key = "" Then
        Exit Function
    End If
    If IsNumeric(key) Then
        '*Èç¹ûÊäÈëµÄÊÇÊý×Ö£¬Ôò×ªÎªkey
        mIndex = CLng(key)
        '*Ð£ÑéºÏ·¨ÐÔ
        If mIndex < 1 Or mIndex > m_TabCount Then
            Exit Function
        End If
    Else
        '*ÊäÈëµÄÊÇkey£¬ÊÇ·ñ´æÔÚÐ£Ñé
        If Not ExistsTab(key) Then
            Exit Function
        End If
        mIndex = Key2Index(key)
    End If
    
    If mIndex <> -1 Then
        '*Èç¹ûµ±Ç°µÄtabÊÇactivetab£¬ÔòÖØÐÂÉèÖÃactivetab
        If m_Tabs(mIndex).Active Then
            '*Èç¹ûÊÇ×îºóÒ»¸ö£¬ÉèÇ°Ò»¸ö
            If mIndex = m_TabCount And m_TabCount > 1 Then
                SetActiveTab m_Tabs(mIndex - 1).key
                MakeTabVisible m_Tabs(mIndex - 1).key
            Else
                If mIndex < m_TabCount Then
                    SetActiveTab m_Tabs(mIndex + 1).key
                    MakeTabVisible m_Tabs(mIndex + 1).key
                End If
            End If
        End If
        
        For i = mIndex To m_TabCount - 1
            With m_Tabs(i)
                .index = i
                .key = m_Tabs(i + 1).key
                .Caption = m_Tabs(i + 1).Caption
                .Width = m_Tabs(i + 1).Width
                .ToolTipText = m_Tabs(i + 1).ToolTipText
                .Active = m_Tabs(i + 1).Active
                .Hover = m_Tabs(i + 1).Hover
            End With
        Next i
        
        '*ÖØÐÂ¶¨ÒåÊý×é
        m_TabCount = m_TabCount - 1
        If m_TabCount = 0 Then
            Erase m_Tabs
        Else
            Set m_Tabs(m_TabCount + 1) = Nothing
            ReDim Preserve m_Tabs(1 To m_TabCount)
        End If
        
        RemoveTab = True
    End If
    DrawTabs
End Function

'**************************************************************
'*Ãû³Æ£ºRemoveAll
'*¹¦ÄÜ£ºÒÆ³ýËùÓÐTAB
'*´«Èë²ÎÊý£º
'*
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame
'*ÈÕÆÚ£º2002-6-26 12:40:01
'**************************************************************
Public Sub RemoveAll()
    
    m_TabCount = 0
    Erase m_Tabs
    DrawTabs
End Sub

'**************************************************************
'*Ãû³Æ£ºExistsTab
'*¹¦ÄÜ£ºÊÇ·ñ´æÔÚ
'*´«Èë²ÎÊý£º
'*
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame
'*ÈÕÆÚ£º2002-6-26 15:32:10
'**************************************************************
Public Function ExistsTab(key As String) As Boolean
Dim i           As Long
    ExistsTab = False
    For i = 1 To m_TabCount
        If m_Tabs(i).key = key Then
            ExistsTab = True
            Exit Function
        End If
    Next i
End Function

'**************************************************************
'*Ãû³Æ£ºKey2Index
'*¹¦ÄÜ£ºÓÉkeyµÃµ½index
'*´«Èë²ÎÊý£º
'*
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame
'*ÈÕÆÚ£º2002-6-26 15:39:12
'**************************************************************
Public Function Key2Index(key As String) As Long
Dim i           As Long
    Key2Index = -1
    For i = 1 To m_TabCount
        If m_Tabs(i).key = key Then
            Key2Index = i
            Exit Function
        End If
    Next i
End Function

'**************************************************************
'*Ãû³Æ£ºIndex2Key
'*¹¦ÄÜ£ºÓÉindexµÃµ½key
'*´«Èë²ÎÊý£º
'*
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame
'*ÈÕÆÚ£º2002-6-26 15:39:12
'**************************************************************
Public Function Index2Key(index As Long) As String
Dim i           As Long
    Index2Key = ""
    If index > 0 And index <= m_TabCount Then
        Index2Key = m_Tabs(index).key
    End If
End Function

'**************************************************************
'*Ãû³Æ£ºDrawTabs
'*¹¦ÄÜ£º»æÖÆ¿Ø¼þ
'*´«Èë²ÎÊý£º
'*
'*·µ»Ø²ÎÊý£º
'*
'*×÷Õß£ºprogame
'*ÈÕÆÚ£º2002-6-26 15:45:44
'**************************************************************
Private Sub DrawTabs()
Dim left            As Single
Dim top             As Single
Dim i               As Long
Dim sWidth          As Single
    UserControl.Cls
    '*ÏÈÅÐ¶ÏÊÇ·ñÄÜ¹»È«²¿ÏÔÊ¾
    If UserControl.Width < TabsWidth Then
        '*ÏÔÊ¾bar
        imgBar.Visible = True
        left = 915
    Else
        imgBar.Visible = False
        left = 0
    End If
    
    If Style = tyTabOnBottom Then
        top = UserControl.Height - TabHeight
    Else
        top = 0
    End If
    
    sWidth = left
    If m_LeftTab - 1 > 0 Then
        '*Èç¹ûÓÐ¹¤¾ßÌõ³öÏÖ£¬»æÖÆÄÇÐ¡°ë±ß²»¿É¼ûµÄtab
        DrawTab m_LeftTab - 1, left - m_Tabs(m_LeftTab - 1).Width - TabHeight / 2, top
    End If
    Dim aLeft       As Single       '*activetabµÄleft
    Dim aWidth      As Single       '*activetabµÄwidth
    For i = m_LeftTab To m_TabCount
        DrawTab i, left, top
        If m_Tabs(i).Active Then
            aLeft = left
            aWidth = m_Tabs(i).Width + TabHeight
        End If
        sWidth = sWidth + m_Tabs(i).Width + TabHeight / 2
        If sWidth + TabHeight / 2 >= UserControl.Width Then
            Exit For
        End If
        left = left + m_Tabs(i).Width + TabHeight / 2
    Next i
    '*»æÖÆ±ßÏß
    If BorderLine Then
        Line (0, top + IIf(Style = tyTabOnTop, 0, TabHeight - 10))-Step(UserControl.Width, 0), &H80000010
        '*ÓÃactivetabµÄµ×É«¸²¸Ç±ßÏß
        Line (aLeft, top + IIf(Style = tyTabOnTop, 0, TabHeight - 10))-Step(aWidth, 0), ActiveTabBackColor
    End If
End Sub

Public Function TabVisible(ByVal index) As Integer
'*TABÊÇ·ñ¿É¼û,-1×ó²»¿É¼û£¬0¿É¼û,1ÓÒ²»¿É¼û,-2ÎªÃ»ÓÐ´Ëtab
Dim sWidth          As Single
Dim mIndex           As Long
Dim i               As Long
    If IsNumeric(index) Then
        mIndex = CLng(index)
    Else
        mIndex = Key2Index(CStr(index))
    End If
    If mIndex < 0 Or mIndex > m_TabCount Then
        TabVisible = -2
    End If
    '*Ê×ÏÈÅÐ¶ÏÊÇ·ñÏÔÊ¾ÁË°´Å¥
    If TabsWidth > UserControl.Width Then
        sWidth = imgBar.Width
    Else
        sWidth = 0
    End If
    If mIndex < m_LeftTab Then
        TabVisible = -1
        Exit Function
    End If
    For i = m_LeftTab To mIndex
        sWidth = sWidth + m_Tabs(i).Width + TabHeight / 2
        If sWidth + TabHeight / 2 >= UserControl.Width Then
            TabVisible = 1
            Exit Function
        End If
    Next i
    TabVisible = 0
End Function

'Private Sub DrawBar()
''*»æÖÆ¹¤¾ßÌõ
'Dim tFontName       As String       '*ÔÝ´æ×ÖÌåÃû³Æ
'Dim tFontSize       As Single       '*ÔÝ´æ×ÖÌå´óÐ¡
'Dim tColor          As OLE_COLOR    '*ÔÝ´æ×ÖÌåÑÕÉ«
'
'    With UserControl
'        tFontName = .FontName
'        tFontSize = .FontSize
'        tColor = .ForeColor
'        .FontName = "Marlett"
'        .FontSize = 12
'        .ForeColor = &H404040
'        '*´òÓ¡
'        .CurrentX = 50
'        .CurrentY = 0
'        Print "3"
'        .CurrentX = 250
'        .CurrentY = 0
'        Print "3"
'        .CurrentX = 400
'        .CurrentY = 0
'        Print "4"
'        .CurrentX = 600
'        .CurrentY = 0
'        Print "4"
'        .FontSize = 11
'        .FontName = "Arial"
'        .CurrentX = 0
'        .CurrentY = 0
'        Print "|"
'        '*»Ö¸´
'        .FontName = tFontName
'        .FontSize = tFontSize
'        .ForeColor = tColor
'    End With
'End Sub

Private Sub DrawTab(index As Long, left As Single, top As Single)
'*ÔÚleft,topÆðµã¿ªÊ¼»æÖÆÒ»¸öTab£¬ÆÕÍ¨µÄ
Dim tStr            As String
Dim tHeight         As Single
Dim tWidth          As Single
Dim X               As Single
Dim Y               As Single
Dim bIcon           As Boolean
    With UserControl
        If m_Tabs(index).Hover Then
            .FontUnderline = True
        Else
            .FontUnderline = False
        End If
    End With
    
    bIcon = ((Not m_ImageList Is Nothing) And m_Tabs(index).Image <> "")
    '*µÃµ½Òª»æÖÆµÄ×Ö·û´®
    tStr = GetFitStr(m_Tabs(index).Caption, m_Tabs(index).Width, bIcon)
    
    '*µÃµ½´òÓ¡Êä³öµÄ¸ß¶ÈºÍ¿í¶È
    tHeight = UserControl.TextHeight(tStr)
    tWidth = UserControl.TextWidth(tStr)
    
    Y = top + (TabHeight - tHeight) / 2
    
    Dim imageWidth          As Single
    '*Èç¹ûÒªÍ¼Ïñ
    If bIcon Then
        imageWidth = TabHeight
    Else
        imageWidth = 0
    End If
    
    Select Case Align
        Case tyLeft
            X = left + imageWidth + 30
        Case tyMiddle
            X = left + (m_Tabs(index).Width - tWidth + imageWidth) / 2
        Case tyRight
            X = left + (m_Tabs(index).Width - tWidth)
    End Select

    '*»æÖÆ
    With UserControl
        If m_Tabs(index).Active Then
            .ForeColor = ActiveTabTextColor
        Else
            .ForeColor = TabTextColor
        End If
        If m_Tabs(index).Hover Then
            .ForeColor = HoverTabTextColor
            If m_Tabs(index).Active Then
                .ForeColor = HoverActiveTabTextColor
            End If
        End If
        '*»æÖÆ±³¾°
        If Style = tyTabOnTop Then
            DrawBackGroundTop index, left, top
        Else
            DrawBackGroundBottom index, left, top
        End If
        '*Êä³ö×Ö·û´®
        .CurrentX = X + TabHeight / 2
        .CurrentY = Y
        Print tStr
        '*»æÖÆÍ¼Ïñ
        If m_Tabs(index).Image <> "" Then
            If m_Tabs(index).Hover Or m_Tabs(index).Active Then
                .PaintPicture m_ImageList.ListImages(m_Tabs(index).Image).Picture, _
                    left + TabHeight / 2 + 30, top + 30, _
                    TabHeight, _
                    TabHeight
            End If
        End If
    End With
    
End Sub

Private Sub DrawBackGroundTop(index As Long, left As Single, top As Single)
Dim i           As Integer
Dim color       As OLE_COLOR
Dim preActive   As Boolean          '*Ç°Ò»¸öTABÊÇ·ñÎªactive
    preActive = False
    If index > 1 Then
        If m_Tabs(index - 1).Active Then
                preActive = True
        End If
    End If
    If m_Tabs(index).Active Then
        color = ActiveTabBackColor
    Else
        color = TabBackColor
    End If
    '*»æÖÆTAB±³¾°ÑÕÉ«
    Line (left + TabHeight / 2, top)-Step(m_Tabs(index).Width, TabHeight - 20), color, BF
    For i = 2 To TabHeight / 20 - 2
        '*Õâ¸öTAB×ó²¿²»±»ÕÚµµ
        If m_Tabs(index).Active Or m_LeftTab = index Then
            Line (left + i * 10, top)-Step(10, (i - 1) * 20), color, BF
        Else
            Line (left + TabHeight / 4 + i * 5 + 10, top + TabHeight / 2 - i * 10)-Step(5, (i - 1) * 20), color, BF
        End If
        
        Line (left + m_Tabs(index).Width + TabHeight / 2 + (i - 1) * 10, top)-Step(10, TabHeight - i * 20), color, BF
    Next i
    '*»æÖÆ×ó¶ËÏß¿ò
    If m_Tabs(index).Active Or m_LeftTab = index Then
        Line (left, top)-Step(TabHeight / 2, TabHeight), vbBlack
        Line (left + 15, top)-Step(TabHeight / 2, TabHeight), vbWhite
    Else
        Line (left + TabHeight / 4, top + TabHeight / 2)-Step(TabHeight / 4, TabHeight / 2), vbBlack
        Line (left + TabHeight / 4 + 15, top + TabHeight / 2)-Step(TabHeight / 4, TabHeight / 2), vbWhite
    End If
    '*»æÖÆÓÒ¶ËÏß¿ò
    Line (left + TabHeight / 2 + m_Tabs(index).Width, top + TabHeight)-Step(TabHeight / 2, -TabHeight), vbBlack
    Line (left + TabHeight / 2 + m_Tabs(index).Width - 15, top + TabHeight)-Step(TabHeight / 2, -TabHeight), &H80000010

    '*»æÖÆµ×²¿»ò¶¥²¿Ïß
    Line (left + TabHeight / 2, top + TabHeight - 5)-Step(m_Tabs(index).Width, 0), &H80000010
End Sub

Private Sub DrawBackGroundBottom(index As Long, left As Single, top As Single)
Dim i           As Integer
Dim color       As OLE_COLOR
Dim preActive   As Boolean          '*Ç°Ò»¸öTABÊÇ·ñÎªactive
    preActive = False
    If index > 1 Then
        If m_Tabs(index - 1).Active Then
                preActive = True
        End If
    End If
    If m_Tabs(index).Active Then
        color = ActiveTabBackColor
    Else
        color = TabBackColor
    End If
    '*»æÖÆTAB±³¾°ÑÕÉ«
    Line (left + TabHeight / 2, top)-Step(m_Tabs(index).Width, TabHeight - 20), color, BF
    For i = 2 To TabHeight / 20 - 2
        '*Õâ¸öTAB×ó²¿²»±»ÕÚµµ
        If m_Tabs(index).Active Or m_LeftTab = index Then
            Line (left + i * 10, top + TabHeight)-Step(10, -(i - 1) * 20), color, BF
        Else
            Line (left + TabHeight / 4 + i * 5, top + TabHeight / 2 - i * 10)-Step(5, (i - 1) * 20), color, BF
        End If
        
        Line (left + m_Tabs(index).Width + TabHeight / 2 + (i - 1) * 10, top + i * 20)-Step(10, TabHeight - i * 20), color, BF
    Next i
    '*»æÖÆ×ó¶ËÏß¿ò
    If m_Tabs(index).Active Or m_LeftTab = index Then
        Line (left, top + TabHeight)-Step(TabHeight / 2, -TabHeight), vbBlack
        Line (left + 15, top + TabHeight)-Step(TabHeight / 2, -TabHeight), vbWhite
    Else
        Line (left + TabHeight / 4, top + TabHeight / 2)-Step(TabHeight / 4, -TabHeight / 2), vbBlack
        Line (left + TabHeight / 4 + 15, top + TabHeight / 2)-Step(TabHeight / 4, -TabHeight / 2), vbWhite
    End If
    '*»æÖÆÓÒ¶ËÏß¿ò
    Line (left + TabHeight / 2 + m_Tabs(index).Width, top)-Step(TabHeight / 2, TabHeight), vbBlack
    Line (left + TabHeight / 2 + m_Tabs(index).Width - 15, top)-Step(TabHeight / 2, TabHeight), &H80000010
    '*»æÖÆµ×²¿»ò¶¥²¿Ïß
    Line (left + TabHeight / 2, top)-Step(m_Tabs(index).Width, 0), &H80000010
End Sub


Private Function GetFitStr(Caption As String, Width As Single, bIcon As Boolean) As String
'*·µ»Ø¿ÉÒÔÊä³öµ½widthÄÚµÄcaption²¿·Ö
Dim tStr            As String
Dim i               As Integer
Dim tWidth          As Single
    GetFitStr = ""
    tWidth = Width
    '*Èç¹ûÒªÍ¼ÏñÏÔÊ¾£¬Ôò¿í¶ÈÒª¼õÉÙ
    If bIcon Then
        tWidth = tWidth - TabHeight
    End If
    
    For i = 1 To Len(Caption)
        If UserControl.TextWidth(left(Caption, i)) <= tWidth Then
            GetFitStr = left(Caption, i)
        Else
            Exit Function
        End If
    Next i
End Function

Private Function TabsWidth() As Single
'*·µ»ØËùÓÐtabºÏ¼Æ¿í¶È
Dim i           As Long
    If m_TabCount > 0 Then
        TabsWidth = TabHeight / 2
    Else
        TabsWidth = 0
    End If
    For i = 1 To m_TabCount
        If m_Tabs(i).Width = -1 Then
            m_Tabs(i).Width = UserControl.TextWidth(m_Tabs(i).Caption) + 150
            If m_Tabs(i).Width > TabWidthMax Then
                m_Tabs(i).Width = TabWidthMax
            End If
        End If
        TabsWidth = TabsWidth + m_Tabs(i).Width + TabHeight / 2
    Next i
End Function

Private Sub SetActiveTab(key As String)
'*ÉèÖÃ»î¶¯µÄTAB
Dim i           As Long
Dim last        As String
    last = SelectTab
    
    For i = 1 To m_TabCount
        m_Tabs(i).Active = False
        If m_Tabs(i).key = key Then
            m_Tabs(i).Active = True
        End If
    Next i
    
    '*½«´ËTABÉèÎª¿É¼û
    MakeTabVisible key
    
    DrawTabs
    
    '*´¥·¢ÊÂ¼þ
    RaiseEvent SelectChang(last, key)
End Sub

Private Sub SetHoverTab(key As String)
'*ÉèÖÃÊó±êÍ£ÁôµÄTAB
Dim i           As Long
    For i = 1 To m_TabCount
        m_Tabs(i).Hover = False
        If m_Tabs(i).key = key Then
            m_Tabs(i).Hover = True
        End If
    Next i
End Sub

Private Sub NoHoverTab()
'*Êó±êÒÆ³ö£¬Ã»ÓÐhoverµÄTAB
Dim i           As Long
    For i = 1 To m_TabCount
        m_Tabs(i).Hover = False
    Next i
End Sub

Public Property Get SelectTab()
'*µÃµ½µ±Ç°»î¶¯µÄTAB
Dim i           As Long
    SelectTab = ""
    For i = 1 To m_TabCount
        If m_Tabs(i).Active Then
            SelectTab = m_Tabs(i).key
            Exit Property
        End If
    Next i
End Property

Public Property Let SelectTab(index)
'*ÉèÖÃ»î¶¯µÄTAB
Dim key         As String
    If IsNumeric(index) Then
        key = Index2Key(CStr(index))
    Else
        key = CStr(index)
    End If
    If ExistsTab(key) Then
        SetActiveTab key
        DrawTabs
    End If
End Property

Public Property Get TabTextColor() As OLE_COLOR
    TabTextColor = m_TabTextColor
End Property

Public Property Let TabTextColor(ByVal New_TabTextColor As OLE_COLOR)
    m_TabTextColor = New_TabTextColor
    PropertyChanged "TabTextColor"
End Property

Public Property Get TabBackColor() As OLE_COLOR
    TabBackColor = m_TabBackColor
End Property

Public Property Let TabBackColor(ByVal New_TabBackColor As OLE_COLOR)
    m_TabBackColor = New_TabBackColor
    PropertyChanged "TabBackColor"
End Property


Public Property Get ActiveTabTextColor() As OLE_COLOR
    ActiveTabTextColor = m_ActiveTabTextColor
End Property

Public Property Let ActiveTabTextColor(ByVal New_ActiveTabTextColor As OLE_COLOR)
    m_ActiveTabTextColor = New_ActiveTabTextColor
    PropertyChanged "ActiveTabTextColor"
End Property

Public Property Get ActiveTabBackColor() As OLE_COLOR
    ActiveTabBackColor = m_ActiveTabBackColor
End Property

Public Property Let ActiveTabBackColor(ByVal New_ActiveTabBackColor As OLE_COLOR)
    m_ActiveTabBackColor = New_ActiveTabBackColor
    PropertyChanged "ActiveTabBackColor"
End Property

Public Property Get HoverTabTextColor() As OLE_COLOR
    HoverTabTextColor = m_HoverTabTextColor
End Property

Public Property Let HoverTabTextColor(ByVal New_HoverTabTextColor As OLE_COLOR)
    m_HoverTabTextColor = New_HoverTabTextColor
    PropertyChanged "HoverTabTextColor"
End Property


Private Sub imgBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*ÒÆ¶¯
    '*ÅÐ¶ÏÑ¡ ÖÐµÄÊÇÄÄ¸ö°´Å¥
    If X > 0 And X < 228 Then
        m_LeftTab = 1
    End If
    If X >= 228 And X <= 457 Then
        m_LeftTab = m_LeftTab - 1
        If m_LeftTab < 0 Then
            m_LeftTab = 1
        End If
    End If
    If X > 457 And X <= 685 Then
        m_LeftTab = m_LeftTab + 1
        If m_LeftTab > m_TabCount Then
            m_LeftTab = m_TabCount
        End If
    End If
    If X > 685 And X <= 915 Then
        m_LeftTab = m_TabCount
    End If
    If m_LeftTab = 0 Then
        m_LeftTab = 1
    End If
    DrawTabs
End Sub

Private Sub m_Tab_tabChange(key As String, item As typeChange)
'*tabÄÚÈÝ¸ü¸Ä£¬ÖØÐÂ»æÖÆ¿Ø¼þ
    Select Case item
        Case tbKey                              '*ÐÞ¸ÄÁËKeyÖµ£¬Òª²é¿´ÊÇ·ñÖØ¸´
            If RepeatKey(m_Tab.key) Then
                m_Tab.key = key         '*»Øµ½Ô­Ê¼×´Ì¬
                '*´íÎó
                Err.Raise vbObjectError + 513, , "Repeat key"
            End If
            Exit Sub
        Case tbWidth                            '*ÐÞ¸ÄÁËwidth
            '*ÐÞ¸ÄwidthÎªÊµ¼ÊµÄÖµ
            AdjustWidth m_Tab.index
        Case tbCaption                          '*ÐÞ¸ÄÁËcaption
            
        Case tbImage                            '*ÐÞ¸ÄÁËimage
            '*Èç¹ûÊÇ´«µÝµÄÊÇimagelistµÄindex£¬½«Ö®×ª»»µ½key
            If IsNumeric(m_Tab.Image) And Not (m_ImageList Is Nothing) Then
                m_Tab.Image = m_ImageList.ListImages(CInt(m_Tab.Image)).key
            End If
        Case tbToolTipText                      '*ÐÞ¸ÄÁËtooltiptext
    End Select
    '*´¥·¢ÊÂ¼þ
    RaiseEvent tabChange(key)
    DrawTabs
End Sub

Private Function RepeatKey(key As String) As Boolean
'*ÊÇ·ñÓÐÖØ¸´µÄkeyÖµ
Dim i           As Long
Dim cnt         As Integer
    cnt = 0
    RepeatKey = False
    For i = 1 To m_TabCount
        If m_Tabs(i).key = key Then
            cnt = cnt + 1
            If cnt = 2 Then
                RepeatKey = True
            End If
        End If
    Next i
End Function

Private Sub UserControl_Click()
'*´¥·¢ÊÂ¼þ
Dim key         As String
    If m_TabCount = 0 Then
        RaiseEvent Click("")
    Else
        Dim iHover          As Long
        iHover = GetMouseTab(m_X, m_Y)
        If iHover = -1 Then
            
            RaiseEvent Click("")
        Else
            key = m_Tabs(iHover).key
            SetActiveTab key
            RaiseEvent Click(key)
        End If
    End If
End Sub

Private Sub UserControl_DblClick()
'*´¥·¢ÊÂ¼þ
Dim key         As String
    If m_TabCount = 0 Then
        RaiseEvent DblClick("")
    Else
        Dim iHover          As Long
        iHover = GetMouseTab(m_X, m_Y)
        If iHover = -1 Then
            RaiseEvent DblClick("")
        Else
            key = m_Tabs(iHover).key
            RaiseEvent DblClick(key)
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    m_TabCount = 0
    m_LeftTab = 1
    Set m_ImageList = Nothing
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_TabTextColor = m_def_TabTextColor
    m_TabBackColor = m_def_TabBackColor
    m_ActiveTabTextColor = m_def_ActiveTabTextColor
    m_ActiveTabBackColor = m_def_ActiveTabBackColor
    m_HoverTabTextColor = m_def_HoverTabTextColor
    Set UserControl.Font = Ambient.Font
    m_TabHeight = m_def_TabHeight
    m_TabWidthMax = m_def_TabWidthMax
    m_Align = m_def_Align
    m_Style = m_def_Style
    m_HoverActiveTabTextColor = m_def_HoverActiveTabTextColor
    m_BorderLine = m_def_BorderLine
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*ÅÐ¶ÏÊó±êÊÇ·ñÍ£ÁôÔÚ¿Ø¼þÉÏ
Dim bOver           As Boolean

    If (X >= 0 And X <= UserControl.Width And Y >= 0 And Y <= UserControl.Height) Then
        bOver = True
    Else
        bOver = False
    End If
    If (imgBar.Visible And X <= 915) Then
        bOver = False
    End If
    '*½«ËùÓÐµÄHoverÇåÎªfalse
    NoHoverTab
    
    If bOver Then
        SetCapture UserControl.hWnd
        Dim iHover          As Long
        iHover = GetMouseTab(X, Y)
        
        If iHover <> -1 Then
            UserControl.MousePointer = 99
            SetHoverTab m_Tabs(iHover).key
        Else
            UserControl.MousePointer = 0
        End If
    Else
        UserControl.MousePointer = 0
        ReleaseCapture
    End If
    DrawTabs
    
    '*´¥·¢ÊÂ¼þ
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Function GetMouseTab(X As Single, Y As Single) As Long
'*·µ»ØÊó±êËùÖ¸µÄtab
Dim sWidth          As Single
Dim i               As Long
    GetMouseTab = -1
    sWidth = IIf(TabsWidth > UserControl.Width, 915, 0)
    '*´Óm_LeftTab¿ªÊ¼µ½²»¿É¼ûÎªÖ¹
    sWidth = sWidth + TabHeight / 2
    For i = m_LeftTab To m_TabCount
        If (X >= sWidth And X <= sWidth + m_Tabs(i).Width + TabHeight / 2) Then
            
            '*´¥·¢ÊÂ¼þ
            RaiseEvent Hover(m_Tabs(i).key)
            
            GetMouseTab = i
            Exit Function
        End If
        sWidth = sWidth + m_Tabs(i).Width + TabHeight / 2
        If sWidth > UserControl.Width Then
            Exit Function
        End If
    Next i
End Function

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_button = Button
    m_shift = Shift
    m_X = X
    m_Y = Y
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TabTextColor = PropBag.ReadProperty("TabTextColor", m_def_TabTextColor)
    m_TabBackColor = PropBag.ReadProperty("TabBackColor", m_def_TabBackColor)
    m_ActiveTabTextColor = PropBag.ReadProperty("ActiveTabTextColor", m_def_ActiveTabTextColor)
    m_ActiveTabBackColor = PropBag.ReadProperty("ActiveTabBackColor", m_def_ActiveTabBackColor)
    m_HoverTabTextColor = PropBag.ReadProperty("HoverTabTextColor", m_def_HoverTabTextColor)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_TabHeight = PropBag.ReadProperty("TabHeight", m_def_TabHeight)
    m_TabWidthMax = PropBag.ReadProperty("TabWidthMax", m_def_TabWidthMax)
    m_Align = PropBag.ReadProperty("Align", m_def_Align)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_HoverActiveTabTextColor = PropBag.ReadProperty("HoverActiveTabTextColor", m_def_HoverActiveTabTextColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000003)
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
    m_BorderLine = PropBag.ReadProperty("BorderLine", m_def_BorderLine)
End Sub

Private Sub UserControl_Resize()
    imgBar.Height = UserControl.Height
    imgBar.left = 0
    imgBar.top = 0
End Sub

Private Sub UserControl_Terminate()
'*Çå³ý¶ÔÏóºÍÊý×é
    Erase m_Tabs
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("TabTextColor", m_TabTextColor, m_def_TabTextColor)
    Call PropBag.WriteProperty("TabBackColor", m_TabBackColor, m_def_TabBackColor)
    Call PropBag.WriteProperty("ActiveTabTextColor", m_ActiveTabTextColor, m_def_ActiveTabTextColor)
    Call PropBag.WriteProperty("ActiveTabBackColor", m_ActiveTabBackColor, m_def_ActiveTabBackColor)
    Call PropBag.WriteProperty("HoverTabTextColor", m_HoverTabTextColor, m_def_HoverTabTextColor)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("TabHeight", m_TabHeight, m_def_TabHeight)
    Call PropBag.WriteProperty("TabWidthMax", m_TabWidthMax, m_def_TabWidthMax)
    Call PropBag.WriteProperty("Align", m_Align, m_def_Align)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("HoverActiveTabTextColor", m_HoverActiveTabTextColor, m_def_HoverActiveTabTextColor)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000003)
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)

    Call PropBag.WriteProperty("BorderLine", m_BorderLine, m_def_BorderLine)
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property


Public Property Get TabHeight() As Integer
    TabHeight = m_TabHeight
End Property

Public Property Let TabHeight(ByVal New_TabHeight As Integer)
'*À¹½Ø²»ºÏ·¨µÄtab¸ß¶È
    If New_TabHeight < 200 Then
        New_TabHeight = 200
    End If
    If New_TabHeight > UserControl.Height Then
        New_TabHeight = UserControl.Height
    End If
    m_TabHeight = New_TabHeight
    PropertyChanged "TabHeight"
End Property

Public Property Get TabWidthMax() As Integer
    TabWidthMax = m_TabWidthMax
End Property

Public Property Let TabWidthMax(ByVal New_TabWidthMax As Integer)
    m_TabWidthMax = New_TabWidthMax
    PropertyChanged "TabWidthMax"
End Property

Public Property Get Align() As typeAlign
    Align = m_Align
End Property

Public Property Let Align(ByVal New_Align As typeAlign)
    m_Align = New_Align
    PropertyChanged "Align"
    DrawTabs
End Property


Public Property Get Style() As typeStyle
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As typeStyle)
    m_Style = New_Style
    PropertyChanged "Style"
    DrawTabs
End Property

Public Property Get HoverActiveTabTextColor() As OLE_COLOR
    HoverActiveTabTextColor = m_HoverActiveTabTextColor
End Property

Public Property Let HoverActiveTabTextColor(ByVal New_HoverActiveTabTextColor As OLE_COLOR)
    m_HoverActiveTabTextColor = New_HoverActiveTabTextColor
    PropertyChanged "HoverActiveTabTextColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    'imgBar.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ImageList() As Object
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As Object)
    Set m_ImageList = New_ImageList
    PropertyChanged "ImageList"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get BorderLine() As Boolean
    BorderLine = m_BorderLine
End Property

Public Property Let BorderLine(ByVal New_BorderLine As Boolean)
    m_BorderLine = New_BorderLine
    PropertyChanged "BorderLine"
    DrawTabs
End Property

