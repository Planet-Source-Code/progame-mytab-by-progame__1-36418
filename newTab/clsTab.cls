VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Key           As String
Private m_Index         As Long
Private m_Caption       As String
Private m_Width         As Single
Private m_ToolTipText   As String
Private m_Image         As String
Private m_Active        As Boolean
Private m_Hover         As Boolean

Public Event tabChange(key As String, item As typeChange)

Public Enum typeChange      '*tabÐÞ¸ÄµÄÊÇÄÄ¸öÊôÐÔ
    tbKey = 0
    tbCaption = 1
    tbWidth = 2
    tbToolTipText = 3
    tbImage = 4
End Enum

Public Property Get key() As String
    key = m_Key
End Property

Public Property Let key(vData As String)
Dim tStr            As String
    If vData <> "" Then
        If Not IsNumeric(vData) Then
            tStr = m_Key
            m_Key = vData
            RaiseEvent tabChange(tStr, tbKey)
        End If
    End If
End Property

Public Property Get index() As Long
    index = m_Index
End Property

Public Property Let index(vData As Long)
    m_Index = vData
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(vData As String)
    m_Caption = vData
    RaiseEvent tabChange(key, tbCaption)
End Property

Public Property Get Width() As Single
    Width = m_Width
End Property

Public Property Let Width(vData As Single)
'* -1 means autofit width
Dim tWidth          As Single
    tWidth = m_Width
    m_Width = vData
    
    If index <> 0 And vData <> tWidth Then
        RaiseEvent tabChange(key, tbWidth)
    End If
    
End Property

Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(vData As String)
    m_ToolTipText = vData
    RaiseEvent tabChange(key, tbToolTipText)
End Property

Public Property Get Image() As String
    Image = m_Image
End Property

Public Property Let Image(vData As String)
    m_Image = vData
    RaiseEvent tabChange(key, tbImage)
End Property

Public Property Get Active() As Boolean
    Active = m_Active
End Property

Friend Property Let Active(vData As Boolean)
    m_Active = vData
End Property

Public Property Get Hover() As Boolean
    Hover = m_Hover
End Property

Friend Property Let Hover(vData As Boolean)
    m_Hover = vData
End Property

Private Sub Class_Initialize()
    Width = -1
    Active = False
    Hover = False
End Sub

Private Sub Class_Terminate()
'
End Sub
