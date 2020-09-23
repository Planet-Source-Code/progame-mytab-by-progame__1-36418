VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTest 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   2085
   ClientTop       =   2895
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   5835
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   135
      TabIndex        =   11
      Top             =   4005
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   688
      ButtonWidth     =   635
      ButtonHeight    =   582
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   1
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Border Line"
      Height          =   285
      Left            =   900
      TabIndex        =   16
      Top             =   2205
      Width           =   1545
   End
   Begin prjmyTab.myTab myTab 
      Height          =   250
      Left            =   90
      TabIndex        =   15
      Top             =   90
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverActiveTabTextColor=   33023
      BackColor       =   -2147483648
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Get Tab"
      Height          =   330
      Left            =   4050
      TabIndex        =   14
      Top             =   1530
      Width           =   1230
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Change Width"
      Height          =   285
      Left            =   4050
      TabIndex        =   13
      Top             =   1170
      Width           =   1230
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add Icon"
      Height          =   330
      Left            =   4050
      TabIndex        =   12
      Top             =   810
      Width           =   1230
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove All"
      Height          =   330
      Left            =   2700
      TabIndex        =   10
      Top             =   2205
      Width           =   1275
   End
   Begin VB.ComboBox cmbActive 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1710
      Width           =   1635
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modify Tab"
      Height          =   330
      Left            =   2700
      TabIndex        =   5
      Top             =   1845
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert Tab"
      Height          =   330
      Left            =   2700
      TabIndex        =   4
      Top             =   1500
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove Tab"
      Height          =   330
      Left            =   2700
      TabIndex        =   3
      Top             =   1155
      Width           =   1275
   End
   Begin VB.ComboBox cmbStyle 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1260
      Width           =   1635
   End
   Begin VB.ComboBox cmbAlign 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "dasdf"
      Top             =   810
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Caption         =   "Add Tab"
      Height          =   330
      Left            =   2700
      TabIndex        =   0
      Top             =   810
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3150
      Top             =   3780
      Width           =   240
   End
   Begin ComctlLib.ImageList imgIcon 
      Left            =   900
      Top             =   3645
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTest.frx":0000
            Key             =   "book"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Active"
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Style"
      Height          =   195
      Left            =   225
      TabIndex        =   8
      Top             =   1305
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Align"
      Height          =   195
      Left            =   270
      TabIndex        =   7
      Top             =   855
      Width           =   345
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    myTab.BorderLine = IIf(Me.Check1.Value = 1, True, False)
End Sub

Private Sub cmbActive_Click()
'*ÉèÖÃactive
    '*by key
    'myTab.SelectTab = myTab.Index2Key(cmbActive.ListIndex + 1)
    '*by index
    myTab.SelectTab = cmbActive.ListIndex + 1
End Sub

Private Sub cmbActive_DropDown()
'*×°ÔØËùÓÐµÄtab
Dim i           As Long
    cmbActive.Clear
    With myTab
        For i = 1 To .TabCount
           cmbActive.AddItem .Tabs(.Index2Key(i)).Caption
        Next i
    End With
End Sub

Private Sub cmbAlign_Click()
    myTab.Align = cmbAlign.ListIndex
End Sub

Private Sub cmbStyle_Click()
    myTab.Style = cmbStyle.ListIndex
End Sub

Private Sub Command1_Click()
    With myTab
        .AddTab "tab" & .TabCount + 1, "TAB" & .TabCount + 1 ', , 1   ', '800
    End With
End Sub



Private Sub Command2_Click()
    With myTab
        .AddTab "insert", "Insert", 1000, "", , 1
    End With
End Sub

Private Sub Command3_Click()
Dim i       As Integer
    i = 1
    With myTab
        .RemoveTab i
    End With
End Sub



Private Sub Command4_Click()
    myTab.Tabs(1).Caption = "Modified"
End Sub

Private Sub Command5_Click()
    myTab.RemoveAll
End Sub

Private Sub Command6_Click()
    myTab.Tabs("tab1").Image = 1
    myTab.SelectTab = 1
End Sub

Private Sub Command7_Click()
    myTab.Tabs(1).Width = 2000
End Sub

Private Sub Command8_Click()
    MsgBox "tab 1 caption:" & myTab.Tabs(1).Caption & vbCrLf _
            & "     reference by index" & vbCrLf _
            & "tab 1 key:" & myTab.Tabs("tab1").key & vbCrLf _
            & "     reference by key"
End Sub

Private Sub Form_Load()
    With cmbAlign
        .AddItem "Left"
        .AddItem "Middle"
        .AddItem "Right"
        .Text = .List(myTab.Align)
    End With
    With cmbStyle
        .AddItem "TabOnBottom"
        .AddItem "TabOnTop"
        .Text = .List(myTab.Style)
    End With
    Me.Check1.Value = IIf(myTab.BorderLine, 1, 0)
    Set myTab.ImageList = imgIcon
    Set Me.Toolbar1.ImageList = imgIcon
    Me.Toolbar1.Buttons(1).Image = "book"
    Me.Image1.Picture = imgIcon.ListImages("book").Picture
    
End Sub


Private Sub myTab_Click(key As String)
    Debug.Print "Event: click  Key:" & key
End Sub

Private Sub myTab_DblClick(key As String)
    Debug.Print "Event: double click  Key:" & key
End Sub

Private Sub myTab_Hover(key As String)
    myTab.ToolTipText = key
    Debug.Print "Event: hover  Key:" & key
End Sub


Private Sub myTab_tabChange(key As String)
    Debug.Print "Event: tab change  Key:" & key
End Sub
