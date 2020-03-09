VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "字体设置"
   ClientHeight    =   5040
   ClientLeft      =   5070
   ClientTop       =   375
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9030
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ButtonChange 
      Caption         =   "换一句"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   21
      Top             =   4740
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9737
            Text            =   "制造本程序：Crazy Urus"
            TextSave        =   "制造本程序：Crazy Urus"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Crazy Urus"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2205
            MinWidth        =   2205
            Text            =   "No.2006001"
            TextSave        =   "No.2006001"
            Object.Tag             =   ""
            Object.ToolTipText     =   "2006年第一个程序"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2734
            MinWidth        =   2734
            TextSave        =   "2020/3/9 星期一"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "1:26"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "查看效果"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   8535
      Begin VB.OptionButton OptionColor 
         Caption         =   "蓝色"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5640
         TabIndex        =   22
         Top             =   3360
         Width           =   855
      End
      Begin ComctlLib.Slider ScrollFontSize 
         Height          =   360
         Left            =   1155
         TabIndex        =   2
         Top             =   2595
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   635
         _Version        =   327682
         LargeChange     =   1
         Min             =   12
         Max             =   72
         SelStart        =   12
         TickStyle       =   3
         Value           =   12
      End
      Begin MSComDlg.CommonDialog ColorDialog 
         Left            =   7800
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "设置字体颜色"
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "自定义"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6600
         TabIndex        =   13
         Top             =   3360
         Width           =   1815
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "绿色"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   12
         Top             =   3360
         Width           =   855
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "黄色"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   11
         Top             =   3360
         Width           =   855
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "橙色"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   10
         Top             =   3360
         Width           =   855
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "红色"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   9
         Top             =   3360
         Width           =   855
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "黑色"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   3360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.ComboBox ComboFont 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2955
         Width           =   2295
      End
      Begin VB.CheckBox CheckStrike 
         Caption         =   "删除线(&S)"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   7
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CheckBox CheckUnderline 
         Caption         =   "下划线(&U)"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   6
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CheckBox CheckItalic 
         Caption         =   "倾斜(&I)"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   3000
         Width           =   1100
      End
      Begin VB.CheckBox CheckBold 
         Caption         =   "加粗(&B)"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   3000
         Width           =   1100
      End
      Begin VB.TextBox TextFontSize 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         TabIndex        =   15
         Text            =   "12"
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox TextOutput 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2175
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   14
         Text            =   "FormMain.frx":0000
         ToolTipText     =   "演示文本"
         Top             =   240
         Width           =   8295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "字号："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "颜色："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "字体："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   540
      End
   End
   Begin VB.TextBox TextInput 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Crazy Urus 追求卓越软件"
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label Label3 
      Caption         =   "请输入要显示的汉字："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   100
      Width           =   2175
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prevOption As Integer

Function SetFontSize(ByVal fontSize As Integer)
    TextOutput.Font.Size = fontSize
    TextFontSize.text = fontSize
    ScrollFontSize.Value = fontSize
End Function

Function SetText(ByVal text As String)
    TextInput.text = text
    TextOutput.text = text
End Function

Private Sub ButtonChange_Click()
    Dim XMLHTTP As Object
    ButtonChange.Enabled = False
     
    Set XMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    XMLHTTP.Open "GET", "https://v1.hitokoto.cn/?c=i&encode=text", False
    XMLHTTP.Send
    If XMLHTTP.Status = 200 And XMLHTTP.readyState = 4 Then
        SetText XMLHTTP.responseText
    Else
        MsgBox "获取内容失败，错误码：" & XMLHTTP.Status, vbExclamation
    End If
    
    ButtonChange.Enabled = True
    Set XMLHTTP = Nothing
End Sub

Private Sub CheckBold_Click()
    If CheckBold.Value = 1 Then TextOutput.FontBold = True Else TextOutput.FontBold = False
End Sub

Private Sub CheckItalic_Click()
    If CheckItalic.Value = 1 Then TextOutput.FontItalic = True Else TextOutput.FontItalic = False
End Sub

Private Sub CheckUnderline_Click()
    If CheckUnderline.Value = 1 Then TextOutput.FontUnderline = True Else TextOutput.FontUnderline = False
End Sub

Private Sub CheckStrike_Click()
    If CheckStrike.Value = 1 Then TextOutput.FontStrikethru = True Else TextOutput.FontStrikethru = False
End Sub

Private Sub ComboFont_Click()
    TextOutput.Font.Name = ComboFont.text
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim fontName As String
    Dim fontNameFirstAscii As Integer
    
    For i = 0 To Screen.FontCount - 1
        fontName = Screen.Fonts(i)
        fontNameFirstAscii = AscW(Mid(fontName, 1, 1))
        
        If (fontNameFirstAscii > 255 Or fontNameFirstAscii < -255) Then
            ComboFont.AddItem fontName
        End If
    Next i
    
    ComboFont.text = TextOutput.Font.Name
End Sub

Private Sub ScrollFontSize_Scroll()
   SetFontSize ScrollFontSize.Value
End Sub

Private Sub OptionColor_Click(Index As Integer)
    Select Case Index
        Case 0
            TextOutput.ForeColor = vbBlack
        Case 1
            TextOutput.ForeColor = 1770192  '红色 #D0021B
        Case 2
            TextOutput.ForeColor = 2336501  '橙色 #F5A623
        Case 3
            TextOutput.ForeColor = 1894392  '黄色 #F8E71C
        Case 4
            TextOutput.ForeColor = 12772176  '绿色 #50E3C2
        Case 5
            TextOutput.ForeColor = 16748568  '蓝色 #1890FF
        Case 6
            On Error GoTo Cancel
            ColorDialog.Action = 3
            
            TextOutput.ForeColor = ColorDialog.Color
            OptionColor(6).Caption = "自定义：#" & DecToHex(ColorDialog.Color)
    End Select
    
    prevOption = Index
    Exit Sub
Cancel:
    OptionColor(prevOption).Value = True
End Sub

Private Sub TextInput_Change()
    TextOutput.text = TextInput.text
End Sub

Private Sub TextFontSize_Change()
    Dim fontSize As Integer
    Dim prevFontSize As Integer
    
    prevFontSize = TextOutput.Font.Size
    fontSize = Val(TextFontSize.text)
    
    If fontSize >= 12 And fontSize <= 72 Then
       SetFontSize fontSize
    ElseIf fontSize <= 12 And fontSize > 0 Then
    Else
        MsgBox fontSize & " 为无效属性值 (12-72)", vbExclamation
        TextFontSize.text = prevFontSize
    End If
End Sub
