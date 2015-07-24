VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "个人信息"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form2"
   ScaleHeight     =   3930
   ScaleWidth      =   5235
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3000
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":0088
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "资料申请"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "进入办公"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改个人密码"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form2.frx":0110
      OLEDBString     =   $"Form2.frx":0198
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "个人信息"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1080
         TabIndex        =   19
         Top             =   2160
         Width           =   90
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "当月工资："
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1200
         TabIndex        =   10
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "所属部门："
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "性别："
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "职务："
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "职工号："
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "欢迎您："
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   855
      Left            =   3120
      TabIndex        =   23
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   705
      Left            =   3120
      TabIndex        =   20
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Label16"
      Height          =   180
      Left            =   3600
      TabIndex        =   17
      Top             =   3600
      Width           =   630
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "北京时间："
      Height          =   180
      Left            =   3240
      TabIndex        =   16
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Label14"
      Height          =   225
      Left            =   3600
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "公司动态："
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label12 
      Height          =   735
      Left            =   3120
      TabIndex        =   13
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
Form2.Hide
End Sub

Private Sub Command2_Click()
Dim bumen As String
bumen = Label11.Caption

Form2.Hide

Select Case bumen
    Case "办公室"
        Form4.Show
    Case "科研处"
        Form5.Show
    Case "信息中心"
        Form6.Show
    Case "综合处"
        Form7.Show
End Select
End Sub

Private Sub Command3_Click()
Adodc2.RecordSource = "select 借阅职工职工号 from 资料外借信息表 where 借阅职工职工号='" & Label8.Caption & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF = True Then
    Form2.Hide
    Form21.Show
Else
    MsgBox "对不起，你上次借阅的资料还没有归还，您暂时不能借阅任何资料", , "提示"
End If
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Timer1.Interval = 300
 Dim a As Integer
    a = Weekday(Now(), vbMonday)
    Select Case a
          Case 1
            Label16.Caption = "星期一"
          Case 2
            Label16.Caption = "星期二"
          Case 3
            Label16.Caption = "星期三"
          Case 4
            Label16.Caption = "星期四"
          Case 5
            Label16.Caption = "星期五"
          Case 6
            Label16.Caption = "星期六"
          Case 7
            Label16.Caption = "星期日"
    End Select
Dim number As String
number = Form1.Text1
Adodc1.RecordSource = "select 职工号,姓名,性别,职务,所在部门,工资 from 基本信息表 where 职工号='" & number & "'"
Adodc1.Refresh

Label2.Caption = Adodc1.Recordset.Fields("姓名").Value
Label8.Caption = number
Label9.Caption = Adodc1.Recordset.Fields("职务").Value
Label10.Caption = Adodc1.Recordset.Fields("性别").Value
Label11.Caption = Adodc1.Recordset.Fields("所在部门").Value
Dim wage As Double
wage = Adodc1.Recordset.Fields("工资").Value
Label18.Caption = Format(wage, "￥0.00")

Adodc1.RecordSource = "动态表"
Adodc1.Refresh
Adodc1.Recordset.MoveLast

Label12.Caption = Adodc1.Recordset.Fields("内容").Value

Adodc2.RecordSource = "select 资料名称 from 资料外借信息表 where 借阅职工职工号='" & number & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF = False Then
    Label19.Caption = Adodc2.Recordset.Fields("资料名称").Value
    Label19.Caption = "您借阅的" + Label19.Caption + "已经获得批准，请及时到信息中心提取该资料"
    Label19.Visible = True
Else
    Label19.Visible = False
End If

Adodc2.RecordSource = "select 项目名称 from 已完成项目表 where 项目申请人职工号='" & number & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF = False Then
    Label7.Caption = Adodc2.Recordset.Fields("项目名称").Value
    Label7.Caption = "您提交的项目“" + Label7.Caption + "”已经获得批准，请及时到财务处领取相关项目资金"
    Label7.Visible = True
Else
    Label7.Visible = False
End If

End Sub

Private Sub Timer1_Timer()
Label14.Caption = Now()
End Sub
