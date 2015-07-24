VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "登录界面"
   ClientHeight    =   3360
   ClientLeft      =   2115
   ClientTop       =   555
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   6780
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   0
      Picture         =   "Form1.frx":52FC
      ScaleHeight     =   2235
      ScaleWidth      =   2835
      TabIndex        =   8
      Top             =   600
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3960
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   $"Form1.frx":80D2
      OLEDBString     =   $"Form1.frx":815A
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
   Begin VB.CommandButton Command2 
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "验证码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "提示：用户名为职工号，密码初始为888888.请在第一次登入后修改您的密码。"
      Height          =   180
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   6210
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "办公系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vCode As String
Private Sub drawvc() '显示校验码
Dim i, vc, px, py As Long
Dim r, g, b As Byte

Randomize                       '初始化随机种子
vc = CLng(8999 * Rnd + 1000)    '生成随机校验码
vCode = vc

Picture2.Cls                    '显示校验码
Picture2.Print vc
'添加噪点（防止自动图像识别）
For i = 0 To 2000               '画2000个噪点
    px = CLng(Picture2.Width * Rnd)  '画点随机位置
    py = CLng(Picture2.Height * Rnd)
    r = CByte(255 * Rnd)         '画点随机颜色
    g = CByte(255 * Rnd)
    b = CByte(255 * Rnd)
    Picture2.Line (px, py)-(px + 1, py + 1), RGB(r, g, b)
Next
End Sub

Private Sub Command1_Click()
Dim name As String
Dim password As String

Adodc1.RecordSource = "select 职工号,密码 from 基本信息表"
Adodc1.Refresh

If Text1.Text = "" Then
    MsgBox "请输入用户名", , "提示"
    GoTo Label
Else
    name = Text1.Text
End If

If Text2.Text = "" Then
    MsgBox "您没有输入密码", , "提示"
    GoTo Label
Else
    password = Text2.Text
End If

If Text3.Text = "" Then
    MsgBox "您没有输入验证码", , "提示"
    GoTo Label
End If

Adodc1.Recordset.MoveFirst

Do While Adodc1.Recordset.EOF <> True
    If name = Adodc1.Recordset.Fields("职工号").Value Then
        If password = Adodc1.Recordset.Fields("密码").Value And Text3.Text = vCode Then
            Form2.Show
            Unload Form1
        ElseIf Text3.Text <> vCode Then
            MsgBox "验证码错误", , "提示"
            Text2.Text = ""
            Text3.Text = ""
            drawvc
        Else
            MsgBox "密码错误", , "提示"
            Text2.Text = ""
            Text3.Text = ""
            drawvc
        End If
        GoTo Label
    Else
        Adodc1.Recordset.MoveNext
    End If
Loop

MsgBox "用户名不存在！", , "提示"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
drawvc

Label:
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Form_Load()
Picture2.FontSize = 14
Picture2.FontBold = True
Picture2.AutoRedraw = True
drawvc
End Sub
