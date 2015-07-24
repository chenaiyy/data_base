VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form20 
   Caption         =   "资料入库"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form20"
   ScaleHeight     =   2070
   ScaleWidth      =   5235
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Connect         =   $"Form20.frx":0000
      OLEDBString     =   $"Form20.frx":0088
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
      Caption         =   "取消"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "资料主题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "资料名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text2.Text = "" Then
    MsgBox "您没有输入资料名称", , "提示"
    GoTo Mark1
End If


If Text4.Text = "" Then
    MsgBox "您没有输入资料主题", , "提示"
    GoTo Mark1
End If

Const a As Long = 0
Const b As Long = 999999
Dim strSNum As String

Mark:
strSNum = "AK47" & CStr(Int(Rnd() * b - a + 1))

Adodc1.RecordSource = "select * from 资料表 where 资料名称='" & strSNum & "'"
Adodc1.Refresh

If Adodc1.Recordset.EOF = True Then
    Adodc1.RecordSource = "资料表"
    Adodc1.Refresh

    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("资料编号").Value = strSNum
    Adodc1.Recordset.Fields("资料名称").Value = Text2.Text
    Adodc1.Recordset.Fields("入库时间").Value = Now
    Adodc1.Recordset.Fields("资料主题").Value = Text4.Text
    Adodc1.Recordset.Update
Else
    GoTo Mark
End If
MsgBox "操作完成", , "提示"

Mark1:
End Sub

Private Sub Command2_Click()
Form20.Hide
Form6.Show
End Sub

