VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
   Caption         =   "项目申请界面"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6900
   LinkTopic       =   "Form11"
   ScaleHeight     =   6165
   ScaleWidth      =   6900
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
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
      Connect         =   $"Form11.frx":0000
      OLEDBString     =   $"Form11.frx":008C
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
      Left            =   3960
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   1575
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   6375
   End
   Begin VB.TextBox Text5 
      Height          =   1095
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "项目的申报材料"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "项目的具体内容"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "工作地点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "项目经费"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "项目名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text1.Text = "" Then
    MsgBox "请输入项目名称", , "提示"
    GoTo Label1
End If

If Text2.Text = "" Then
    MsgBox "请输入项目经费", , "提示"
    GoTo Label1
End If

If Text3.Text = "" Then
    MsgBox "请输入项目工作地点", , "提示"
    GoTo Label1
End If

If Text5.Text = "" Then
    MsgBox "请输入项目具体内容", , "提示"
    GoTo Label1
End If

If Text6.Text = "" Then
    MsgBox "请填写项目的申报材料", , "提示"
    GoTo Label1
End If

Adodc1.RecordSource = "项目申请表"
Adodc1.Refresh
Adodc1.Recordset.AddNew

Adodc1.Recordset.Fields("项目名称") = Text1.Text
Adodc1.Recordset.Fields("项目经费") = Format(Text2.Text)
Adodc1.Recordset.Fields("工作地点") = Text3.Text
Adodc1.Recordset.Fields("项目申请人职工号") = Form2.Label8.Caption
Adodc1.Recordset.Fields("项目具体内容") = Text5.Text
Adodc1.Recordset.Fields("项目申报材料") = Text6.Text
Adodc1.Recordset.Update

MsgBox "项目申报完成", , "提示"
Form11.Hide
Form5.Show

Label1:

End Sub

Private Sub Command2_Click()
Form11.Hide
Form5.Show
End Sub
