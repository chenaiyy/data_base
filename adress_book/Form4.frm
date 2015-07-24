VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "综合查询"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form4"
   ScaleHeight     =   5355
   ScaleWidth      =   11490
   StartUpPosition =   3  '窗口缺省
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":0000
      Height          =   4335
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=学生信息表\Student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=学生信息表\Student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from 学生信息表"
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
      Caption         =   "选择项"
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   2040
         TabIndex        =   11
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1200
         TabIndex        =   10
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Caption         =   "平均成绩"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Top             =   1560
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "年龄"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "性别"
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "女"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "男"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "返回"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "查询"
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   3600
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   2040
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   1680
         Y1              =   1680
         Y2              =   1680
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public I As Integer
Dim Sql As String
Dim db As Database
Dim gradelow As Integer
Dim gradehigh As Integer

Private Sub Command1_Click()
On Error GoTo LABEL
Set db = OpenDatabase("学生信息表\Student.mdb")
Sql = "select 学生课程表.学号,AVG(成绩) as 平均成绩 into Temp from 学生课程表,选课表 where 学生课程表.课程代号=选课表.课程代号 GROUP BY 学生课程表.学号"
db.Execute (Sql)      '建立一个临时表，来存储平均成绩的信息，相当于视图
db.Close

Sql = "select 学生信息表.学号,姓名,专业,性别,年龄,平均成绩 from Temp,学生信息表 where 学生信息表.学号=Temp.学号"
If Check2.Value = 1 Then         '如果点击了平均成绩，则加入平均成绩的查询信息
gradelow = Int(Text3.Text)
gradehigh = Int(Text4.Text)
Sql = Sql + " and 平均成绩 between " & gradelow & " and " & gradehigh
End If
If Check1.Value = 1 Then         '如果点击了年龄，则加入平均成绩的查询信息
Sql = Sql + " and 年龄 between'" & Text1.Text & "' and'" & Text2.Text & "'"
End If
'If Frame2.Enabled = 1 Then
Select Case I
Case 0
Sql = Sql + " and 性别='男'"     '性别查询
Case 1
Sql = Sql + " and 性别='女'"
End Select
'End If
Sql = Sql + " order by 平均成绩 desc"

Adodc1.RecordSource = Sql
Adodc1.Refresh
DataGrid1.Refresh

LABEL:
Set db = OpenDatabase("学生信息表\Student.mdb")
Sql = "DROP Table Temp"
db.Execute (Sql)
db.Close
End Sub

Private Sub Command2_Click()
Unload Form4            '关闭当前窗口，返回主窗口
Form1.Show
End Sub

Private Sub Option1_Click(Index As Integer)
I = Index
End Sub
