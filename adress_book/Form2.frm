VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "ѧ��������Ϣ"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15495
   LinkTopic       =   "Form2"
   ScaleHeight     =   5070
   ScaleWidth      =   15495
   StartUpPosition =   3  '����ȱʡ
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   4335
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   840
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ѧ����Ϣ��\Student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ѧ����Ϣ��\Student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from ѧ����Ϣ��"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "ѡ����"
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.Frame Frame1 
         Caption         =   "ѡ����"
         Height          =   4455
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2535
         Begin VB.CheckBox Check1 
            Caption         =   "����"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   24
            Top             =   2880
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "�Ա�"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "רҵ"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "ѧ��"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   7
            ItemData        =   "Form2.frx":0015
            Left            =   960
            List            =   "Form2.frx":006D
            Sorted          =   -1  'True
            TabIndex        =   20
            Text            =   "����"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   6
            ItemData        =   "Form2.frx":0135
            Left            =   960
            List            =   "Form2.frx":018D
            Sorted          =   -1  'True
            TabIndex        =   19
            Text            =   "ѧ��"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   5
            ItemData        =   "Form2.frx":02FD
            Left            =   960
            List            =   "Form2.frx":0307
            TabIndex        =   18
            Text            =   "רҵ"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   4
            ItemData        =   "Form2.frx":032D
            Left            =   960
            List            =   "Form2.frx":0337
            TabIndex        =   17
            Text            =   "�Ա�"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "����"
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   3600
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "��ѯ"
            Height          =   495
            Left            =   1320
            TabIndex        =   15
            Top             =   3600
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Height          =   270
            Left            =   1080
            TabIndex        =   14
            Top             =   3000
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   270
            Left            =   1800
            TabIndex        =   13
            Top             =   3000
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "����"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   735
         End
         Begin VB.Line Line2 
            X1              =   1440
            X2              =   1800
            Y1              =   3120
            Y2              =   3120
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "����"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1800
         TabIndex        =   9
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1080
         TabIndex        =   8
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��ѯ"
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   3
         ItemData        =   "Form2.frx":0343
         Left            =   960
         List            =   "Form2.frx":034D
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   2
         ItemData        =   "Form2.frx":0359
         Left            =   960
         List            =   "Form2.frx":0363
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         ItemData        =   "Form2.frx":0389
         Left            =   960
         List            =   "Form2.frx":03E1
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         ItemData        =   "Form2.frx":0551
         Left            =   960
         List            =   "Form2.frx":05A9
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   1440
         X2              =   1800
         Y1              =   3120
         Y2              =   3120
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Private Sub Command4_Click()
Unload Form2             '���ص�������
Form1.Show
End Sub


Private Sub Command3_Click()
Sql = "select ѧ��,����,�Ա�,רҵ,����,������¼,����  from ѧ����Ϣ�� where ѧ�� is not null"
If Check1(1).Value = 1 Then
Sql = Sql + " and ���� = '" & Combo1(7).Text & "'" '��������ѯ
End If
If Check1(2).Value = 1 Then
Sql = Sql + " and ѧ��= '" & Combo1(6).Text & "'" '��ѧ�Ų�ѯ
End If
If Check1(3).Value = 1 Then
Sql = Sql + " and רҵ= '" & Combo1(5).Text & "'" '��רҵ��ѯ
End If
If Check1(4).Value = 1 Then
Sql = Sql + " and �Ա�= '" & Combo1(4).Text & "'" '���Ա��ѯ
End If
If Check1(5).Value = 1 Then
Sql = Sql + " and ���� between'" & Text4.Text & "' and'" & Text3.Text & "'"
Sql = Sql + "order by ���� ASC"    '�������ѯ
End If
Adodc1.RecordSource = Sql
Adodc1.Refresh
DataGrid1.Refresh

End Sub


