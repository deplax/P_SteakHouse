VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Menu 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   0  '없음
   Caption         =   "Menu"
   ClientHeight    =   6735
   ClientLeft      =   11475
   ClientTop       =   8490
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin prjFrmSkinV8.frmSkinV8 frmSkinV81 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11880
      Caption         =   "Menu"
      Begin prjFrmSkinV8.jcbutton CmdClose 
         Height          =   735
         Left            =   2040
         TabIndex        =   18
         Top             =   5520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Close"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton CmdDel 
         Height          =   615
         Left            =   2400
         TabIndex        =   17
         Top             =   3840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1085
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Del"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton CmdSave 
         Height          =   615
         Left            =   1320
         TabIndex        =   16
         Top             =   3840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1085
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Save"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton CmdAdd 
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   3840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1085
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Add"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5535
         Left            =   3480
         TabIndex        =   14
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   9763
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
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
               LCID            =   1042
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
               LCID            =   1042
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
      Begin VB.TextBox Text4 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   1440
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   1440
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   1440
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         Height          =   180
         Left            =   1440
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Wine"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Steak"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   1440
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   2640
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   6720
         Top             =   6240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Portion"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "SteakName"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Price"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "SteakCode"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "StockCode"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "StockCode"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   2640
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   3210
         Left            =   120
         Picture         =   "Menu.frx":0000
         Top             =   3460
         Width           =   3210
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dB As Connection
Dim RS As Recordset
Dim currentR As Long
Dim totalR As Long

Private Sub CmdAdd_Click()
    Adodc1.Recordset.AddNew
    Dim countR As Integer
    Combo1.Clear
    
    If Option1(1).Value = True Then
        sql = "select StockName from StockT where StockCode like 'W%' order by StockCode"
    Else
        sql = "select StockName from StockT where StockCode like 'S%' order by StockCode"
    End If
    RS.Open sql, dB, adOpenStatic, adLockOptimistic
    
    countR = RS.RecordCount
    For i = 1 To countR
        Combo1.AddItem RS.Fields("StockName")
        RS.MoveNext
    Next
    RS.Close
    

End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdDel_Click()
Dim YBt As Integer
    
    If Not (Adodc1.Recordset.BOF And Adodc1.Recordset.EOF) Then
    YBt = MsgBox("삭제하시겠습니까?", vbYesNo, "Delete")
        If YBt = vbYes Then
            Adodc1.Recordset.Delete
            Adodc1.Recordset.MoveNext
            If Adodc1.Recordset.EOF Then Adodc1.Recordset.MovePrevious
        Else
            DataGrid1.SetFocus
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    If Not (Adodc1.Recordset.EOF And Adodc1.Recordset.EOF) Then
        If Adodc1.Recordset.EditMode = adEditAdd Then
            Adodc1.Recordset.Update
        Else
            Adodc1.Recordset.Update
        End If
    End If
End Sub

Private Sub Combo1_Click()
    
        sql = "select StockCode from StockT where StockName = '" & Combo1.Text & "'"
        RS.Open sql, dB, adOpenStatic, adLockOptimistic
        Label1(1).Caption = RS.Fields("StockCode")
        RS.Close

End Sub

Private Sub DataGrid1_Click()
    ComboLoad
End Sub

Private Sub Form_Load()

    Adodc1.ConnectionString = "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
    
    Dim sql As String
    Set dB = New Connection
    dB.CursorLocation = adUseClient
    dB.Open "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"

    SteakMode
    
End Sub
Private Sub SteakMode()
    SetClear

    Adodc1.RecordSource = "select * from SteakT order by SteakCode"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh
    Set RS = New Recordset
'    RS.Requery
    
    Label1(2).Caption = "SteakCode"
    Label1(4).Caption = "SteakName"
    Label1(3).Caption = "Price"
    Label1(0).Caption = "StockCode"
    Label1(5).Caption = "Portion"
    Label1(1).Caption = "StockCode"
    
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Combo1.Text = ""
    
    Set Text1.DataSource = Adodc1
    Set Text2.DataSource = Adodc1
    Set Text3.DataSource = Adodc1
    Set Text4.DataSource = Adodc1
    Set Label1(1).DataSource = Adodc1
    
    Text1.DataField = "SteakCode"
    Text2.DataField = "SteakName"
    Text3.DataField = "Price"
    Text4.DataField = "Portion"
    Label1(1).DataField = "StockCode"
    
    ComboLoad
End Sub

Private Sub WineMode()
    SetClear

    Adodc1.RecordSource = "select * from WineT order by WineCode"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh
    Set RS = New Recordset
'    RS.Requery
    
    Label1(2).Caption = "WineCode"
    Label1(4).Caption = "WineName"
    Label1(3).Caption = "Price"
    Label1(0).Caption = "StockCode"
    Label1(5).Caption = "Vintage"
    Label1(1).Caption = "StockCode"
    
    Text1.DataChanged = True
    Set Text1.DataSource = Adodc1
    Set Text2.DataSource = Adodc1
    Set Text3.DataSource = Adodc1
    Set Text4.DataSource = Adodc1
    Set Label1(1).DataSource = Adodc1
  

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Combo1.Text = ""

    Text1.DataField = "WineCode"
    Text2.DataField = "WineName"
    Text3.DataField = "Price"
    Text4.DataField = "Vintage"
    Label1(1).DataField = "StockCode"
    
    ComboLoad
End Sub

Private Sub SetClear()
    
    Set DataGrid1.DataSource = Nothing
    Set RS = Nothing
    Set Text1.DataSource = Nothing
    Set Text2.DataSource = Nothing
    Set Text3.DataSource = Nothing
    Set Text4.DataSource = Nothing
    Set Label1(1).DataSource = Nothing
    Text1.DataField = ""
    Text2.DataField = ""
    Text3.DataField = ""
    Text4.DataField = ""
    Label1(1).DataField = ""
End Sub


Private Sub ComboLoad()
    sql = "select StockName from StockT where StockCode = '" & Label1(1).Caption & "'"
    RS.Open sql, dB, adOpenStatic, adLockOptimistic
    Dim countR As String
    countR = RS.RecordCount
    If countR <> 0 Then
        Combo1.Text = RS.Fields("StockName")
    End If
    RS.Close
End Sub


Private Sub Option1_Click(Index As Integer)
    If Option1(1).Value = True Then
        WineMode
    Else
        SteakMode
    End If
End Sub
