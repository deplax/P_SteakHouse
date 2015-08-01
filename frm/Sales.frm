VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Sales 
   BackColor       =   &H80000009&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   5640
   ClientTop       =   2895
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin prjFrmSkinV8.frmSkinV8 frmSkinV81 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _ExtentX        =   16536
      _ExtentY        =   13573
      Caption         =   "Sales"
      Begin VB.TextBox Text7 
         Appearance      =   0  '평면
         Height          =   375
         Left            =   11760
         TabIndex        =   7
         Text            =   "Text6"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  '평면
         Height          =   375
         Left            =   9480
         TabIndex        =   6
         Text            =   "Text6"
         Top             =   1440
         Width           =   1575
      End
      Begin prjFrmSkinV8.jcbutton CmdSerch2 
         Height          =   495
         Left            =   13560
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Caption         =   "Search(~)"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton CmdSearch 
         Height          =   495
         Left            =   13560
         TabIndex        =   5
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Caption         =   "Search"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   13560
         Top             =   6600
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Connect         =   "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
         OLEDBString     =   "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc2"
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4335
         Left            =   7560
         TabIndex        =   10
         Top             =   2640
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7646
         _Version        =   393216
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
      Begin prjFrmSkinV8.jcbutton CmdClose 
         Height          =   495
         Left            =   13560
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   6240
         Top             =   6600
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Connect         =   "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
         OLEDBString     =   "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
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
      Begin VB.TextBox Text3 
         Appearance      =   0  '평면
         Height          =   375
         Left            =   12120
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '평면
         Height          =   375
         Left            =   10800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         Height          =   375
         Left            =   9480
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4335
         Left            =   240
         TabIndex        =   1
         Top             =   2640
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7646
         _Version        =   393216
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
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "WineSalse"
         Height          =   255
         Index           =   3
         Left            =   7560
         TabIndex        =   34
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "SteakSalse"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "(yyyymmdd)"
         Height          =   255
         Index           =   1
         Left            =   11760
         TabIndex        =   32
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "(yyyymmdd)"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   31
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   255
         Left            =   11160
         TabIndex        =   30
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Day(dd)"
         Height          =   255
         Index           =   2
         Left            =   12120
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Month(mm)"
         Height          =   255
         Index           =   1
         Left            =   10800
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "Year(yyyy)"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblTNP 
         BackStyle       =   0  '투명
         Caption         =   "lblTNP"
         Height          =   255
         Left            =   9600
         TabIndex        =   26
         Top             =   7680
         Width           =   975
      End
      Begin VB.Label lblTS 
         BackStyle       =   0  '투명
         Caption         =   "lblTS"
         Height          =   255
         Left            =   5520
         TabIndex        =   25
         Top             =   7680
         Width           =   975
      End
      Begin VB.Label lblNPW 
         BackStyle       =   0  '투명
         Caption         =   "lblNPW"
         Height          =   255
         Left            =   12960
         TabIndex        =   24
         Top             =   7080
         Width           =   975
      End
      Begin VB.Label lblPCW 
         BackStyle       =   0  '투명
         Caption         =   "lblPCW"
         Height          =   255
         Left            =   10920
         TabIndex        =   23
         Top             =   7080
         Width           =   975
      End
      Begin VB.Label lblGSW 
         BackStyle       =   0  '투명
         Caption         =   "lblGSW"
         Height          =   255
         Left            =   8760
         TabIndex        =   22
         Top             =   7080
         Width           =   975
      End
      Begin VB.Label lblNPS 
         BackStyle       =   0  '투명
         Caption         =   "lblNPS"
         Height          =   255
         Left            =   5760
         TabIndex        =   21
         Top             =   7080
         Width           =   975
      End
      Begin VB.Label lblPCS 
         BackStyle       =   0  '투명
         Caption         =   "lblPCS"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   7080
         Width           =   975
      End
      Begin VB.Label lblGSS 
         BackStyle       =   0  '투명
         Caption         =   "lblGSS"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   7080
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Total Net Profit"
         Height          =   255
         Index           =   7
         Left            =   8040
         TabIndex        =   18
         Top             =   7680
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Total Sales"
         Height          =   255
         Index           =   6
         Left            =   4320
         TabIndex        =   17
         Top             =   7680
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Net Profit"
         Height          =   255
         Index           =   5
         Left            =   12000
         TabIndex        =   16
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Net Profit"
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   15
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Production Cost"
         Height          =   375
         Index           =   3
         Left            =   9840
         TabIndex        =   14
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Production Cost"
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   13
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Gross Salse"
         Height          =   255
         Index           =   1
         Left            =   7560
         TabIndex        =   12
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Gross Salse"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Image Image1 
         Appearance      =   0  '평면
         Height          =   1560
         Left            =   1680
         Picture         =   "Sales.frx":0000
         Top             =   720
         Width           =   5640
      End
   End
End
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrAll As String
Dim Stry As String
Dim Strm As String
Dim Strd As String

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSearch_Click()

    Dim tagstr As String
    Stry = Format(Text1.Text, "0000")
    Strm = Format(Text2.Text, "00")
    Strd = Format(Text3.Text, "00")
    StrAll = Stry & Strm & Strd

    
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
    MsgBox ("검색할 날짜를 입력해주세요^^")
ElseIf Text1.Text = "" And Text2.Text = "" Then
'dd
    tagstr = "TO_CHAR(ORDERT.ORDERTIME,'dd') = '" & StrAll & "'"
ElseIf Text2.Text = "" And Text3.Text = "" Then
'yyyy
    tagstr = "TO_CHAR(ORDERT.ORDERTIME,'yyyy') = '" & StrAll & "'"
ElseIf Text1.Text = "" And Text3.Text = "" Then
'mm
    tagstr = "TO_CHAR(ORDERT.ORDERTIME,'mm') = '" & StrAll & "'"
ElseIf Text3.Text = "" Then
'yyyymm
    tagstr = "TO_CHAR(ORDERT.ORDERTIME,'yyyymm') = '" & StrAll & "'"
ElseIf Text1.Text = "" Then
'mmdd
    tagstr = "TO_CHAR(ORDERT.ORDERTIME,'mmdd') = '" & StrAll & "'"
ElseIf Text2.Text = "" Then
    MsgBox ("년과 일만 입력할 수 없습니다.")
Else
    tagstr = "TO_CHAR(ORDERT.ORDERTIME,'yyyymmdd') = '" & StrAll & "'"
End If

    Adodc1.RecordSource = "SELECT ORDERT.ORDERNO, STEAKT.STEAKNAME, STEAKT.PRICE, ORDERT.QUANTITY, STEAKT.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) AS ProtionPrice, STEAKT.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) * ORDERT.QUANTITY AS NetProfit, ORDERT.ORDERTIME From ORDERT, STOCKT, STEAKT WHERE ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STOCKT.STOCKCODE = STEAKT.STOCKCODE AND " & tagstr & ""
    Adodc1.Refresh
    
    
    'SELECT ORDERT.ORDERNO, WINET.WINENAME, WINET.PRICE, ORDERT.QUANTITY, WINET.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY, 2) AS ProtionPrice, WINET.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY,  2) * ORDERT.QUANTITY AS NetProfit From ORDERT, STOCKT, WINET WHERE ORDERT.WINECODE = WINET.WINECODE AND STOCKT.STOCKCODE = WINET.STOCKCODE AND TO_CHAR(ORDERT.ORDERTIME,'dd') = '09';
    
    
    Adodc2.RecordSource = "SELECT ORDERT.ORDERNO, WINET.WINENAME, WINET.PRICE, ORDERT.QUANTITY, WINET.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY, 2) AS ProtionPrice, WINET.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY,  2) * ORDERT.QUANTITY AS NetProfit From ORDERT, STOCKT, WINET WHERE ORDERT.WINECODE = WINET.WINECODE AND STOCKT.STOCKCODE = WINET.STOCKCODE AND " & tagstr & ""
    Adodc2.Refresh

    Set DataGrid2.DataSource = Adodc2
    DataGrid2.Refresh
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh

    CalcT

End Sub

Private Sub CmdSerch2_Click()

    Dim tagstr As String


    Adodc1.RecordSource = "SELECT ORDERT.ORDERNO, STEAKT.STEAKNAME, STEAKT.PRICE, ORDERT.QUANTITY, STEAKT.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) AS ProtionPrice, STEAKT.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) * ORDERT.QUANTITY AS NetProfit, ORDERT.ORDERTIME From ORDERT, STOCKT, STEAKT WHERE ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STOCKT.STOCKCODE = STEAKT.STOCKCODE AND ORDERT.ORDERTIME >= TO_DATE('" & Text6.Text & "','yyyymmdd') AND ORDERT.ORDERTIME <= TO_DATE('" & Text7.Text & "','yyyymmdd')"
    Adodc1.Refresh
    
    
    'SELECT ORDERT.ORDERNO, WINET.WINENAME, WINET.PRICE, ORDERT.QUANTITY, WINET.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY, 2) AS ProtionPrice, WINET.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY,  2) * ORDERT.QUANTITY AS NetProfit From ORDERT, STOCKT, WINET WHERE ORDERT.WINECODE = WINET.WINECODE AND STOCKT.STOCKCODE = WINET.STOCKCODE AND TO_CHAR(ORDERT.ORDERTIME,'dd') = '09';
    
    
    Adodc2.RecordSource = "SELECT ORDERT.ORDERNO, WINET.WINENAME, WINET.PRICE, ORDERT.QUANTITY, WINET.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY, 2) AS ProtionPrice, WINET.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY,  2) * ORDERT.QUANTITY AS NetProfit From ORDERT, STOCKT, WINET WHERE ORDERT.WINECODE = WINET.WINECODE AND STOCKT.STOCKCODE = WINET.STOCKCODE AND ORDERT.ORDERTIME >= TO_DATE('" & Text6.Text & "','yyyymmdd') AND ORDERT.ORDERTIME <= TO_DATE('" & Text7.Text & "','yyyymmdd')"
    Adodc2.Refresh

    Set DataGrid2.DataSource = Adodc2
    DataGrid2.Refresh
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh

    CalcT

End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""

    Text6.Text = ""
    Text7.Text = ""

Dim tagstr As String

'AND TO_CHAR(ORDERT.ORDERTIME,'yy') = '" & StrAll & "'

'만약 y가 0이 아니라면 m이 0이 아니고 d도 0이 아니라면
'yyyymmdd
'yyyy, mm, dd, yyyymm, yyyymmdd, mmdd,

'AND TO_CHAR(ORDERT.ORDERTIME,'yyyy') = '" & Format(text1.text, "0000") & "'

    Dim StrDay As String

    StrDay = Day(Date)
    'SELECT ORDERT.ORDERNO, STEAKT.STEAKNAME, STEAKT.PRICE, ORDERT.QUANTITY, STEAKT.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) AS ProtionPrice, STEAKT.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) * ORDERT.QUANTITY AS NetProfit, ORDERT.ORDERTIME From ORDERT, STOCKT, STEAKT WHERE ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STOCKT.STOCKCODE = STEAKT.STOCKCODE AND TO_CHAR(ORDERT.ORDERTIME,'dd') = '09';
    Adodc1.RecordSource = "SELECT ORDERT.ORDERNO, STEAKT.STEAKNAME, STEAKT.PRICE, ORDERT.QUANTITY, STEAKT.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) AS ProtionPrice, STEAKT.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) * ORDERT.QUANTITY AS NetProfit, ORDERT.ORDERTIME From ORDERT, STOCKT, STEAKT WHERE ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STOCKT.STOCKCODE = STEAKT.STOCKCODE AND TO_CHAR(ORDERT.ORDERTIME,'dd') = '" & Format(Day(Date), "00") & "'"
    Adodc1.Refresh
    
    
    'SELECT ORDERT.ORDERNO, WINET.WINENAME, WINET.PRICE, ORDERT.QUANTITY, WINET.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY, 2) AS ProtionPrice, WINET.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY,  2) * ORDERT.QUANTITY AS NetProfit From ORDERT, STOCKT, WINET WHERE ORDERT.WINECODE = WINET.WINECODE AND STOCKT.STOCKCODE = WINET.STOCKCODE AND TO_CHAR(ORDERT.ORDERTIME,'dd') = '09';
    
    
    Adodc2.RecordSource = "SELECT ORDERT.ORDERNO, WINET.WINENAME, WINET.PRICE, ORDERT.QUANTITY, WINET.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY, 2) AS ProtionPrice, WINET.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY,  2) * ORDERT.QUANTITY AS NetProfit From ORDERT, STOCKT, WINET WHERE ORDERT.WINECODE = WINET.WINECODE AND STOCKT.STOCKCODE = WINET.STOCKCODE AND TO_CHAR(ORDERT.ORDERTIME,'dd') = '" & Format(Day(Date), "00") & "'"
    Adodc2.Refresh
    
    
    'Text5.Text = Adodc1.Recordset.Fields(3)
    
    Set DataGrid2.DataSource = Adodc2
    DataGrid2.Refresh
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh
    
    CalcT
    
End Sub

Private Sub CalcT()
    Dim CountR1 As Integer
    Dim CountR2 As Integer
    Dim i As Integer
    CountR1 = Adodc1.Recordset.RecordCount
    CountR2 = Adodc2.Recordset.RecordCount
    
    lblGSS.Caption = "0"
    lblPCS.Caption = "0"
    lblNPS.Caption = "0"
    
    lblGSW.Caption = "0"
    lblPCW.Caption = "0"
    lblNPW.Caption = "0"
    
    lblTS.Caption = "0"
    lblTNP.Caption = "0"
    
    If CountR1 <> 0 Then
        Adodc1.Recordset.MoveFirst
        For i = 1 To CountR1
            lblGSS.Caption = lblGSS.Caption + Adodc1.Recordset.Fields(4)
            lblPCS.Caption = lblPCS.Caption + Adodc1.Recordset.Fields(5)
            lblNPS.Caption = lblNPS.Caption + Adodc1.Recordset.Fields(6)
            Adodc1.Recordset.MoveNext
        Next
    End If
    
    If CountR2 <> 0 Then
        Adodc2.Recordset.MoveFirst
        For i = 1 To CountR2
            lblGSW.Caption = lblGSW.Caption + Adodc2.Recordset.Fields(4)
            lblPCW.Caption = lblPCW.Caption + Adodc2.Recordset.Fields(5)
            lblNPW.Caption = lblNPW.Caption + Adodc2.Recordset.Fields(6)
            Adodc2.Recordset.MoveNext
        Next
    End If
    
    lblTS.Caption = Format(Val(lblGSS.Caption) + Val(lblGSW.Caption), "##,##0.00")
    lblTNP.Caption = Format(Val(lblNPS.Caption) + Val(lblNPW.Caption), "##,##0.00")
End Sub

