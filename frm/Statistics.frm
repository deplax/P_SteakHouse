VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Statistics 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   0  '없음
   Caption         =   "Statistics"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin prjFrmSkinV8.frmSkinV8 frmSkinV81 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   16960
      Caption         =   "Statistics"
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   6495
         Left            =   720
         OleObjectBlob   =   "Statistics.frx":0000
         TabIndex        =   1
         Top             =   1920
         Width           =   9375
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   8160
         Top             =   8640
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
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
   End
End
Attribute VB_Name = "Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Adodc1.RecordSource = "SELECT ORDERT.ORDERNO, STEAKT.STEAKNAME, STEAKT.PRICE, ORDERT.QUANTITY, STEAKT.PRICE * ORDERT.QUANTITY AS Total , ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) AS ProtionPrice, STEAKT.PRICE * ORDERT.QUANTITY - ROUND(STOCKT.PRICE / STOCKT.QUANTITY * STEAKT.PORTION, 2) * ORDERT.QUANTITY AS NetProfit, ORDERT.ORDERTIME From ORDERT, STOCKT, STEAKT WHERE ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STOCKT.STOCKCODE = STEAKT.STOCKCODE AND TO_CHAR(ORDERT.ORDERTIME,'dd') = '" & Format(Day(Date), "00") & "'"
    Adodc1.Refresh
    
    Set MSChart1.DataSource = Adodc1
    MSChart1.Refresh
End Sub


