VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Stock 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   0  '없음
   Caption         =   "Stock"
   ClientHeight    =   8535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin prjFrmSkinV8.frmSkinV8 frmSkinV81 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   15055
      Caption         =   "Stock"
      Begin prjFrmSkinV8.jcbutton CmdDel 
         Height          =   375
         Left            =   9360
         TabIndex        =   21
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         Caption         =   "Delete"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton CmdSave 
         Height          =   375
         Left            =   8040
         TabIndex        =   20
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
         Height          =   375
         Left            =   6720
         TabIndex        =   19
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
      Begin prjFrmSkinV8.jcbutton jcbutton1 
         Height          =   375
         Left            =   9120
         TabIndex        =   18
         Top             =   7920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin MSACAL.Calendar Calendar2 
         DataField       =   "PURCHASED"
         DataSource      =   "Adodc1"
         Height          =   1935
         Left            =   7560
         TabIndex        =   15
         Top             =   960
         Width           =   3015
         _Version        =   524288
         _ExtentX        =   5318
         _ExtentY        =   3413
         _StockProps     =   1
         BackColor       =   16777215
         Year            =   2011
         Month           =   6
         Day             =   6
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   0
         GridFontColor   =   4210752
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   4210752
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar Calendar1 
         DataField       =   "RENOUNCED"
         DataSource      =   "Adodc1"
         Height          =   1935
         Left            =   3960
         TabIndex        =   14
         Top             =   960
         Width           =   3015
         _Version        =   524288
         _ExtentX        =   5318
         _ExtentY        =   3413
         _StockProps     =   1
         BackColor       =   16777215
         Year            =   2011
         Month           =   6
         Day             =   6
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   0
         GridFontColor   =   4210752
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   4210752
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         DataField       =   "NOTE"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3000
         Width           =   8655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         DataField       =   "COUNTRY"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   4
         Left            =   1920
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         DataField       =   "QUANTITY"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         DataField       =   "PRICE"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         DataField       =   "STOCKNAME"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '평면
         DataField       =   "STOCKCODE"
         DataSource      =   "Adodc1"
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   720
         Top             =   7920
         Width           =   3135
         _ExtentX        =   5530
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
         Appearance      =   0
         BackColor       =   -2147483643
         ForeColor       =   4210752
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
         OLEDBString     =   "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from StockT "
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Stock.frx":0000
         Height          =   3855
         Left            =   720
         TabIndex        =   1
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6800
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Renounce Day"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   7560
         TabIndex        =   17
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Purchase Day"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   16
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Note"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   840
         TabIndex        =   13
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Country"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   840
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Quantity"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   840
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Price"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "StockName"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "StockCode"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   6555
         Left            =   30
         Picture         =   "Stock.frx":0015
         Top             =   1920
         Width           =   5190
      End
   End
End
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Adodc1.Caption = Adodc1.Recordset.AbsolutePosition & "/" & Adodc1.Recordset.RecordCount
End Sub


Private Sub CmdAdd_Click()
    Adodc1.Recordset.AddNew
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

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    Dim grSQL As String
    
    Adodc1.Recordset.Close
    Select Case ColIndex
        Case 0
            grSQL = "select * from StockT order by StockCode "
        Case 1
            grSQL = "select * from StockT order by price "
        Case 2
            grSQL = "select * from StockT order by quantity "
        Case 3
            grSQL = "select * from StockT order by purchaseD "
        Case 4
            grSQL = "select * from StockT order by renounceD "
        Case 5
            grSQL = "select * from StockT order by country "
        Case 6
            grSQL = "select * from StockT order by note "
        Case 7
            grSQL = "select * from StockT order by StockName "
    End Select

    Adodc1.RecordSource = grSQL
    Adodc1.Refresh
End Sub

Private Sub Form_Initialize()
    FadeIN Me
End Sub

Private Sub Form_Load()
    Adodc1.Refresh
End Sub

Private Sub jcbutton1_Click()
    Unload Me
End Sub
