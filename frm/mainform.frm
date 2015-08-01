VERSION 5.00
Begin VB.Form frmSkin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "Steak House"
   ClientHeight    =   11190
   ClientLeft      =   1050
   ClientTop       =   1245
   ClientWidth     =   9855
   Icon            =   "mainform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   9855
   Begin prjFrmSkinV8.frmSkinV8 UserControl11 
      Height          =   11220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   19791
      Caption         =   "Steak House"
      Begin prjFrmSkinV8.jcbutton CmdMenu 
         Height          =   975
         Left            =   7560
         TabIndex        =   33
         Top             =   7200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Menu"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdSales 
         Height          =   975
         Left            =   7560
         TabIndex        =   32
         Top             =   8400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Sales"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.TextBox TxtMoney 
         Appearance      =   0  '평면
         Height          =   270
         Left            =   7080
         TabIndex        =   30
         Text            =   "TxtMoney"
         Top             =   5280
         Width           =   1695
      End
      Begin prjFrmSkinV8.jcbutton cmdCheck 
         Height          =   855
         Left            =   5880
         TabIndex        =   26
         Top             =   6120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1508
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Check"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.CheckBox reservation 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   3240
         TabIndex        =   21
         Top             =   10680
         Width           =   1335
      End
      Begin VB.CheckBox reservation 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   20
         Top             =   10680
         Width           =   1335
      End
      Begin VB.CheckBox reservation 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   19
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CheckBox reservation 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   18
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CheckBox reservation 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   17
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CheckBox reservation 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   16
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CheckBox reservation 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   15
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CheckBox reservation 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Reservation"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   14
         Top             =   3120
         Width           =   1695
      End
      Begin prjFrmSkinV8.jcbutton CmdClose 
         Height          =   975
         Left            =   7560
         TabIndex        =   13
         Top             =   9600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Close"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdStatistics 
         Height          =   975
         Left            =   5880
         TabIndex        =   12
         Top             =   9600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Statistics"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdStock 
         Height          =   975
         Left            =   5880
         TabIndex        =   11
         Top             =   8400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Stock"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdOrder 
         Height          =   975
         Left            =   5880
         TabIndex        =   10
         Top             =   7200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Order"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.ListBox List1 
         Appearance      =   0  '평면
         Height          =   3270
         Left            =   5760
         TabIndex        =   9
         Top             =   1080
         Width           =   3375
      End
      Begin prjFrmSkinV8.jcbutton Table 
         Height          =   1935
         Index           =   7
         Left            =   3120
         TabIndex        =   8
         Top             =   8640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Table08"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton Table 
         Height          =   1935
         Index           =   6
         Left            =   600
         TabIndex        =   7
         Top             =   8640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Table07"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton Table 
         Height          =   1935
         Index           =   5
         Left            =   3120
         TabIndex        =   6
         Top             =   6120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Table06"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton Table 
         Height          =   1935
         Index           =   4
         Left            =   600
         TabIndex        =   5
         Top             =   6120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Table05"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton Table 
         Height          =   1935
         Index           =   3
         Left            =   3120
         TabIndex        =   4
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Table04"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton Table 
         Height          =   1935
         Index           =   2
         Left            =   600
         TabIndex        =   3
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Table03"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton Table 
         Height          =   1935
         Index           =   1
         Left            =   3120
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Table02"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton Table 
         Height          =   1935
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3413
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Table01"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label lblItem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No    Item(q)          Ea      Unit      Sum"
         Height          =   255
         Left            =   5760
         TabIndex        =   31
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label ChargeFld 
         BackStyle       =   0  '투명
         Caption         =   "ChargeValue"
         Height          =   255
         Left            =   7080
         TabIndex        =   29
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label lblCharge 
         BackStyle       =   0  '투명
         Caption         =   "Charge"
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label lblMoney 
         BackStyle       =   0  '투명
         Caption         =   "Money"
         Height          =   255
         Left            =   6120
         TabIndex        =   27
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label TableNoFld 
         BackStyle       =   0  '투명
         Caption         =   "TableNoValue"
         Height          =   255
         Left            =   7080
         TabIndex        =   25
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label lblTableNo 
         BackStyle       =   0  '투명
         Caption         =   "TableNo"
         Height          =   255
         Left            =   6120
         TabIndex        =   24
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label TotalFld 
         BackStyle       =   0  '투명
         Caption         =   "TotalValue"
         Height          =   255
         Left            =   7080
         TabIndex        =   23
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '투명
         Caption         =   "Total"
         Height          =   255
         Left            =   6120
         TabIndex        =   22
         Top             =   4920
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   10740
         Left            =   15
         Picture         =   "mainform.frx":87CA
         Top             =   405
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Shadow As clsShadow


Dim dB As Connection
Dim RS As Recordset
Dim RS3 As Recordset
Dim currentR As Long
Dim totalR As Long

Private Sub cmdCheck_Click()
    If TableNoFld.Caption <> "None" Then
        Dim sql As String
        Dim sql2 As String
        Dim sql3 As String
        Dim sqlex As String
        Dim sqlex2 As String
        Dim CountR3 As Integer
        Dim i As Integer
        sql = "select * from TableT where TableNo = " & TableNoFld.Caption & ""
        RS.Open sql, dB, adOpenKeyset, adLockPessimistic
        
        If RS.Fields("State") = 0 Then
            sqlex = "update TableT set State = 1 where TableNo = " & TableNoFld.Caption & ""
        Else
        '손님이 나갈때
            sql2 = "select OrderListT.ORDERLISTNO from ORDERLISTT, ORDERT WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND OrderT.WINECODE IS NULL AND ORDERLISTT.TableNo = " & TableNoFld.Caption & ""
            RS3.Open sql2, dB, adOpenKeyset, adLockPessimistic
        '스테이크의 경우
            CountR3 = RS3.RecordCount
            If CountR3 <> 0 Then
            '시켰다면
                RS3.MoveFirst
                For i = 1 To CountR3
                    sql3 = "Update STOCKT SET QUANTITY = (SELECT DISTINCT stockt.QUANTITY - steakt.PORTION*(SELECT ORDERT.QUANTITY FROM ORDERLISTT, ORDERT WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERLISTT.ORDERLISTNO = " & RS3.Fields("OrderListNo") & ") AS T From ORDERLISTT, ORDERT, STOCKT, STEAKT, WINET WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STEAKT.STOCKCODE = STOCKT.STOCKCODE AND ORDERLISTT.ORDERLISTNO = " & RS3.Fields("OrderListNo") & ") WHERE stockcode = (SELECT DISTINCT stockT.STOCKCODE From ORDERLISTT, ORDERT, STOCKT, STEAKT, WINET WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STEAKT.STOCKCODE = STOCKT.STOCKCODE AND ORDERLISTT.ORDERLISTNO = " & RS3.Fields("OrderListNo") & ")"
'                    Update STOCKT SET QUANTITY = (SELECT DISTINCT stockt.QUANTITY - steakt.PORTION*(SELECT ORDERT.QUANTITY FROM ORDERLISTT, ORDERT WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERLISTT.ORDERLISTNO = 13) AS T From ORDERLISTT, ORDERT, STOCKT, STEAKT, WINET WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STEAKT.STOCKCODE = STOCKT.STOCKCODE) WHERE stockcode = (SELECT DISTINCT stockT.STOCKCODE From ORDERLISTT, ORDERT, STOCKT, STEAKT, WINET WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERT.STEAKCODE = STEAKT.STEAKCODE AND STEAKT.STOCKCODE = STOCKT.STOCKCODE AND ORDERLISTT.ORDERLISTNO = 13)
                    'List2.AddItem RS3.Fields("OrderListNo")
                    dB.Execute sql3
                    RS3.MoveNext
                Next
            End If
            RS3.Close
            
            sql2 = "select OrderListT.ORDERLISTNO from ORDERLISTT, ORDERT WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND OrderT.STEAKCODE IS NULL AND ORDERLISTT.TableNo = " & TableNoFld.Caption & ""
            RS3.Open sql2, dB, adOpenKeyset, adLockPessimistic
        '와인의 경우
            CountR3 = RS3.RecordCount
            If CountR3 <> 0 Then
                RS3.MoveFirst
                For i = 1 To CountR3
                    sql3 = "Update STOCKT SET QUANTITY = QUANTITY - (SELECT ORDERT.QUANTITY FROM ORDERLISTT, ORDERT WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERLISTT.ORDERLISTNO = " & RS3.Fields("OrderListNo") & ") WHERE stockcode = (SELECT DISTINCT stockT.STOCKCODE From ORDERLISTT, ORDERT, STOCKT, STEAKT, WINET WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERT.WINECODE = WINET.WINECODE AND WINET.STOCKCODE = STOCKT.STOCKCODE AND ORDERLISTT.ORDERLISTNO = " & RS3.Fields("OrderListNo") & ")"
                    'Update STOCKT SET QUANTITY = QUANTITY - (SELECT ORDERT.QUANTITY FROM ORDERLISTT, ORDERT WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERLISTT.ORDERLISTNO = 16) WHERE stockcode = (SELECT DISTINCT stockT.STOCKCODE From ORDERLISTT, ORDERT, STOCKT, STEAKT, WINET WHERE ORDERLISTT.ORDERNO = ORDERT.ORDERNO AND ORDERT.WINECODE = WINET.WINECODE AND WINET.STOCKCODE = STOCKT.STOCKCODE AND ORDERLISTT.ORDERLISTNO = 16);
                    'List2.AddItem RS3.Fields("OrderListNo")
                    dB.Execute sql3
                    RS3.MoveNext
                Next
            End If
            RS3.Close
            
            sqlex = "update TableT set State = 0 where TableNo = " & TableNoFld.Caption & ""
            sqlex2 = "delete from OrderListT where TableNo = " & TableNoFld.Caption & ""
            dB.Execute sqlex2
        End If
        dB.Execute sqlex
        
        RS.Close
        TableSet
    End If
    
    ListUpdate (TableNoFld.Caption - 1)
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdMenu_Click()
    Menu.Show
End Sub

Private Sub CmdOrder_Click()
    If TableNoFld = "None" Then
        MsgBox ("주문할 테이블을 선택해 주세요")
    Else
        Order.Show
    End If
End Sub

Private Sub CmdSales_Click()
    Sales.Show
End Sub

Private Sub CmdStatistics_Click()
    Statistics.Show
End Sub

Private Sub CmdStock_Click()
    Stock.Show
End Sub

Private Sub Form_Initialize()
FadeIN Me
End Sub

Private Sub Form_Load()
'//그림자소스
    Set Shadow = New clsShadow
    Call Shadow.Shadow(Me)
    Shadow.Color = vbBlack
    Shadow.Depth = 4
'//그림자소스끝

ApplyIcon hWnd

'===================================================

    
    Set dB = New Connection
    dB.CursorLocation = adUseClient
    dB.Open "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
    Set RS = New Recordset
    Set RS3 = New Recordset
    
    TableSet
    ReserveSet
   
'==================================================
'컨트롤 초기화

    TableNoFld.Caption = "None"
    TotalFld.Caption = ""
    TxtMoney.Text = ""
    ChargeFld.Caption = ""
    
End Sub
'예약 체크용
Private Sub ReservationCheck()
Dim i As Integer
For i = 0 To 7
    If Table(i).ButtonStyle = eVistaAero Then
        reservation(i).Enabled = False
    Else
        reservation(i).Enabled = True
    End If
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
FadeOUT Me
'Cancel = 1
End Sub
Private Sub TableSet()
    Dim sql As String
    sql = "select * from TableT order by TableNo"
    RS.Open sql, dB, adOpenStatic, adLockOptimistic
    RS.MoveFirst
    
'===================================================
'테이블 버튼 초기화
    Dim i As Integer
    For i = 0 To 7
        If RS.Fields("State") = 1 Then
            Table(i).ButtonStyle = eVistaAero
            Table(i).BackColor = &HC0C0FF
        Else
            Table(i).ButtonStyle = eOutlook2007
            Table(i).BackColor = &HFFD1AD
        End If
        RS.MoveNext
    Next
    ReservationCheck
    RS.Close
End Sub


Private Sub reservation_Click(Index As Integer)
    
    Dim sql As String
    
    If reservation(Index).Value = 1 Then
        sql = "update TableT set reserve = 1 where tableno = " & Index & " + 1"
        Table(Index).BackColor = &HE0E0E0
        Table(Index).Enabled = False
    Else
        sql = "update TableT set reserve = 0 where tableno = " & Index & " + 1"
        Table(Index).Enabled = True
        Table(Index).BackColor = &HFFD1AD
    End If
'    RS.Close
'RS.Open sql, dB, adOpenKeyset, adLockPessimistic
    dB.Execute sql
 '   RS.Close
End Sub

Private Sub Table_Click(Index As Integer)
    TableNoFld.Caption = Index + 1
    
    ListUpdate (Index)
    
End Sub


Private Sub ListUpdate(Index As Integer)

    List1.Clear
    Dim sql As String
    sql = "SELECT OrderListT.ORDERLISTNO, SteakT.STEAKNAME, OrderT.QUANTITY, SteakT.PRICE, OrderT.Burn, OrderT.QUANTITY * SteakT.PRICE AS Total From OrderListT, OrderT, SteakT WHERE OrderListT.ORDERNO = OrderT.ORDERNO AND OrderT.STEAKCODE = SteakT.STEAKCODE AND OrderListT.TABLENO = " & Index + 1 & ""
    RS.Open sql, dB, adOpenKeyset, adLockPessimistic
    
    Dim Str As String
 
    totalR = RS.RecordCount
    
    If totalR <> 0 Then
        RS.MoveFirst
    End If
    List1.Clear
    List1.AddItem ("== ===== ===== Steak ===== ===== ==")
    Dim i, Tot
    For i = 0 To totalR - 1
        Str = RS.Fields(0) & "   " & RS.Fields(1) & "(" & RS.Fields(4) & ")   " & RS.Fields(2) & "   " & RS.Fields(3) & "   " & RS.Fields(2) * RS.Fields(3)
        Tot = Tot + RS.Fields(5)
        List1.AddItem (Str)
        RS.MoveNext
    Next
    
    RS.Close
    
    sql = "SELECT OrderListT.ORDERLISTNO, WineT.WINENAME, OrderT.QUANTITY, WineT.PRICE, OrderT.Burn, OrderT.QUANTITY * WineT.PRICE AS Total From OrderListT, OrderT, WineT WHERE OrderListT.ORDERNO = OrderT.ORDERNO AND OrderT.WineCODE = WineT.WineCODE AND OrderListT.TABLENO =  " & Index + 1 & ""
    RS.Open sql, dB, adOpenKeyset, adLockPessimistic
    
    totalR = RS.RecordCount
    If totalR <> 0 Then
        RS.MoveFirst
    End If
    List1.AddItem ("== ===== ===== Wine ===== ===== ==")
    For i = 0 To totalR - 1
        Str = RS.Fields(0) & "   " & RS.Fields(1) & "(" & RS.Fields(4) & ")   " & RS.Fields(2) & "   " & RS.Fields(3) & "   " & RS.Fields(2) * RS.Fields(3)
        Tot = Tot + RS.Fields(5)
        List1.AddItem (Str)
        RS.MoveNext
    Next
    
    RS.Close
    
    TotalFld = Tot
End Sub

Private Sub ReserveSet()
    Dim sql As String
    sql = "select * from TableT order by TableNo"
    RS.Open sql, dB, adOpenStatic, adLockOptimistic
    RS.MoveFirst
    
'===================================================
'예약 버튼 초기화
    Dim i As Integer
    For i = 0 To 7
        If RS.Fields("Reserve") = 1 Then
            reservation(i).Value = 1
        Else
            reservation(i).Value = 0
        End If
        RS.MoveNext
    Next
    ReservationCheck
    RS.Close
End Sub



Public Sub ListReLoad(TB As Integer)
    ListUpdate (TB)
End Sub

Private Sub TxtMoney_Change()
    ChargeFld.Caption = Val(TxtMoney.Text) - Val(TotalFld.Caption)
End Sub
