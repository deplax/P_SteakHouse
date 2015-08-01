VERSION 5.00
Begin VB.Form Order 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   11265
   ClientTop       =   1455
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin prjFrmSkinV8.frmSkinV8 frmSkinV81 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11245
      Caption         =   "Order"
      Begin prjFrmSkinV8.jcbutton CmdDel 
         Height          =   975
         Left            =   3480
         TabIndex        =   28
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "Delete"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.ListBox List1 
         Appearance      =   0  '평면
         Height          =   3450
         Left            =   4080
         TabIndex        =   27
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   1800
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
         Begin VB.OptionButton OpBurn 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "Rare"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton OpBurn 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "Medium"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   25
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OpBurn 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            Caption         =   "Welldone"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   24
            Top             =   480
            Width           =   1095
         End
      End
      Begin prjFrmSkinV8.jcbutton CmdClear 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   4320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "Clear"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   9
         Left            =   3360
         TabIndex        =   17
         Top             =   3960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "0"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   8
         Left            =   3000
         TabIndex        =   16
         Top             =   3960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "9"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   15
         Top             =   3960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "8"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   14
         Top             =   3960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "7"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   13
         Top             =   3960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "6"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   4
         Left            =   3360
         TabIndex        =   12
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "5"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "4"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   10
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "3"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "2"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "Wine"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   2880
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin prjFrmSkinV8.jcbutton CmdNumber 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   6
         Top             =   3600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "1"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.TextBox TxtQuantity 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   5
         Text            =   "Quantity"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.ComboBox CmbWine 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   1800
         TabIndex        =   4
         Text            =   "Null"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox CmbSteak 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   1800
         TabIndex        =   3
         Text            =   "Null"
         Top             =   1680
         Width           =   1935
      End
      Begin prjFrmSkinV8.jcbutton CmdAdd 
         Height          =   975
         Left            =   1680
         TabIndex        =   2
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "Add"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin prjFrmSkinV8.jcbutton CmdClose 
         Height          =   975
         Left            =   5280
         TabIndex        =   1
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1720
         ButtonStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Caption         =   "Close"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin VB.Label TableNoFld 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "TableNoValue"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "LIST"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "ITEM"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblTableNo 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         Caption         =   "TableNo"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   5595
         Left            =   2400
         Picture         =   "Order.frx":0000
         Top             =   720
         Width           =   5685
      End
   End
End
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dB As Connection
Dim RS As Recordset
Dim RS2 As Recordset
Dim currentR As Long
Dim totalR As Long
Dim Str As String

Private Sub CmdAdd_Click()
    
    If TxtQuantity.Text = "Quantity" Or Str = "" Then
        MsgBox "수량을 입력해주세요^^"
    Else
        Dim sql1 As String
        Dim sql2 As String
        Dim sqlCode As String
        
        Dim BurnSwitch As Integer
        
        Set RS2 = New Recordset
        
        If OpBurn(0).Value = True Then
            BurnSwitch = 0
        ElseIf OpBurn(1).Value = True Then
            BurnSwitch = 1
        Else
            BurnSwitch = 2
        End If
            
        
        
        If Option1(0).Value = True Then
            '스테이크 옵션이 선택되어 있으면
            sqlCode = "select SteakCode from SteakT where SteakName = '" & CmbSteak.Text & "'"
            RS2.Open sqlCode, dB, adOpenStatic, adLockOptimistic
            sql1 = "insert into OrderT(OrderNo, OrderTable, SteakCode, Quantity, Burn) values(Order_no.nextval, " & TableNoFld.Caption & ", '" & RS2.Fields("SteakCode") & "', '" & TxtQuantity.Text & "', '" & OpBurn(1).Caption & "')"
            RS2.Close
            
        Else
            sqlCode = "select WineCode from WineT where WineName = '" & CmbWine.Text & "'"
            RS2.Open sqlCode, dB, adOpenStatic, adLockOptimistic
            sql1 = "insert into OrderT(OrderNo, OrderTable, WineCode, Quantity) values(Order_no.nextval, " & TableNoFld.Caption & ", '" & RS2.Fields("WineCode") & "', '" & TxtQuantity.Text & "')"
            RS2.Close
        End If
        sql2 = "insert into OrderListT values(OrderList_No.NextVal, Order_No.CurrVal, " & TableNoFld.Caption & ")"
        dB.Execute sql1
        dB.Execute sql2
        
        frmSkin.ListReLoad (TableNoFld.Caption - 1)
        ListLoad
        Str = ""
        TxtQuantity.Text = "Quantity"
    End If
End Sub

Private Sub CmdClear_Click()
    TxtQuantity.Text = "Quantity"
    Str = ""
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdDel_Click()
    
    Dim Nostr As String
    Nostr = Trim(Left(List1.Text, 4))
    
    Dim sql1 As String
    Dim sql2 As String
    Dim sqlCode As String

    Set RS2 = New Recordset
    
    sql1 = "delete from OrderListT where OrderListNo = '" & Nostr & "'"
    sql2 = "delete from OrderT where OrderNo in (select OrderNo from OrderListT where OrderListNo = '" & Nostr & "')"

    dB.Execute sql2
    dB.Execute sql1
    
    frmSkin.ListReLoad (TableNoFld.Caption - 1)
    ListLoad
End Sub

Private Sub CmdNumber_Click(Index As Integer)
    Str = Str + CmdNumber(Index).Caption
    TxtQuantity.Text = Str
End Sub


Private Sub Form_Load()
    Image1.Width = 4780

    Option1(0).Value = True
    TableNoFld.Caption = frmSkin.TableNoFld.Caption
    
    Dim i As Integer
    
    Dim sql As String
    
    Set dB = New Connection
    dB.CursorLocation = adUseClient
    dB.Open "PROVIDER=MSDASQL;dsn=pro304;uid=pro304;pwd=qlalf;database=ora_na;"
    Set RS = New Recordset
    sql = "select SteakName from SteakT order by SteakCode"
    RS.Open sql, dB, adOpenStatic, adLockOptimistic
    
    totalR = RS.RecordCount

    If totalR <> 0 Then
        RS.MoveFirst
    End If
    With CmbSteak
        For i = 1 To totalR
            .AddItem RS.Fields("SteakName")
            RS.MoveNext
        Next
    End With
    
    RS.Close
    
    sql = "select WineName from WineT order by WineCode"
    RS.Open sql, dB, adOpenStatic, adLockOptimistic
    
    totalR = RS.RecordCount

    If totalR <> 0 Then
        RS.MoveFirst
    End If
    With CmbWine
        For i = 1 To totalR
            .AddItem RS.Fields("WineName")
            RS.MoveNext
        Next
    End With
    
    CmbSteak.ListIndex = 0
    CmbWine.ListIndex = 0
    RS.Close
    
    OpBurn(1).Value = True
    ListLoad
End Sub
Private Sub ListLoad()
    List1.Clear
    Dim ListCnt As Integer
    Dim i As Integer
    ListCnt = frmSkin.List1.ListCount
    For i = 1 To ListCnt
        frmSkin.List1.ListIndex = i - 1
        List1.AddItem frmSkin.List1.Text
    Next
End Sub


Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        CmbWine.Enabled = False
        CmbSteak.Enabled = True
        OpBurn(0).Enabled = True
        OpBurn(1).Enabled = True
        OpBurn(2).Enabled = True
    Else
        CmbSteak.Enabled = False
        CmbWine.Enabled = True
        OpBurn(0).Enabled = False
        OpBurn(1).Enabled = False
        OpBurn(2).Enabled = False
    End If
End Sub

