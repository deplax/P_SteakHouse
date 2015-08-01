VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   0  '없음
   Caption         =   "frmSkin v7 Demo"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   Icon            =   "frmSkin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows 기본값
   Begin prjFrmSkinV8.jcbutton jcbutton1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   2175
      _ExtentX        =   3836
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
      BackColor       =   14935011
      Caption         =   $"frmSkin.frx":87CA
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
      ColorScheme     =   2
   End
   Begin prjFrmSkinV8.frmSkinV8 UserControl11 
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3069
      Caption         =   "frmSkinV8"
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'최근 한달 반 동안 썼던 폼스킨입니다.
'어제 새로운거 만들어서 이거 배포해용 ㅇㅇ
'디자인은 제가 한건 아니고 굴러다니는 테마 예뻐서 배껴서 조금 수정했어용..ㅋㅋ
'참, 이번엔 Label컨트롤, 이미지 컨트롤 폼스킨에 넣을수있습니다.
'사용하시고 프로그램 배포하실때  'J나킴(J-NaKiM)의 폼스킨' 요거 사용해주시면 감사하겠습니다.
'물론 안하셔도되고요
'일단, 폼스킨 소스안에 포함된 파일들은
'페이드인 페이드아웃 소스 (껏다 킬때 투명->반투명->불투명) 소스
'그림자 소스
'엠피제로님의 32x32 아이콘을 적용시키는소스
'LunaButton - 롤오버 버튼 컨트롤인데 제가 롤오버 버튼을 폼스킨안에 직접 넣을려고했었는데 ㅠㅠ
'                     혼자 해보려고하니깐 안되서 검색해서 다른사람들이 만든 롤오버 소스를 사용해봤지만
'                     유저 컨트롤에선 이상하게 작동해서 --;;;;
'                     엠피제로님이나 상폭이님 보면 여러 이미지를 사용하는게아니라 모든 버튼이미지를
'                     위아래로내리면서 하는 방법으로 하고싶지만
'                     좋은 예제 소스를 구하기가 힘들었고... 상폭이님 소스는 보기 힘들었고...!!
'                     결론은 귀찮았고!!! 그냥 루나버튼 사용했습니다. ㅇㅇ
'frmSkinV8 폼스킨 핵심 유저 컨트롤입니다.
'폼스킨을 사용하시는 방법은, 이 프로젝트파일에서 프로젝트를 시작하시는 방법도 있구
'아니면 기존 프로젝트의 이름을 'prjFrmSkinV8' 로 설정하신뒤
'lunabutton 컨트롤을 불러오신뒤,
'frmSkinv8 컨트롤을 불러오신후
'나머지 모듈/클래스 추가하시면됩니다.
'폼스킨은 자유롭게 사용하셔도됩니다!
'폼의 제목을 바꾸는 방법은 유저컨트롤의 속성 Caption을 바꾸시구요
'폼의 캡션도 바꿔주세요.
'http://jnakim.com/


Dim Shadow As clsShadow

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

'//앰피제로님의 아이콘 적용소스
ApplyIcon hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
FadeOUT Me
'Cancel = 1
End Sub

Private Sub jcbutton1_Click()
    MsgBox "JN frmSkin V8"
End Sub
