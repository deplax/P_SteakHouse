VERSION 5.00
Begin VB.Form frmSkin 
   BorderStyle     =   0  '����
   Caption         =   "frmSkin v7 Demo"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   Icon            =   "frmSkin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows �⺻��
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
'�ֱ� �Ѵ� �� ���� ��� ����Ų�Դϴ�.
'���� ���ο�� ���� �̰� �����ؿ� ����
'�������� ���� �Ѱ� �ƴϰ� �����ٴϴ� �׸� ������ �貸�� ���� �����߾��..����
'��, �̹��� Label��Ʈ��, �̹��� ��Ʈ�� ����Ų�� �������ֽ��ϴ�.
'����Ͻð� ���α׷� �����ϽǶ�  'J��Ŵ(J-NaKiM)�� ����Ų' ��� ������ֽø� �����ϰڽ��ϴ�.
'���� ���ϼŵ��ǰ��
'�ϴ�, ����Ų �ҽ��ȿ� ���Ե� ���ϵ���
'���̵��� ���̵�ƿ� �ҽ� (���� ų�� ����->������->������) �ҽ�
'�׸��� �ҽ�
'�������δ��� 32x32 �������� �����Ű�¼ҽ�
'LunaButton - �ѿ��� ��ư ��Ʈ���ε� ���� �ѿ��� ��ư�� ����Ų�ȿ� ���� ���������߾��µ� �Ф�
'                     ȥ�� �غ������ϴϱ� �ȵǼ� �˻��ؼ� �ٸ�������� ���� �ѿ��� �ҽ��� ����غ�����
'                     ���� ��Ʈ�ѿ��� �̻��ϰ� �۵��ؼ� --;;;;
'                     �������δ��̳� �����̴� ���� ���� �̹����� ����ϴ°Ծƴ϶� ��� ��ư�̹�����
'                     ���Ʒ��γ����鼭 �ϴ� ������� �ϰ������
'                     ���� ���� �ҽ��� ���ϱⰡ �������... �����̴� �ҽ��� ���� �������...!!
'                     ����� �����Ұ�!!! �׳� �糪��ư ����߽��ϴ�. ����
'frmSkinV8 ����Ų �ٽ� ���� ��Ʈ���Դϴ�.
'����Ų�� ����Ͻô� �����, �� ������Ʈ���Ͽ��� ������Ʈ�� �����Ͻô� ����� �ֱ�
'�ƴϸ� ���� ������Ʈ�� �̸��� 'prjFrmSkinV8' �� �����Ͻŵ�
'lunabutton ��Ʈ���� �ҷ����ŵ�,
'frmSkinv8 ��Ʈ���� �ҷ�������
'������ ���/Ŭ���� �߰��Ͻø�˴ϴ�.
'����Ų�� �����Ӱ� ����ϼŵ��˴ϴ�!
'���� ������ �ٲٴ� ����� ������Ʈ���� �Ӽ� Caption�� �ٲٽñ���
'���� ĸ�ǵ� �ٲ��ּ���.
'http://jnakim.com/


Dim Shadow As clsShadow

Private Sub Form_Initialize()
FadeIN Me
End Sub

Private Sub Form_Load()
'//�׸��ڼҽ�
    Set Shadow = New clsShadow
    Call Shadow.Shadow(Me)
    Shadow.Color = vbBlack
    Shadow.Depth = 4
'//�׸��ڼҽ���

'//�������δ��� ������ ����ҽ�
ApplyIcon hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
FadeOUT Me
'Cancel = 1
End Sub

Private Sub jcbutton1_Click()
    MsgBox "JN frmSkin V8"
End Sub
