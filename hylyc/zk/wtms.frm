VERSION 5.00
Begin VB.Form wtms 
   Caption         =   "��������"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   8295
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "������Ȼ�����Ԥ�����֧��ϵͳ"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   $"wtms.frx":0000
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "wtms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'��ʾ�ܿش���
zk.Show

'�ر�������������
Unload Me

End Sub
