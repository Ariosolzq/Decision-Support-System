VERSION 5.00
Begin VB.Form dy 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9285
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "参数a、b1、b2的值"
      Height          =   3495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      Begin VB.CommandButton button1 
         Caption         =   "计算"
         Height          =   735
         Left            =   2880
         TabIndex        =   7
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox Txtb2 
         Enabled         =   0   'False
         Height          =   270
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Txtb1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Txta 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "回归系数b2"
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "回归系数b1"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "常数项a"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "dy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub button1_Click()

    '调用最小二乘法模型程序
    Shell ("C:\hylyc\mx\zxecf\mbwj\zxecf.exe")
    
    Dim strSQL As String
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn
    
    '定义变量a,b1,b2存放从cs表中得到的a,b1,b2字段的值
    Dim a, b1, b2 As Double
    
    '等待最小二乘法模型程序执行完毕
    For i = O To 9999
        strSQL = " select * from cs"
        ADOrs.Open strSQL
        a = ADOrs("a")
        b1 = ADOrs("b1")
        b2 = ADOrs("b2")
        ADOrs.Close
    
        If a <> 0 Then
            Exit For
        End If
    Next i

'定义变量m,n,p分别存放将a,b1,b2转化为字符串后的值
Dim m, n, p As String
m = CStr(a)
n = CStr(b1)
p = CStr(b2)

'若m左侧的第一个字苻为"."，则在笫一个字符前加“0”
If Left(m, 1) = "." Then
    m = "0" & m
End If

'若n左侧的第一个字苻为"."，则在笫一个字符前加“0”
If Left(n, 1) = "." Then
    n = "0" & n
End If


'若p左侧的第一个字苻为"."，则在笫一个字符前加“0”
If Left(p, 1) = "." Then
    p = "0" & p
End If

'在文本框中显示m，n，p的值
Txta.Text = m
Txtb1.Text = n
Txtb2.Text = p
    
'让文本框成为只读
Txta.Enabled = False
Txtb1.Enabled = False
Txtb2.Enabled = False
    
End Sub

Private Sub Form_Load()

    '调用最小二乘法模型程序
    Shell ("C:\hylyc\mx\zxecf\mbwj\zxecf.exe")
    
    Dim strSQL As String
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn
    
    '定义变量a,b1,b2存放从cs表中得到的a,b1,b2字段的值
    Dim a, b1, b2 As Double
    
    '等待最小二乘法模型程序执行完毕
    For i = O To 9999
        strSQL = " select * from cs"
        ADOrs.Open strSQL
        a = ADOrs("a")
        b1 = ADOrs("b1")
        b2 = ADOrs("b2")
        ADOrs.Close
    
        If a <> 0 Then
            Exit For
        End If
        Next i

'定义变量m,n,p分别存放将a,b1,b2转化为字符串后的值
Dim m, n, p As String
m = CStr(a)
n = CStr(b1)
p = CStr(b2)

'若m左侧的第一个字苻为"."，则在笫一个字符前加“0”
If Left(m, 1) = "." Then
    m = "0" & m
End If

'若n左侧的第一个字苻为"."，则在笫一个字符前加“0”
If Left(n, 1) = "." Then
    n = "0" & n
End If


'若p左侧的第一个字苻为"."，则在笫一个字符前加“0”
If Left(p, 1) = "." Then
    p = "0" & p
End If

'在文本框中显示m，n，p的值
Txta.Text = m
Txtb1.Text = n
Txtb2.Text = p
    
'让文本框成为只读
Txta.Enabled = False
Txtb1.Enabled = False
Txtb2.Enabled = False

End Sub


