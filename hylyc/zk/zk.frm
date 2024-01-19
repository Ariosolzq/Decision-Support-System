VERSION 5.00
Begin VB.Form zk 
   Caption         =   "年度货运量预测决策支持系统"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9855
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Txt2 
      Height          =   375
      Left            =   6960
      TabIndex        =   22
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Txt1 
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Cmdtc 
      Caption         =   "退出"
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Cmdjxyc 
      Caption         =   "继续预测"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Cmdyc 
      Caption         =   "预测"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "年度工业总产值、年度基建投资"
      Height          =   1575
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   8655
      Begin VB.TextBox Txtndhyl 
         Enabled         =   0   'False
         Height          =   390
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtndjjtz 
         Height          =   375
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Txtndgyzcz 
         Height          =   375
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "单位：（百万吨）"
         Height          =   375
         Left            =   7080
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "年度货运量"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "单位：（亿元）"
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "单位：（亿元）"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "年度基建投资"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "年度工业总产值"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "参数a、b1、b2的值"
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   8655
      Begin VB.TextBox Txtb2 
         Enabled         =   0   'False
         Height          =   270
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Txtb1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Txta 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "回归系数b2"
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "回归系数b1"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "常数项"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "年度货运量预测决策支持系统"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "zk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdjxyc_Click()

    Txtndgyzcz.Text = ""
    Txtndjjtz.Text = ""
    Txtndhyl.Text = ""
    Txtndgyzcz.SetFocus
    Cmdyc.Enabled = True
    Cmdjxyc.Enabled = False
    

End Sub

Private Sub Cmdtc_Click()
    
    ADOcn.Close
    End
    
End Sub

Private Sub Cmdyc_Click()

'判断输入的年度工业总产值是否为空，若为空则弹出提示信息
If Txtndgyzcz.Text = "" Then
    MsgBox "请输入年度工业总产值！"
    Txtndgyzcz.Text = ""
    Txtndgyzcz.SetFocus
    Exit Sub
End If

'判断输入的年度工业总产值是否为数字，若不为数字则弹出提示信息，不为正数则弹出提示信息
If Not IsNumeric(Txtndgyzcz.Text) Then
    MsgBox "输入的年度工业总产值只能是数字！"
    Txtndgyzcz.Text = ""
    Txtndgyzcz.SetFocus
    Exit Sub
Else
    If Val(Txtndgyzcz.Text) <= 0 Then
        MsgBox " 输入的年度工业总产值必须为正数"
        Txtndgyzcz.Text = ""
        Txtndgyzcz.SetFocus
        Exit Sub
    End If
End If


'判断输入的年度基建投资是否为空，若为空则弹出提示信息
If Txtndjjtz.Text = "" Then
    MsgBox "请输入年度基建投资！"
    Txtndjjtz.Text = ""
    Txtndjjtz.SetFocus
    Exit Sub
End If

'判断输入的年度基建投资是否为数字，若不为数字则弹出提示信息，不为正数则弹出提示信息
If Not IsNumeric(Txtndjjtz.Text) Then
    MsgBox "输入的年度基建投资只能是数字！"
    Txtndjjtz.Text = ""
    Txtndjjtz.SetFocus
    Exit Sub
Else
    If Val(Txtndjjtz.Text) <= 0 Then
        MsgBox " 输入的年度基建投资必须为正数"
        Txtndgyzcz.Text = ""
        Txtndgyzcz.SetFocus
        Exit Sub
    End If
End If

'如果是第1次预测
If Txt1.Text = "" Then

    '记录输入的年度工业总产值
    Txt1.Text = Txtndgyzcz.Text
    
    '记录输入的年度基建投资
    Txt2.Text = Txtndjjtz.Text
    
    '用刚输入的年度工业总产值更新yc表中的x1字段
    Dim strSQL As String
    strSQL = "update yc set x1 = " & Val(Txtndgyzcz.Text) & " where id = 1"
    ADOcn.Execute strSQL
    
    '用刚输入的年度基建投资更新yc表中的x2字段
    strSQL = "update yc set x2 = " & Val(Txtndjjtz.Text) & " where id = 1"
    ADOcn.Execute strSQL
    
    Dim Sqlmodel As String
    Dim ADOrs_Model As New Recordset
    Set ADOrs_Model.ActiveConnection = ADOcn

    '从zd表中获得二元线性回归方程模型文件的名称和路径
    Sqlmodel = " select * from zd where id = 2"
    ADOrs_Model.Open Sqlmodel
    filepath = ADOrs_Model("mxwjlj")
    modelName = ADOrs_Model("mxwjm")
    ADOrs_Model.Close
    
    '调用二元线性回归方程模型程序
    Shell (filepath & modelName)
    
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn

    '定义变量y , 存放从yc表中得到的字段y的值
    Dim y As Double
    
    '等待二元线性回归方程模型程序执行完毕
    For i = 0 To 9999
        strSQL = " select y from yc where id = 1"
        ADOrs.Open strSQL
        y = ADOrs("y")
        ADOrs.Close

        If y > 1 Then
            Exit For
        End If
        Next i
    
    If y > 1 Then
    Else
        For i = 0 To 9999
        strSQL = " select y from yc where id = 1"
        ADOrs.Open strSQL
        y = ADOrs("y")
        ADOrs.Close

        If y > 1 Then
            Exit For
        End If
        Next i
    End If
    Txtndhyl.Text = y

'如果不是第一次预测
Else

    Set ADOrs.ActiveConnection = ADOcn
    '变量y_1用于存放执行模型程序前从yc表中得到的字段y的值
    strSQL = " select x1,x2,y from yc where id = 1 "
    ADOrs.Open strSQL
    y_1 = ADOrs("y")
    ADOrs.Close
    
    '如果前一次输入的年度工业总产值、年度基建投资与这次输入的相同
    If Txtndgyzcz.Text = Txt1.Text And Txtndjjtz.Text = Txt2.Text Then
        Txtndhyl.Text = y_1
    
    '如果前一次输入的年度工业总产值、年度基建投资与这次输入的不同
    Else
        Txt1.Text = Txtndgyzcz.Text
        Txt2.Text = Txtndjjtz.Text
        strSQL = " update yc set x1=" & Val(Txtndgyzcz.Text) & " where id = 1 "
        ADOcn.Execute strSQL
        
        strSQL = " update yc set x2=" & Val(Txtndjjtz.Text) & " where id = 1 "
        ADOcn.Execute strSQL
         
        Set ADOrs_Model.ActiveConnection = ADOcn
        
        Sqlmodel = " select * from zd where id = 2"
        ADOrs_Model.Open Sqlmodel
        filepath = ADOrs_Model("mxwjlj")
        modelName = ADOrs_Model("mxwjm")
        ADOrs_Model.Close
        
        Shell (filepath & modelName)
        
        '变量y_2用于存放执行模型程序后从yc表中得到的字段y的值
        For i = 0 To 9999
            strSQL = "select y from yc where id = 1"
            ADOrs.Open strSQL
            y_2 = ADOrs("y")
            ADOrs.Close

            '二元线性回归方程模型程序执行完华
            If y_2 <> y_1 Then
                Exit For
            End If
            Next i
        
        Txtndhyl.Text = y_2
    
    End If
    
End If

Cmdjxyc.Enabled = True
Cmdyc.Enabled = False

End Sub

Private Sub Form_Load()

    Dim Sqlmodel As String
    Dim ADOrs_Model As New Recordset
    Set ADOrs_Model.ActiveConnection = ADOcn
    
    '从zd表中获得最小二乘法模型文件的名称和路径
    Sqlmodel = " select * from zd where id = 1 "
    ADOrs_Model.Open Sqlmodel
    filepath = ADOrs_Model("mxwjlj")
    modelName = ADOrs_Model("mxwjm")
    ADOrs_Model.Close
    
    '调用最小二乘法模型程序
    Shell (filepath & modelName)
    
    Dim strSQL As String
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn
    
    '定义变量a,b1,b2存放从cs表中得到的a,b1,b2字段的值
    Dim a As Double, b1 As Double, b2 As Double
    
    '等待最小二乘法模型程序执行完毕
    For i = 0 To 99999
        strSQL = " select * from cs "
        ADOrs.Open strSQL
        a = ADOrs("a")
        b1 = ADOrs("b1")
        b2 = ADOrs("b2")
        ADOrs.Close
        
        If a <> 0 Then
            Exit For
        End If
        Next i

    If a = 0 Then
        Shell (filepath & modelName)
    
    For i = 0 To 99999
        strSQL = " select * from cs "
        ADOrs.Open strSQL
        a = ADOrs("a")
        b1 = ADOrs("b1")
        b2 = ADOrs("b2")
        ADOrs.Close
        
        If a <> 0 Then
            Exit For
        End If
        Next i
    
    End If
    
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
    
    Txta.Text = m
    Txtb1.Text = n
    Txtb2.Text = p
    
    '让文本框成为只读
    Txta.Enabled = False
    Txtb1.Enabled = False
    Txtb2.Enabled = False
    Cmdjxyc.Enabled = False
    
End Sub
