VERSION 5.00
Begin VB.Form zk 
   Caption         =   "��Ȼ�����Ԥ�����֧��ϵͳ"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9855
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "�˳�"
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Cmdjxyc 
      Caption         =   "����Ԥ��"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Cmdyc 
      Caption         =   "Ԥ��"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ȹ�ҵ�ܲ�ֵ����Ȼ���Ͷ��"
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
         Caption         =   "��λ��������֣�"
         Height          =   375
         Left            =   7080
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "��Ȼ�����"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "��λ������Ԫ��"
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "��λ������Ԫ��"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "��Ȼ���Ͷ��"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "��ȹ�ҵ�ܲ�ֵ"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����a��b1��b2��ֵ"
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
         Caption         =   "�ع�ϵ��b2"
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "�ع�ϵ��b1"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "������"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "��Ȼ�����Ԥ�����֧��ϵͳ"
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

'�ж��������ȹ�ҵ�ܲ�ֵ�Ƿ�Ϊ�գ���Ϊ���򵯳���ʾ��Ϣ
If Txtndgyzcz.Text = "" Then
    MsgBox "��������ȹ�ҵ�ܲ�ֵ��"
    Txtndgyzcz.Text = ""
    Txtndgyzcz.SetFocus
    Exit Sub
End If

'�ж��������ȹ�ҵ�ܲ�ֵ�Ƿ�Ϊ���֣�����Ϊ�����򵯳���ʾ��Ϣ����Ϊ�����򵯳���ʾ��Ϣ
If Not IsNumeric(Txtndgyzcz.Text) Then
    MsgBox "�������ȹ�ҵ�ܲ�ֵֻ�������֣�"
    Txtndgyzcz.Text = ""
    Txtndgyzcz.SetFocus
    Exit Sub
Else
    If Val(Txtndgyzcz.Text) <= 0 Then
        MsgBox " �������ȹ�ҵ�ܲ�ֵ����Ϊ����"
        Txtndgyzcz.Text = ""
        Txtndgyzcz.SetFocus
        Exit Sub
    End If
End If


'�ж��������Ȼ���Ͷ���Ƿ�Ϊ�գ���Ϊ���򵯳���ʾ��Ϣ
If Txtndjjtz.Text = "" Then
    MsgBox "��������Ȼ���Ͷ�ʣ�"
    Txtndjjtz.Text = ""
    Txtndjjtz.SetFocus
    Exit Sub
End If

'�ж��������Ȼ���Ͷ���Ƿ�Ϊ���֣�����Ϊ�����򵯳���ʾ��Ϣ����Ϊ�����򵯳���ʾ��Ϣ
If Not IsNumeric(Txtndjjtz.Text) Then
    MsgBox "�������Ȼ���Ͷ��ֻ�������֣�"
    Txtndjjtz.Text = ""
    Txtndjjtz.SetFocus
    Exit Sub
Else
    If Val(Txtndjjtz.Text) <= 0 Then
        MsgBox " �������Ȼ���Ͷ�ʱ���Ϊ����"
        Txtndgyzcz.Text = ""
        Txtndgyzcz.SetFocus
        Exit Sub
    End If
End If

'����ǵ�1��Ԥ��
If Txt1.Text = "" Then

    '��¼�������ȹ�ҵ�ܲ�ֵ
    Txt1.Text = Txtndgyzcz.Text
    
    '��¼�������Ȼ���Ͷ��
    Txt2.Text = Txtndjjtz.Text
    
    '�ø��������ȹ�ҵ�ܲ�ֵ����yc���е�x1�ֶ�
    Dim strSQL As String
    strSQL = "update yc set x1 = " & Val(Txtndgyzcz.Text) & " where id = 1"
    ADOcn.Execute strSQL
    
    '�ø��������Ȼ���Ͷ�ʸ���yc���е�x2�ֶ�
    strSQL = "update yc set x2 = " & Val(Txtndjjtz.Text) & " where id = 1"
    ADOcn.Execute strSQL
    
    Dim Sqlmodel As String
    Dim ADOrs_Model As New Recordset
    Set ADOrs_Model.ActiveConnection = ADOcn

    '��zd���л�ö�Ԫ���Իع鷽��ģ���ļ������ƺ�·��
    Sqlmodel = " select * from zd where id = 2"
    ADOrs_Model.Open Sqlmodel
    filepath = ADOrs_Model("mxwjlj")
    modelName = ADOrs_Model("mxwjm")
    ADOrs_Model.Close
    
    '���ö�Ԫ���Իع鷽��ģ�ͳ���
    Shell (filepath & modelName)
    
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn

    '�������y , ��Ŵ�yc���еõ����ֶ�y��ֵ
    Dim y As Double
    
    '�ȴ���Ԫ���Իع鷽��ģ�ͳ���ִ�����
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

'������ǵ�һ��Ԥ��
Else

    Set ADOrs.ActiveConnection = ADOcn
    '����y_1���ڴ��ִ��ģ�ͳ���ǰ��yc���еõ����ֶ�y��ֵ
    strSQL = " select x1,x2,y from yc where id = 1 "
    ADOrs.Open strSQL
    y_1 = ADOrs("y")
    ADOrs.Close
    
    '���ǰһ���������ȹ�ҵ�ܲ�ֵ����Ȼ���Ͷ��������������ͬ
    If Txtndgyzcz.Text = Txt1.Text And Txtndjjtz.Text = Txt2.Text Then
        Txtndhyl.Text = y_1
    
    '���ǰһ���������ȹ�ҵ�ܲ�ֵ����Ȼ���Ͷ�����������Ĳ�ͬ
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
        
        '����y_2���ڴ��ִ��ģ�ͳ�����yc���еõ����ֶ�y��ֵ
        For i = 0 To 9999
            strSQL = "select y from yc where id = 1"
            ADOrs.Open strSQL
            y_2 = ADOrs("y")
            ADOrs.Close

            '��Ԫ���Իع鷽��ģ�ͳ���ִ���껪
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
    
    '��zd���л����С���˷�ģ���ļ������ƺ�·��
    Sqlmodel = " select * from zd where id = 1 "
    ADOrs_Model.Open Sqlmodel
    filepath = ADOrs_Model("mxwjlj")
    modelName = ADOrs_Model("mxwjm")
    ADOrs_Model.Close
    
    '������С���˷�ģ�ͳ���
    Shell (filepath & modelName)
    
    Dim strSQL As String
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn
    
    '�������a,b1,b2��Ŵ�cs���еõ���a,b1,b2�ֶε�ֵ
    Dim a As Double, b1 As Double, b2 As Double
    
    '�ȴ���С���˷�ģ�ͳ���ִ�����
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
    
    '�������m,n,p�ֱ��Ž�a,b1,b2ת��Ϊ�ַ������ֵ
    Dim m, n, p As String
    m = CStr(a)
    n = CStr(b1)
    p = CStr(b2)
    
    '��m���ĵ�һ������Ϊ"."��������һ���ַ�ǰ�ӡ�0��
    If Left(m, 1) = "." Then
        m = "0" & m
    End If
    
    '��n���ĵ�һ������Ϊ"."��������һ���ַ�ǰ�ӡ�0��
    If Left(n, 1) = "." Then
        n = "0" & n
    End If
    
    '��p���ĵ�һ������Ϊ"."��������һ���ַ�ǰ�ӡ�0��
    If Left(p, 1) = "." Then
        p = "0" & p
    End If
    
    Txta.Text = m
    Txtb1.Text = n
    Txtb2.Text = p
    
    '���ı����Ϊֻ��
    Txta.Enabled = False
    Txtb1.Enabled = False
    Txtb2.Enabled = False
    Cmdjxyc.Enabled = False
    
End Sub
