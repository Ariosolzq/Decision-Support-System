Attribute VB_Name = "eyxxhg"
'����ȫ�ֶ������ADOcn , ���ڴ��������ݿ������
Public ADOcn As Connection
Sub Main()

'�������ݿ������ַ��ͱ�W�� strAccess,��Ϊ�丳ֵ����������Access���ݿ�
Dim strAccess As String
strAccess = " Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\hylyc\sj\hylyc.mdb"

'����VB����DoEvents,����ת��ϵͳ����Ȩ
DoEvents

    '�����û�н��������ݿ�����ӣ��������´��봴��
    Set ADOcn = New Connection
    ADOcn.Open strAccess
    Dim strSQL As String
    
    '����һ��¼������,����ADOcn����
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn
    strSQL = " select * from cs"

    'ִ��Select���
    ADOrs.Open strSQL

    '�������a,b1,b2�ֱ���cs���е�a,b1,b2�ֶ�ֵ
    '�������x1,x2,y�ֱ�����ȹ�ҵ�ܲ�ֵ,��Ȼ���Ͷ��,��Ȼ�����
    Dim a, b1, b2, x1, x2, y As Double
    a = ADOrs("a")
    b1 = ADOrs("b1")
    b2 = ADOrs("b2")
    ADOrs.Close
    
    '�������Sql1��Ϊ�丳ֵ���õ�������yc���е���ȹ�ҵ�ܲ�ֵ,��Ȼ���Ͷ��
    Sql1 = " select x1,x2 from yc where id = 1 "
    ADOrs.Open Sql1
    x1 = ADOrs("x1")
    x2 = ADOrs("x2")
    
    '������ȹ�ҵ�ܲ�ֵ,��Ȼ���Ͷ�ʣ�Ԥ����Ȼ�����
    y = Val(a) + Val(b1 * x1) + Val(b2 * x2)

    '�������Sql2��Ϊ�丳ֵ������yc���е���Ȼ�����
    Sql2 = " update yc set y = " & y & " where id = 1"
    ADOcn.Execute Sql2
    
    '�رռ�¼����������Ӷ���
    ADOrs.Close
ADOcn.Close
    
End Sub
