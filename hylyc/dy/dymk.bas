Attribute VB_Name = "dymk"
'����ȫ�ֶ������ADOcn , ���ڴ��������ݿ������
Public ADOcn As Connection
Public Sub Main()

    '�������ݿ������ַ��ͱ���slrAccess,��Ϊ�丳ֵ,��������Access���ݿ�
    Dim strAccess As String
    strAccess = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = C:\hylyc\sj\hylyc.mdb"
    
    '����SQL�ַ��ͱ���
    Dim strSQL As String
    
    '����VB����DoEvents , ����ת��ϵͳ����Ȩ
    DoEvents

        '���������ݿ������
        Set ADOcn = New Connection
        
        '����Access���ݿ�
        ADOcn.Open strAccess

        '��ʼ��cs��
        strSQL = "delete * from cs"
        ADOcn.Execute strSQL
        
        strSQL = " insert into cs(a ,b1,b2) values(0,0,0)"
        ADOcn.Execute strSQL
        
        '��ʾ������
        dy.Show
        
End Sub
