Attribute VB_Name = "zkmk"
'����ȫ�ֶ������ADOcn�����ڴ��������ݿ������
Public ADOcn As Connection
Public Sub main()

    '�������ݿ������ַ��ͱ���strAccess,��Ϊ�丳ֵ,��������Access���ݿ�
    Dim strAccess As String
    strAccess = " Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = C:\hylyc\sj\hylyc.mdb"

    '����VB����DoEvents , ����ת��ϵͳ����Ȩ
    DoEvents

        '�����û�н��������ݿ������.�������´��봴��
        Set ADOcn = New Connection

        '����Access���ݿ�
        ADOcn.Open strAccess
    
        '��ʼ��cs��
        Dim strSQL As String
        
        strSQL = " delete * from cs"
        ADOcn.Execute strSQL

        strSQL = " insert into cs(id,a,b1,b2) values(0,0,0,0)"
        ADOcn.Execute strSQL

        '��ʼ��yc��
        strSQL = " delete * from yc"
        ADOcn.Execute strSQL

        strSQL = " insert into yc(id,x1,x2,y) values(1,0,0,0)"
        ADOcn.Execute strSQL
    
    '��ʾ������������
    wtms.Show

End Sub

