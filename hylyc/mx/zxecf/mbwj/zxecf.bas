Attribute VB_Name = "zxecf"
'����ȫ�ֶ�����ADOcn_mx,���ڴ��������ݿ������
Public ADOcn_mx As Connection
Sub Main()

    '�������ݿ������ַ��ͱ���strAccess����Ϊ�丳ֵ����������Access���ݿ�
    Dim strAccess As String
    strAccess = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = C:\hylyc\sj\hylyc.mdb"
    
    '����VB����DoEvents������ת��ϵͳ����Ȩ
    DoEvents
    
    '���������ݿ������
    Set ADOcn_mx = New Connection
    ADOcn_mx.Open strAccess

    Dim strSQL As String
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn_mx
    strSQL = " select * from tjsj"
    
    'ִ��Select��䣬1����ֻ��
    ADOrs.Open strSQL, ADOcn_mx, 1
    
    '�������num,��ŴӼ�¼������������и���
    Dim num As Integer
    num = ADOrs.RecordCount
    
    '�������v1���x1��ֵ������v2���x2��ֵ.����v3���y��ֵ
    '����sum1���x1�ĺͣ�sum2���x2�ĺͣ�sum3���x1��ƽ���ͣ�sum4���x2��ƽ����
    '����sum5���x1*x2�ĺͣ�sum6���x1*y�ĺͣ�sum7���x2*y�ĺͣ�sum8���y�ĺͣ�i��ѭ������
    '����b1��Żع�ϵ��b1��ֵ��b2��Żع�ϵ��b2��ֵ
    Dim i As Integer, v1, v2, v3, sum1, sum2, sum3, sum4, sum5, sum6, sum7, sum8 As Double
    For i = 1 To num
        v1 = ADOrs("x1").Value
        v2 = ADOrs("x2").Value
        v3 = ADOrs("y").Value
    
        '����x1�ĺ�
        sum1 = sum1 + v1
        
        '����x2�ĺ�
        sum2 = sum2 + v2
        
        '����y�ĺ�
        sum8 = sum8 + v3
        
        '����x1ƽ���ĺ�
        sum3 = sum3 + v1 * v1
        
        '����x1ƽ���ĺ�
        sum4 = sum4 + v2 * v2
        
        '����x1*x2�ĺ�
        
        sum5 = sum5 + v1 * v2
        
        '����x1*y�ĺ�
        sum6 = sum6 + v1 * v3
        
        '����x2*y�ĺ�
        sum7 = sum7 + v2 * v3
        ADOrs.MoveNext
        
Next i

'�رռ�¼��
ADOrs.Close

'�������Ҳ��ų�����,�����M b��Żع�ϵ��.�����* Sqll ,��Ÿ����ַ���
Dim a As Double, b1 As Double, b2 As Double, Sql1 As String
b1 = (sum4 * sum6 - sum5 * sum7) / (sum3 * sum4 - sum5 * sum5)
b2 = (sum3 * sum7 - sum5 * sum6) / (sum3 * sum4 - sum5 * sum5)
a = (sum8 - b1 * sum1 - b2 * sum2) / num

'�ü������a,b1,b2��ֵ�������ݿ��в�����cs�е�a,b1,b2��ֵ
Sql1 = " update cs set a = " & a & " , b1 = " & b1 & ", b2= " & b2

'ִ�и������
ADOcn_mx.Execute Sql1

'�ر����ݿ����Ӷ���
ADOcn_mx.Close

End Sub
