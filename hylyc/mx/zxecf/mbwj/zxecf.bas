Attribute VB_Name = "zxecf"
'声明全局对象变出ADOcn_mx,用于创建与数据库的连接
Public ADOcn_mx As Connection
Sub Main()

    '定义数据库连接字符型变量strAccess，并为其赋值，用以连接Access数据库
    Dim strAccess As String
    strAccess = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = C:\hylyc\sj\hylyc.mdb"
    
    '调用VB函数DoEvents，用于转移系统控制权
    DoEvents
    
    '建立与数据库的连接
    Set ADOcn_mx = New Connection
    ADOcn_mx.Open strAccess

    Dim strSQL As String
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn_mx
    strSQL = " select * from tjsj"
    
    '执行Select语句，1代表只读
    ADOrs.Open strSQL, ADOcn_mx, 1
    
    '定义变量num,存放从记录集读入的数据行个数
    Dim num As Integer
    num = ADOrs.RecordCount
    
    '定义变量v1存放x1的值，变量v2存放x2的值.变量v3存放y的值
    '变量sum1存放x1的和，sum2存放x2的和，sum3存放x1的平方和，sum4存放x2的平方和
    '变量sum5存放x1*x2的和，sum6存放x1*y的和，sum7存放x2*y的和，sum8存放y的和，i是循环变量
    '变量b1存放回归系数b1的值，b2存放回归系数b2的值
    Dim i As Integer, v1, v2, v3, sum1, sum2, sum3, sum4, sum5, sum6, sum7, sum8 As Double
    For i = 1 To num
        v1 = ADOrs("x1").Value
        v2 = ADOrs("x2").Value
        v3 = ADOrs("y").Value
    
        '计算x1的和
        sum1 = sum1 + v1
        
        '计算x2的和
        sum2 = sum2 + v2
        
        '计算y的和
        sum8 = sum8 + v3
        
        '计算x1平方的和
        sum3 = sum3 + v1 * v1
        
        '计算x1平方的和
        sum4 = sum4 + v2 * v2
        
        '计算x1*x2的和
        
        sum5 = sum5 + v1 * v2
        
        '计算x1*y的和
        sum6 = sum6 + v1 * v3
        
        '计算x2*y的和
        sum7 = sum7 + v2 * v3
        ADOrs.MoveNext
        
Next i

'关闭记录集
ADOrs.Close

'定义变量也存放常数项,定义变M b存放回归系数.定义变* Sqll ,存放更新字符中
Dim a As Double, b1 As Double, b2 As Double, Sql1 As String
b1 = (sum4 * sum6 - sum5 * sum7) / (sum3 * sum4 - sum5 * sum5)
b2 = (sum3 * sum7 - sum5 * sum6) / (sum3 * sum4 - sum5 * sum5)
a = (sum8 - b1 * sum1 - b2 * sum2) / num

'用计算出的a,b1,b2的值更新数据库中参数表cs中的a,b1,b2的值
Sql1 = " update cs set a = " & a & " , b1 = " & b1 & ", b2= " & b2

'执行更新语句
ADOcn_mx.Execute Sql1

'关闭数据库连接对象
ADOcn_mx.Close

End Sub
