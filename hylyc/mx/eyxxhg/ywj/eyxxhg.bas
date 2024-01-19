Attribute VB_Name = "eyxxhg"
'声明全局对象变量ADOcn , 用于创建与数据库的连接
Public ADOcn As Connection
Sub Main()

'定义数据库连接字符型变W： strAccess,并为其赋值，用以连接Access数据库
Dim strAccess As String
strAccess = " Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\hylyc\sj\hylyc.mdb"

'调用VB函数DoEvents,用于转移系统控制权
DoEvents

    '如果还没有建立与数据库的连接，则用以下代码创建
    Set ADOcn = New Connection
    ADOcn.Open strAccess
    Dim strSQL As String
    
    '声明一记录集对象,并与ADOcn关联
    Dim ADOrs As New Recordset
    Set ADOrs.ActiveConnection = ADOcn
    strSQL = " select * from cs"

    '执行Select语句
    ADOrs.Open strSQL

    '定义变量a,b1,b2分别存放cs表中的a,b1,b2字段值
    '定义变量x1,x2,y分别存放年度工业总产值,年度基建投资,年度货运量
    Dim a, b1, b2, x1, x2, y As Double
    a = ADOrs("a")
    b1 = ADOrs("b1")
    b2 = ADOrs("b2")
    ADOrs.Close
    
    '定义变量Sql1并为其赋值，得到保存在yc表中的年度工业总产值,年度基建投资
    Sql1 = " select x1,x2 from yc where id = 1 "
    ADOrs.Open Sql1
    x1 = ADOrs("x1")
    x2 = ADOrs("x2")
    
    '根据年度工业总产值,年度基建投资，预测年度货运量
    y = Val(a) + Val(b1 * x1) + Val(b2 * x2)

    '定义变量Sql2并为其赋值，更新yc表中的年度货运量
    Sql2 = " update yc set y = " & y & " where id = 1"
    ADOcn.Execute Sql2
    
    '关闭记录集对象和连接对象
    ADOrs.Close
ADOcn.Close
    
End Sub
