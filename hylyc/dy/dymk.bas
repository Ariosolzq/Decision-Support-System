Attribute VB_Name = "dymk"
'声明全局对象变量ADOcn , 用于创建与数据库的连接
Public ADOcn As Connection
Public Sub Main()

    '定义数据库连接字符型变量slrAccess,并为其赋值,用以连接Access数据库
    Dim strAccess As String
    strAccess = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = C:\hylyc\sj\hylyc.mdb"
    
    '定义SQL字符型变量
    Dim strSQL As String
    
    '调用VB函数DoEvents , 用于转移系统控制权
    DoEvents

        '建立与数据库的连接
        Set ADOcn = New Connection
        
        '连接Access数据库
        ADOcn.Open strAccess

        '初始化cs表
        strSQL = "delete * from cs"
        ADOcn.Execute strSQL
        
        strSQL = " insert into cs(a ,b1,b2) values(0,0,0)"
        ADOcn.Execute strSQL
        
        '显示主窗体
        dy.Show
        
End Sub
