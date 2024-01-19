Attribute VB_Name = "zkmk"
'声明全局对象变量ADOcn，用于创建与数据库的连接
Public ADOcn As Connection
Public Sub main()

    '定义数据库连接字符型变量strAccess,并为其赋值,用以连接Access数据库
    Dim strAccess As String
    strAccess = " Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = C:\hylyc\sj\hylyc.mdb"

    '调用VB函数DoEvents , 用于转移系统控制权
    DoEvents

        '如果还没有建立与数据库的连接.则用以下代码创建
        Set ADOcn = New Connection

        '连接Access数据库
        ADOcn.Open strAccess
    
        '初始化cs表
        Dim strSQL As String
        
        strSQL = " delete * from cs"
        ADOcn.Execute strSQL

        strSQL = " insert into cs(id,a,b1,b2) values(0,0,0,0)"
        ADOcn.Execute strSQL

        '初始化yc表
        strSQL = " delete * from yc"
        ADOcn.Execute strSQL

        strSQL = " insert into yc(id,x1,x2,y) values(1,0,0,0)"
        ADOcn.Execute strSQL
    
    '显示问题描述窗体
    wtms.Show

End Sub

