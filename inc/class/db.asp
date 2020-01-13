<%
'@title: Class_DB
'@author: ekede.com
'@date: 2017-11-29
'@description: 数据库操作类

Class Class_DB
    '
    Private conn_
	Private isDebug_
	
    '@conn: 数据库连接
	
    Public Property Get conn
        Set conn = conn_
    End Property

    Public Property Let conn(Value)
        Set conn_ = Value
    End Property

    '初始化

    Private Sub Class_Initialize()
		If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
    End Sub

    Private Sub Class_Terminate()
    End Sub
	
    '@OpenConn(db_type, db_path,db_version, db_name, db_user, db_pass): 打开数据库连接

    Public Sub OpenConn(db_type, db_version, db_path, db_name, db_user, db_pass)
        on error resume next
        Dim dpath, TempStr

        '文件数据库检测
        If db_type = 1 Or db_type = 2 Then dpath=Server.MapPath(PATH_ROOT&db_path&db_name)
        '
        Select Case db_type
            Case 1 'Access
			   If db_version = 1 Then
                  TempStr = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="&dpath&";DefaultDir=;"
			   ElseIf db_version = 2 Then
			      TempStr = "Provider=Microsoft.jet.OLEDB.4.0;Data Source="&dpath
			   Else
			      TempStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &dpath
			   End If
            Case 2 'Excel
			   If db_version = 1 Then
                  TempStr = "DRIVER={Microsoft Excel Driver (*.xls)};DBQ="&dpath&";DefaultDir=;" 
			   ElseIf db_version = 2 Then
			      TempStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&dpath&";Extended Properties=Excel 8.0;"
			   Else
			      TempStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&dpath&";Extended Properties=Excel 12.0;"
			   End If
			Case 3 'MSQL
                TempStr = "Driver={SQL Server};server="&db_path&";uid="&db_user&";pwd="&db_pass&";database="&db_name&""
            Case 4 'MYSQL
                TempStr = "Driver={mySQL};Server="&db_path&";Port=3306;Option=131072;Stmt=; Database="&db_name&";Uid="&db_user&";Pwd="&db_pass&";"
            Case 5 'ORACLE
                TempStr = "Driver={Microsoft ODBC for Oracle};Server="&db_path&";Uid="&db_user&";Pwd="&db_pass&";"
            Case 6 'Godaddy
                TempStr = "filedsn=" & Server.MapPath(db_path) & "/" & db_name
            Case Else
			    OutErr("check database setting")
        End Select
        '
        Set conn_ = server.CreateObject("ADODB.CONNECTION")
        conn_.Open TempStr
        If Err Then OutErr(Err.description)

    End Sub

    '@CloseConn(): 关闭数据库连接

    Public Sub CloseConn()
        conn_.Close
        Set conn_ = Nothing
    End Sub

    '@Query(sTable, sFileds, sWhere, sOrder, sGroup, sCursorType, sLockType): 查 - 返回Recordset对象
	'如果单单是读取，不涉及更新操作，那就用1，1
	'如果涉及读取及更新操作，可以用1,3 或3,2

    Public Function Query(sTable, sFileds, sWhere, sOrder, sGroup, sCursorType, sLockType)
        On Error Resume Next
        Dim sql, rs
        sql = SqlQuery(sTable, sFileds, sWhere, sOrder, sGroup)
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, conn_, sCursorType, sLockType
        If Err Then
            OutErr(sql&chr(10)&Err.Description)
            rs.Close
            Set rs = Nothing
            Exit Function
        Else
            Set Query = rs
        End If
    End Function

    '@Add(sTable, sFileds, sValues): 增 - 返回id

    Public Function Add(sTable, sFileds, sValues)
        On Error Resume Next
        Dim sql
        sql = SqlAdd(sTable, sFileds, sValues)
        '
        conn_.Execute(sql)
        If Err Then
            OutErr(sql&chr(10)&Err.Description)
            Add = 0
            Exit Function
        Else
            Add = AutoId(sTable) -1 '新添加id
        End If
    End Function

    '@Edit(sTable, sFileds, sValues, sWhere): 改 - 返回状态

    Public Function Edit(sTable, sFileds, sValues, sWhere)
        On Error Resume Next
        Dim sql
        sql = SqlEdit(sTable, sFileds, sValues, sWhere)
        '
        conn_.Execute(sql)
        '
        If Err Then
            OutErr(sql&chr(10)&Err.Description)
            Edit = 0
            Exit Function
        Else
            Edit = 1
        End If
    End Function

    '@Del(sTable, sWhere): 删 - 返回状态

    Public Function Del(sTable, sWhere)
        On Error Resume Next
        Dim sql
        sql = SqlDel(sTable, sWhere)
        '
        conn_.Execute(sql)
        If Err Then
            OutErr(sql&chr(10)&Err.Description)
            Del = 0
            Exit Function
        Else
            Del = AutoId(sTable) -1 '新添加id
        End If
    End Function
	
    '@AutoId(ByVal TableName):自动ID

    Public Function AutoId(ByVal TableName)
        On Error Resume Next
        Dim rs, Sql, TempNo
        Set rs = Server.CreateObject("adodb.recordset")
        Sql = "SELECT * FROM "&TableName
        rs.Open Sql, conn_, 3, 3
        If rs.EOF Then
            AutoId = 1
        Else
            Do While Not rs.EOF
                TempNo = rs.Fields(0).Value
                rs.MoveNext
                If rs.EOF Then AutoId = TempNo + 1
            Loop
        End If
        rs.Close
        Set rs = Nothing
        '
        If Err Then
		    OutErr(sql&chr(10)&Err.Description)
            AutoId = 0
            Exit Function
        End If

    End Function
	
	'@SqlQuery(sTable, sFileds, sWhere, sOrder, sGroup): sql语句查询

    Public Function SqlQuery(sTable, sFileds, sWhere, sOrder, sGroup)
	    SqlQuery = SqlBuild("select", sTable, sFileds, "", sWhere, sOrder, sGroup)
	End Function
	
	'@SqlAdd(sTable, sFileds, sValues): sql语句添加
	
    Public Function SqlAdd(sTable, sFileds, sValues)
        SqlAdd = SqlBuild("insert", sTable, sFileds, sValues, "", "", "")
    End Function
	
	'@SqlEdit(sTable, sFileds, sValues, sWhere): sql语句修改
	
    Public Function SqlEdit(sTable, sFileds, sValues, sWhere)
        SqlEdit = SqlBuild("update", sTable, sFileds, sValues, sWhere, "", "")
    End Function
	
	'@SqlDel(sTable, sWhere): sql语句删除
	
    Public Function SqlDel(sTable, sWhere)
        SqlDel = SqlBuild("delete", sTable, "", "", sWhere, "", "")
    End Function
	
    'sql语句生成

    Private Function SqlBuild(sType, sTable, sFileds, sValues, sWhere, sOrder, sGroup)
        Dim TempStr
        '主语句
        Select Case sType
            Case "select"
                If sFileds = "" Then sFileds = "*"
                TempStr = "select "&sFileds&" from "&sTable&""
            Case "delete"
                TempStr = "delete from "&sTable&""
            Case "insert"
                TempStr = "insert into "&sTable&" ("&sFileds&") values ("&sValues&")"
            Case "update"
                If sFileds<>"" Then
                    TempStr = "update "&sTable&" set "&sFileds&" = "&sValues&""
                Else
                    TempStr = "update "&sTable&" set "&sValues&""
                End If
        End Select
        '条件，排序
        If sWhere<>"" Then TempStr = TempStr&" where "&sWhere
        If sGroup<>"" Then TempStr = TempStr&" Group by "&sGroup
        If sOrder<>"" Then TempStr = TempStr&" order by "&sOrder
        '
        SqlBuild = TempStr
    End Function
	
	'@SqlExecute(sql): sql语句执行
	
	Public Function SqlExecute(sql)
	    On Error Resume Next
	    If  left(lcase(sql),6)="select" Then
		    Set SqlExecute = conn_.Execute(sql)
			If Err Then OutErr(sql&chr(10)&Err.Description)
		Else
			conn_.Execute(sql)
			If Err Then
			   SqlExecute=0
			   OutErr(sql&chr(10)&Err.Description)
			Else
			   SqlExecute=1
			End If
		End If
	End Function
	
	'@CreateAccess(db_path,db_name): 创建Access数据库
	
    Public Sub CreateAccess(db)
        on error resume Next
		Dim adox
		Set adox= Server.CreateObject("ADOX.Catalog") 
		    adox.Create "Provider = Microsoft.jet.OLEDB.4.0;Data Source=" & Server.MapPath(PATH_ROOT&db)
		Set adox= Nothing
        If Err Then OutErr("can't create access"&chr(10)&Err.Description)
    End Sub
	
	'@CompactAccess(db1,db2): 压缩Access数据库
	
    Public Sub CompactAccess(db1,db2)
        on error resume next
		Set Engine = CreateObject("JRO.JetEngine")
			Engine.CompactDatabase _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(PATH_ROOT&db1), _
			"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(PATH_ROOT&db2) 
		Set Engine = nothing 
        If Err Then OutErr("can't compact access"&chr(10)&Err.Description)
    End Sub
	
	'错误提示

	Public Sub OutErr(ErrMsg)
	    Err.Clear
		If isDebug_ = true Then
			Response.charset = "utf-8"
			Response.Write ErrMsg
			Response.End
		End If
	End Sub

End Class
%>