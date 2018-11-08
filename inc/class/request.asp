<%
'@title: Class_Request
'@author: ekede.com
'@date: 2017-12-6
'@description: Request对象,服务器信息

Class Class_Request

    '@bytes: byte二进制流
    '@querystr: Get集合
	'@forms: Form集合
    '@servers: Server集合
    '@status404: 404状态
    '@standardAddr: 标准网址
    '@realAddr: 真实网址
	
    Dim querystr, forms, servers
	
	Dim https, queryStrings
	
    Dim status404, standardAddr, realAddr
	
    '@baseAddr: script目录所在地址
    '@basePicAddr: 图片网站根目录
	
    Dim baseAddr,basePicAddr
	
	'二进制流
    Public Property Get bytes()
        bytes  = Request.BinaryRead(Request.TotalBytes) 'multipart/form-data
    End Property
    '
    Private Sub Class_Initialize()
        Set querystr = Server.CreateObject("Scripting.Dictionary")
        Set forms = Server.CreateObject("Scripting.Dictionary")
        Set servers = Server.CreateObject("Scripting.Dictionary")
        '
        ColDic Request.ServerVariables, servers
        ColDic request.QueryString, querystr
		If instr(servers("HTTP_CONTENT_TYPE"),"application/x-www-form-urlencoded")>0 Then
           ColDic Request.Form, forms
		End If
        '协议
        If servers("Https") = "on" Then
            https = "https://"
        Else 'off
            https = "http://"
        End If
        '修正IIS5 加port
		queryStrings = servers("QUERY_STRING") '参数:a=1&b=2
        If (servers("SERVER_SOFTWARE") = "Microsoft-IIS/5.1" And InStr(QueryStrings, ";")>0) Then
            queryStrings = Replace(queryStrings, servers("SERVER_NAME"), servers("SERVER_NAME")&":"&servers("SERVER_PORT"))
            queryStrings = Replace(queryStrings, "http://", https) '验证
        End If
        '
        Real_Addr()     'realAddr
        status_404()    'status404
        Standard_Addr() 'standardAddr
        Base_Addr()     'baseAddr
		Base_Pic_Addr() 'basePicAddr
    End Sub


    Private Sub class_terminate()
        Set querystr = Nothing
        Set forms = Nothing
        Set servers = Nothing
    End Sub

    '404状态

    Private Sub Status_404()
        If InStr(servers("QUERY_STRING"), ";")>0 Then
            status404 = True
        Else
            status404 = False
        End If
    End Sub

    'realAddr : http://localhost/sys/404.asp?404;http://localhost:80/sys/en/

    Private Sub Real_Addr()
        Dim Url
        Url = https&servers("HTTP_HOST")&servers("URL")
        If queryStrings <>"" Then Url = Url&"?"& queryStrings
        If servers("SERVER_PORT") = "80" Then Url = Replace(Url, ":80/", "/") '去80
        If servers("SERVER_PORT") = "443" Then Url = Replace(Url, ":443/", "/") '去443
        realAddr = Url
    End Sub

    'standardAddr : http://localhost:80/sys/en/

    Private Sub Standard_Addr()
        Dim Url
        If queryStrings = "" Then
            Url = https&servers("HTTP_HOST")&servers("URL")
        Else
            If InStr(queryStrings, ";")>0 Then
                Url = Right(queryStrings, Len(queryStrings) - InStr(queryStrings, ";"))
            Else
                Url = https&servers("HTTP_HOST")&servers("URL")&"?"& queryStrings
            End If
        End If
        If servers("SERVER_PORT") = "80" Then Url = Replace(Url, ":80/", "/") '去80  http
        If servers("SERVER_PORT") = "443" Then Url = Replace(Url, ":443/", "/") '去443 https
        standardAddr = Url
    End Sub

    'baseAddr : http://localhost/sys/

    Private Sub Base_Addr()
        Dim Url
        Url = Replace(https&servers("HTTP_HOST")&servers("URL"), "index.asp", "")
        If servers("SERVER_PORT") = "80" Then Url = Replace(Url, ":80/", "/") '去80  http
        If servers("SERVER_PORT") = "443" Then Url = Replace(Url, ":443/", "/") '去443 https
        baseAddr = Url
    End Sub
	
    Private Sub Base_Pic_Addr() '回根目录
	    Dim counter,arr,str,i
	    if PATH_ROOT <> "" then
		   If Instr(PATH_ROOT,"../")>0 Then
	          counter=ubound(Split(PATH_ROOT,"../"))
			  arr=split(baseAddr,"/")
			  for i = 0 to Ubound(arr)-(counter+1)
			     str=str&arr(i)&"/"
			  next
		   Else
		      str = baseAddr&PATH_ROOT
		   End If
		Else
		   str = baseAddr
		End If
        basePicAddr = str
    End Sub
	
    '调用 - collection转换dictionary

    Public Sub ColDic(col, dic)
        For Each Keys in col
            dic(Keys) = col(Keys)
        Next
    End Sub

    '服务器参数举例
	'servers("SERVER_SOFTWARE") 'Microsoft-IIS/5.1  Microsoft-IIS/6.0  Microsoft-IIS/7.5
	'servers("HTTP_HOST")       'localhost:8080
	'servers("SERVER_NAME")     'localhost
	'servers("SERVER_PORT")     '端口:80,8080
	'servers("SCRIPT_NAME")     '网页:/test/xxx.asp
	'servers("URL")             '网页:/test/xxx.asp
	'servers("QUERY_STRING")    '参数:a=1&b=2
	'servers("Remote_Addr")     'IP: 127.0.0.1 计算地区
	'servers("HTTP_REFERER")    '来访页: http://www.xxx.com/xxx.asp?id=xxx 完整地址
	'servers("HTTP_USER_AGENT") '操作系统,浏览器,版本
	'servers("HTTP_ACCEPT_LANGUAGE") '语言
	'servers("HTTP_ACCEPT")     '文档类型

End Class
%>