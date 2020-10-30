<%
'@title: Class_Ext_Http
'@author: ekede.com
'@date: 2018-08-24
'@description: 模拟Http请求

Class Class_Ext_Http

    Private isDebug_
    Private slResolveTimeout, slConnectTimeout, slSendTimeout, slReceiveTimeout
    Private http_
    Private form_
	Private header_
	Private cookie_
	'
	Private cache_,fun_
	
    '@rStatus: 请求状态
    '@rHeader: header信息
	'@rBody: 内容二进制
	'@rText: 内容文本
	Dim rStatus,rHeader,rBody,rText
	
    '@rCookie: cookie字符串

    Public Property Get rCookie()
	    rCookie = GetCookie()
    End Property

	'@cookie_on: 是否开启cookie
	
    Dim cookie_on
	
    '@cache: cache对象依赖

    Public Property Let cache(Value)
        Set cache_ = Value
    End Property
	
    '@items: item直接赋值,字符串或字节数组都可以

    Public Property Let items(Value)
        form_ = Value
    End Property
	
    '@isDebug: 是否设置为调试模式
	
    Public Property Let isDebug(Value) 
        isDebug_ = Value
    End Property
	

    Private Sub Class_Initialize()
		If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
        slResolveTimeout = 20000   '解析DNS名字的超时时间,20秒
        slConnectTimeout = 20000   '建立Winsock连接的超时时间,20秒
        slSendTimeout = 30000      '发送数据的超时时间,30秒
        slReceiveTimeout = 30000   '接收response的超时时间,30秒
        Set http_ = Server.CreateObject("MSXML2.ServerXMLHTTP")
        Set header_ = Server.CreateObject("Scripting.Dictionary")
		'
		cookie_on = False
        Set cookie_ = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate
        Set cookie_ = Nothing
		'
        Set header_ = Nothing
        Set http_ = Nothing
    End Sub

    '@Send(ByRef method,ByVal url): 发送请求

    Public Sub Send(ByRef method,ByVal url)
	    On Error Resume Next
	    Dim c,s
		'Header Read
		If cookie_on Then c = GetC(url)
		'
		http_.setTimeouts slResolveTimeout, slConnectTimeout, slSendTimeout, slReceiveTimeout
		'GET
		If method = "GET" Then
			If form_ <> "" Then
			  If Instr(url,"?")>0 Then
				 url = url&"&"&form_
			  Else
				 url = url&"?"&form_
			  End If
			End If
			http_.Open "GET", url, False
			'
			If c <> "" Then http_.setRequestHeader "Cookie", c
			For Each k in header_
				http_.setRequestHeader k, header_(k)
			Next
			'
			http_.Send()
			If Err Then OutErr("Http Get Fail:"&Err.Number&":"&Err.Description)
		End If
		'POST
		If method = "POST" Then
		    http_.Open method, url, False
		    http_.setRequestHeader "Content-Length", Len(form_)
			'
			If c <> "" Then http_.setRequestHeader "Cookie", c
			For Each k in header_
				http_.setRequestHeader k, header_(k)
			Next
			'
			http_.Send(form_)
			If Err Then OutErr("Http POST Fail:"&Err.Number&":"&Err.Description)
	    End If
		'
		While http_.readyState <> 4
			http_.waitForResponse 1000
		Wend
		'
		rStatus = http_.Status
		rHeader = http_.getAllResponseHeaders()
		rText = http_.ResponseText
		rBody = http_.ResponseBody
		'Header Write
	    If cookie_on Then SetC url,rHeader
		
    End Sub

    '@AddItem(ByRef Key, ByRef Value): 添加表单键值

    Public Sub AddItem(ByRef Key, ByRef Values)
        On Error Resume Next
        If form_ = "" Then
            form_ = Key + "=" + Server.URLEncode(Values)
        Else
            form_ = form_ + "&" + Key + "=" + Server.URLEncode(Values)
        End If
    End Sub

    '@SetHeader(key, Value): 设置头信息

    Public Sub SetHeader(ByRef key, ByRef Values)
        header_(key) = Values
    End Sub
	
    '读Cookie
	
    Private Function GetC(ByRef url)
	    Dim c
        key = "cookie/"&GetDomainKey(url)&".txt"
	    c = cache_.GetCache(key)
		If c = -1 Then
		   GetC = ""
           CleanCookie()
		Else
		   GetC = c
           CleanCookie()
		   SetCookie(c)
		End If
    End Function

    '写Cookie
	
    Private Function SetC(ByRef url,ByRef str)
        SetC = HeadCookie(str)
		If SetC Then 
           key = "cookie/"&GetDomainKey(url)&".txt"
		   cache_.SetCache key,rCookie
		End If
    End Function
	
	'--------------------------------- Http Header Cookie
	
    '@GetCookie(): Cookie_对象 -> str标准字符串, 返回cookie字符串

    Private Function GetCookie()
	    Dim str
		For Each x in cookie_
            If str = "" Then
			   str = x&"="&cookie_(x)
			Else
			   str = str & "; " & x &"="&cookie_(x)
			End If
		Next
        GetCookie = str
    End Function
	
    '@SetCookie(ByRef str): str标准字符串 -> cookie_对象, 初始化cookie_
	
    Private Sub SetCookie(ByRef str)
		arr = Split(str,"; ")
        For i = 0 To UBound(arr)
		    k = Left(arr(i),InStr(arr(i),"=")-1)
			v = Mid(arr(i),InStr(arr(i),"=")+1,Len(arr(i))-InStr(arr(i),"="))
			cookie_(k)=v
        Next
    End Sub
	
    '@CleanCookie(): 清空cookie_

    Private Sub CleanCookie()
        cookie_.RemoveAll
    End Sub

    '@HeadCookie(ByRef str): Header字符串 => Cookie_对象, 更新cookie_
	
    Private Function HeadCookie(ByRef str)
	    Dim c,k,v,arr,arrr,i,s
	    HeadCookie = False
		Set c = MatchesExp(str,"Set-Cookie: ([^=]+)=([^;]+|);")
		For Each x in c
			HeadCookie = True
			k = x.SubMatches(0)
			v = x.SubMatches(1)
			If v = "" Then
			   If cookie_.Exists(k) Then cookie_.Remove(k)
			Else
			   If InStr(v,"=")>0 Then
			      s=""
			      arr=Split(v,"&")
				  For i = 0 To UBound(arr)
				      arrr=Split(arr(i),"=")
					  If UBound(arrr)=1 Then
					     If arrr(1)<>"" Then
						    If s = "" Then
						       s = arrr(0) &"="&arrr(1)
							Else
						       s = s & "&" &arrr(0) &"="&arrr(1)
							End If
						 End If
					  End If
				  Next
				  If s="" Then 
			         If cookie_.Exists(k) Then cookie_.Remove(k)
			      Else
				     cookie_(k)=s
				  End If
			   Else
			      cookie_(k)=v
			   End If
			End If
		Next
		Set c = Nothing
    End Function

	'查找字符串并返回集合
	
	Private Function MatchesExp(ByRef strng,ByRef patrn)
		Dim regEx
			Set regEx = New RegExp
			regEx.Pattern = patrn
			regEx.IgnoreCase = true
			regEx.Global = True
			Set MatchesExp = regEx.Execute(strng)
		set regEx=nothing
	End Function
	
	'取域名key,不准确凑合用
	
	Private Function GetDomainKey(ByRef url)
		Dim str, num
		str= Replace(url, "://", "")
		num = InStr(str, "/")
		If num > 0 Then
			GetDomainKey = Left(str, num -1)
		Else
			GetDomainKey = str
		End If
	End Function

    'Err
	
    Private Sub OutErr(ByRef str)
		Err.clear
        If IsDebug_ = true Then
            Response.charset = "utf-8"
            Response.Write str
            Response.End
        End If
    End Sub

End Class
%>