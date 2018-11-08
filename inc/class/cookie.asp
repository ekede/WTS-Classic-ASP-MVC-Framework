<%
'@title: Class_Cookie
'@author: ekede.com
'@date: 2017-12-6
'@description: Cookie对象

Class Class_Cookie

    Private path_
	Private domain_
	Private expire_
	Private encode_
	
	'@path: cookie路径
	
    Public Property Get path
        Set path = path_
    End Property
	
	'@domain: cookie域名

    Public Property Let domain(Value)
        Set domain_ = Value
    End Property
	
	'@expire: cookie过期时间
	
    Public Property Let expire(Value)
        Set expire_ = Value
    End Property

	'@encode: cookie加密
	
    Public Property Let encode(Value)
        encode_ = Value
    End Property

    Private Sub Class_Initialize()
	   'path_=Left(Request.ServerVariables("script_name"),inStrRev(Request.ServerVariables("script_name"),"/"))
	   'expire_=Date+1
	    encode_=False
    End Sub
    Private Sub class_terminate()
    End Sub

	'--------------------------------- Server cookie
	
    '@GetC(k1,k2): 读

    Public Function GetC(k1,k2)
	    Dim v
	    If k2  = "" Then
           v = Request.Cookies(k1)
		Else
           v = Request.Cookies(k1)(k2)
		End If
		If encode_ = True Then
           GetC = DecodeC(v)
		Else
           GetC = v
		End If
    End Function
	
    '@SetC(k1,k2,v,d,p,e): 写 -key1,key2,Value,Domain,Path,Expires

    Public Sub SetC(k1,k2,ByVal v,ByVal d,ByVal p,ByVal e)
		If encode_ = True Then v = EncodeC(v)
		'
	    If  k2 = "" Then
			Response.Cookies(k1) = v
		Else
			Response.Cookies(k1)(k2) = v
		End If
		'
        If d="" And domain_ <> "" Then d = domain_
		If d<>"" Then Response.Cookies(k1).Domain= d
        '
		If p="" And path_ <> "" Then p = path_
		If p<>"" Then Response.Cookies(k1).Path= p
		'
		If e="" And expire_ <> "" Then e = expire_
		If e<>"" Then Response.Cookies(k1).Expires = e
    End Sub

    '@DelC(k1,k2,d,p): 删

    Public Sub DelC(k1,k2,d,p)
	     If k2 <> "" Then
            SetC k1,k2,"",d,p,""
	     Else
            SetC k1,"","",d,p,(Now()-1)
		 End If
    End Sub

    '@CleanC(d,p): 清

    Public Sub CleanC(d,p)
        For Each k In Request.Cookies
		    DelC k,"",d,p
        Next
    End Sub

	'编码cookies, 编码处理后的信息，字符以"a"隔开

	Private Function EncodeC(contentStr)
		Dim i,returnStr
		For i = Len(contentStr) to 1 Step -1
			returnStr = returnStr & Ascw(Mid(contentStr,i,1))
			If (i <> 1) Then returnStr = returnStr & "a"
		Next
		EncodeC = returnStr
	End Function

	'解码cookies ,解码处理后的信息

	Private Function DecodeC(contentStr)
		Dim i
		Dim StrArr,StrRtn
		StrArr = Split(contentStr,"a")
		For i = 0 to UBound(StrArr)
			If isNumeric(StrArr(i)) = True Then
			    StrRtn = Chrw(StrArr(i)) & StrRtn
			Else
				StrRtn = contentStr
				Exit Function
			End If
		Next
		DecodeC = StrRtn
	End Function

End Class
%>