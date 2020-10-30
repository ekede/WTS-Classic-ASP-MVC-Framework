<%
'@title: Class_Function
'@author: ekede.com
'@date: 2018-07-22
'@description: 全局函数类

Class Class_Function

    Private Sub Class_Initialize()
    End Sub
    Private Sub Class_Terminate()
    End Sub
	
	'@HtmlEncodes(ByRef str): 字符串 - Html编码
	
	Public Function HtmlEncodes(ByRef str)
		Dim fString
		fString = Replace(str, ">", "&gt;")
		fString = Replace(fString, "<", "&lt;")
		'fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, Chr(34), "&quot;")
		fString = Replace(fString, Chr(39), "&#39;")
		fString = Replace(fString, Chr(13), "")
		fString = Replace(fString, Chr(10) & Chr(10), "</P><P>")
		fString = Replace(fString, Chr(10), "<BR>")
		HtmlEncodes = fString
	End Function
	
	'@HtmlDncodes(ByRef str): 字符串 - Html解码
	
	Public Function HtmlDecodes(ByRef str)
		Dim fString
		fString = Replace(str, "&gt;", ">")
		fString = Replace(fString, "&lt;", "<")
		'fString = Replace(fString,"&nbsp;",chr(32))
		fString = Replace(fString, "&quot;", Chr(34))
		fString = Replace(fString, "&#39;", Chr(39))
		fString = Replace(fString, "", Chr(13))
		fString = Replace(fString, "</P><P>", Chr(10) & Chr(10))
		fString = Replace(fString, "<BR>", Chr(10))
		HtmlDecodes = fString
	End Function
	
	'**************************************************
	'对String对象编码以便它们能在所有计算机上可读,所有空格、标点、重音符号以及其他非ASCII字符都用%xx编码代替其中xx等于表示该字符的十六进制数
	
	'@UrlEncodes(ByRef str): 字符串 - URL编码
	
	Public Function UrlEncodes(ByRef str)
		If str = "" Then Exit Function
		UrlEncodes = server.URLEncode(str)
	End Function
	
	'@UrlDecodes(ByVal str): 字符串 - URL解码

	Public Function UrlDecodes(ByVal str)
		Dim start,final,length,char,i,butf8,pass
		Dim leftstr,rightstr,finalstr
		Dim b0,b1,bx,blength,position,u,utf8
		On Error Resume Next
		
		b0 = Array(192,224,240,248,252,254)
		str = Replace(str,"+"," ")
		pass = 0
		utf8 = -1
		
		length = Len(str) : start = InStr(str,"%") : final = InStrRev(str,"%")
		If start = 0 Or length < 3 Then URLDecodes = str : Exit Function
		leftstr = Left(str,start - 1) : rightstr = Right(str,length - 2 - final)
		
		For i = start To final
		char = Mid(str,i,1)
		If char = "%" Then
		bx = URLDecode_Hex(Mid(str,i + 1,2))
		If bx > 31 And bx < 128 Then
		i = i + 2
		finalstr = finalstr & ChrW(bx)
		ElseIf bx > 127 Then
		i = i + 2
		If utf8 < 0 Then
		butf8 = 1 : blength = -1 : b1 = bx
		For position = 4 To 0 Step -1
		If b1 >= b0(position) And b1 < b0(position + 1) Then
		blength = position
		Exit For
		End If
		Next
		If blength > -1 Then
		For position = 0 To blength
		b1 = URLDecode_Hex(Mid(str,i + position * 3 + 2,2))
		If b1 < 128 Or b1 > 191 Then butf8 = 0 : Exit For
		Next
		Else
		butf8 = 0
		End If
		If butf8 = 1 And blength = 0 Then butf8 = -2
		If butf8 > -1 And utf8 = -2 Then i = start - 1 : finalstr = "" : pass = 1
		utf8 = butf8
		End If
		If pass = 0 Then
		If utf8 = 1 Then
		b1 = bx : u = 0 : blength = -1
		For position = 4 To 0 Step -1
		If b1 >= b0(position) And b1 < b0(position + 1) Then
		blength = position
		b1 = (b1 xOr b0(position)) * 64 ^ (position + 1)
		Exit For
		End If
		Next
		If blength > -1 Then
		For position = 0 To blength
		bx = URLDecode_Hex(Mid(str,i + 2,2)) : i = i + 3
		If bx < 128 Or bx > 191 Then u = 0 : Exit For
		u = u + (bx And 63) * 64 ^ (blength - position)
		Next
		If u > 0 Then finalstr = finalstr & ChrW(b1 + u)
		End If
		Else
		b1 = bx * &h100 : u = 0
		bx = URLDecode_Hex(Mid(str,i + 2,2))
		If bx > 0 Then
		u = b1 + bx
		i = i + 3
		Else
		If Left(str,1) = "%" Then
		u = b1 + Asc(Mid(str,i + 3,1))
		i = i + 2
		Else
		u = b1 + Asc(Mid(str,i + 1,1))
		i = i + 1
		End If
		End If
		finalstr = finalstr & Chr(u)
		End If
		Else
		pass = 0
		End If
		End If
		Else
		finalstr = finalstr & char
		End If
		Next
		URLDecodes = leftstr & finalstr & rightstr
	End Function
	'
	Function URLDecode_Hex(ByVal h)
		On Error Resume Next
		h = "&h" & Trim(h) : URLDecode_Hex = -1
		If Len(h) <> 4 Then Exit Function
		If isNumeric(h) Then URLDecode_Hex = cInt(h)
	End Function
	
	'**************************************************
	
	'@AddSlashes(ByVal str): 字符串 - 单引号转义
	
	Public Function AddSlashes(ByRef str)
		AddSlashes = Replace(str, "'", "''")
	End Function
	
	'**************************************************
	
	'@StrLength(ByRef str): 字符串 - 长度
	
	Public Function StrLength(ByRef str)
		On Error Resume Next
		Dim WINNT_CHINESE
		WINNT_CHINESE = (Len("中国") = 2)
		If WINNT_CHINESE Then
			Dim l, t, c
			Dim i
			l = Len(str)
			t = l
			For i = 1 To l
				c = Asc(Mid(str, i, 1))
				If c<0 Then c = c + 65536
				If c>255 Then
					t = t + 1
				End If
			Next
			StrLength = t
		Else
			StrLength = Len(str)
		End If
		If Err.Number<>0 Then Err.Clear
	End Function
	
	'@StrReplace(ByRef str,ByRef str1,ByRef str2): 字符串 - 替换,解决Null空值出错的问题

	Public Function StrReplace(ByRef str,ByRef str1,ByRef str2)
		If IsNull(str2) Or IsNull(str1) Then
			StrReplace = str
			Exit Function
		End If
		If InStr(Str, str1) = 0 Then
			StrReplace = str
			Exit Function
		End If
		StrReplace = Replace(str, str1, str2)
	End Function
	
	'@StrLeft(Byval str,ByRef strLeng,ByRef strExt): 字符串 - 截字符串

	Public Function StrLeft(Byval str,ByRef strLeng,ByRef strExt)
		str = ReplaceExp(str, "(<[^>]*>|\r\n|\r|\n)", "") '去掉html符号,换行符
		If strLeng > 0 Then Str = Left(str, strLeng)
		If strExt = "point" Then
			str = str&" ..."
		Else
			str = str&strExt
		End If
		StrLeft = str
	End Function
	
	'@StrBr(ByRef str): 字符串 - 替换换行符
	
	Public Function StrBr(ByRef str)
		StrBr = Replace(str, Chr(10), "<br/>")
		'strbr=replace(strbr,chr(13),"<br/>")
	End Function
	
	'@StripTags(str): 字符串 - 去除html标签和Asp标签
	
	Function StripTags(ByRef str) 
		StripTags = ReplaceExp(str,"<(.[^>]*)>","")
	End Function
	
	'**************************************************
	
	'@StrChars(ByRef str, ByVal chars): 字符串 - 搜索字符（字符串str是否由字符集chars组成）
	
	Public Function StrChars(ByRef str, ByVal chars)
		Dim i,u
		StrChars = True
		If  IsNull(str) Or IsEmpty(str) Then Exit Function
		If  chars = "" Then chars = "0123456789abcdefghijklmnopqrstuvwxyz-_/" '合法字符:给出正确模板 
		'
		For i = 1 To Len(str)
			u = Mid(str, i, 1)
			If InStr(chars, u) = 0 Then
			   StrChars = False
			   Exit For
			End If
		Next
	End Function
	
	'@StrCheck(ByRef str, ByVal chars): 字符串 - 搜索字符 (字符串str是否包含指定字符集chars）
	
	Public Function StrCheck(ByRef str, ByVal chars)
		Dim i, u
		StrCheck = False
		If  IsNull(str) Or IsEmpty(str) Then Exit Function
		If  chars = "" Then chars = "~@#$%^*(){}[]'\/<>?.:;,+!| " '非法字符:给出错误模板
		'
		For i = 1 To Len(chars)
			u = Mid(chars, i, 1)
			If InStr(str, u) > 0 Then
			   StrCheck = True
			   Exit For
			End If
		Next
	End Function
	
	'@StrEqual(ByRef str,ByRef strs,ByRef ge): 字符串 - 搜索单词
	
	Public Function StrEqual(ByRef str,ByRef strs,ByRef ge)
		Dim arr, i
		StrEqual = False
		If  IsNull(str) Or IsEmpty(str) Then Exit Function
        If  strs = "" Then Exit Function
        If  ge = "" Then ge = ","
        '
		arr = Split(strs, ge)
		For i = 0 To UBound(arr)
			If  arr(i) = str Then
				StrEqual = True
				Exit Function
			End If
		Next
	End Function
	
	'@StrCompare(ByRef a,ByRef t,ByRef b): 字符串 - 比较两个字符串的大小,区分大小写
	
	Public Function StrCompare(ByRef a,ByRef t,ByRef b)
		Dim isStr, b_comp
		isStr = False
		If VarType(a) = 8 Or VarType(b) = 8 Then
		    isStr = True
		    If IsNumeric(a) And IsNumeric(b) Then isStr = False
		    If IsDate(a) And IsDate(b) Then isStr = False
		End If
		If isStr Then
		    b_comp = StrComp(a,b,0)
		    Select Case LCase(t)
				Case "lt", "<" StrCompare = (b_comp = -1)
				Case "gt", ">" StrCompare = (b_comp = 1)
				Case "eq", "=" StrCompare = (b_comp = 0)
				Case "lte", "<=" StrCompare = (b_comp = -1 Or b_comp = 0)
				Case "gte", ">=" StrCompare = (b_comp = 1 Or b_comp = 0)
		    End Select
		Else
		    Select Case LCase(t)
				Case "lt", "<" StrCompare = (a < b)
				Case "gt", ">" StrCompare = (a > b)
				Case "eq", "=" StrCompare = (a = b)
				Case "lte", "<=" StrCompare = (a <= b)
				Case "gte", ">=" StrCompare = (a >= b)
		    End Select
		End If
	End Function
	
	'**************************************************
	
	'@TrimBoth(ByRef str1,ByRef str2): 字符串 - 去掉头尾指定字符
	
	Public Function TrimBoth(ByRef str1,ByRef str2)
		Dim str, strLeng
		str = TrimVBcrlf(str1) '去空白
		'
		strLeng = Len(str2)
		If strLeng<>0 Then
			If Left(str, strLeng) = str2 Then str = Right(str, Len(str) - strLeng)
			If Right(str, strLeng) = str2 Then str = Left(str, Len(str) - strLeng)
		End If
		TrimBoth = str
	End Function
	
	'@TrimVBcrlf(ByRef str): 字符串 - 去掉头尾连续空白字符
	
	Public Function TrimVBcrlf(ByRef str)
		TrimVBcrlf = RTrimVBcrlf(LTrimVBcrlf(str))
	End Function
	
	'@LTrimVBcrlf(ByRef str): 字符串 - 去掉开头空白字符
	
	Public Function LTrimVBcrlf(ByRef str)
		Dim pos, isBlankChar
		pos = 1
		isBlankChar = true
		While isBlankChar
			If Mid(str, pos, 1) = Chr(32) Then
				pos = pos + 1
			ElseIf Mid(str, pos, 1) = Chr(10) Then 'VBlf 换行
				pos = pos + 1
			ElseIf Mid(str, pos, 1) = Chr(13) Then 'VBcr 回车
				pos = pos + 1
		   'ElseIf Mid(str,pos,2)=VBcrlf then '换行+回车
		   'pos=pos+2
			Else
				isBlankChar = false
			End If
		Wend
		LTrimVBcrlf = Right(str, Len(str) - pos + 1)
	End Function
	
	'@RTrimVBcrlf(ByRef str): 字符串 - 去掉末尾空白字符
	
	Public Function RTrimVBcrlf(ByRef str)
		Dim pos, isBlankChar
		pos = Len(str)
		isBlankChar = true
		While isBlankChar And pos>= 1
			If Mid(str, pos, 1) = Chr(32) Then
				pos = pos -1
			ElseIf Mid(str, pos, 1) = Chr(10) Then
				pos = pos -1
			ElseIf Mid(str, pos, 1) = Chr(13) Then
				pos = pos -1
			   'ElseIf mid(str,pos-1,2)=VBcrlf then
			   'pos=pos-2
			Else
				isBlankChar = false
			End If
		Wend
		RTrimVBcrlf = RTrim(Left(str, pos))
	End Function

	'**************************************************

	'@GetRandomizeCode(ByRef num): 字符串 - 生成序列号
	
	Public Function GetRandomizeCode(ByRef num)
		Randomize
		Dim strRandArray, intRandlen, strRandomize, i
		strRandArray = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
		intRandlen = 16 ''定义随机码的长度
		If num<>"" Then intRandlen = num
		For i = 1 To intRandlen
			strRandomize = strRandomize & strRandArray(Int((21 * Rnd)))
		Next
		GetRandomizeCode = strRandomize
	End Function
	
	'@GetSequenceId(): 字符串 - 生成流水号
	
	Public Function GetSequenceId()
		Dim ranNum
		ranNum = Int(9 * Rnd) + 10
		GetSequenceId = Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&ranNum
	End Function
	
	'**************************************************
	
	'@GetDomain(ByRef url): 字符串 - 取域名
	
	Public Function GetDomain(ByRef url)
		Dim str, num
		If url = "" Then Exit Function
		str = url
		str= Replace(Str, "http://", "")
		str = Replace(Str, "https://", "")
		num = InStr(str, "/")
		If num > 0 Then
			GetDomain = Left(str, num -1)
		Else
			GetDomain = str
		End If
	End Function

	'@GetExt(ByRef fileName): 字符串 - 取扩展名
	
	Public Function GetExt(ByRef fileName)
		If InStr(fileName, ".") = 0 Then Exit Function
		GetExt = Mid(fileName, InStrRev(fileName, "."))
	End Function

	'**************************************************
	
	'@IIF(ByRef a,ByRef b,ByRef c): 字符串 - 三元运算 a？b：c
	
	Public Function IIF(ByRef a,ByRef b,ByRef c)
		If a Then IIF = b Else IIF = c
	End Function
	
	'**************************************************
	
	'@CheckExp(ByRef strng,ByRef patrn): 正则 - 查找字符串是否存在
	
	Public Function CheckExp(ByRef strng,ByRef patrn)
		Dim regEx
		Set regEx = New RegExp
			regEx.Pattern = patrn
			regEx.IgnoreCase = true
			regEx.Global = True
			CheckExp = regEx.Test(strng)
		Set regEx = Nothing
	End Function
	
	'@MatchesExp(ByRef strng,ByRef patrn): 正则 - 匹配字符串,返回匹配集合
	
	Public Function MatchesExp(ByRef strng,ByRef patrn)
		Dim regEx
			Set regEx = New RegExp
			regEx.Pattern = patrn
			regEx.IgnoreCase = true
			regEx.Global = True
			Set MatchesExp = regEx.Execute(strng)
		set regEx=nothing
	End Function
	
	'@ReplaceExp(ByRef strng,ByRef patrn,ByRef replaces): 正则 - 替换字符串
	
	Public Function ReplaceExp(ByRef strng,ByRef patrn,ByRef replaces)
		Dim regEx
		Set regEx = New RegExp
			regEx.Pattern = patrn
			regEx.IgnoreCase = True
			regEx.Global = True
			If regEx.Test(strng) Then
				ReplaceExp = regEx.Replace(strng, replaces)
			Else
				ReplaceExp = strng
			End If
		Set regEx = Nothing
	End Function
	
    '@IsValid(ByRef act,ByRef strng): 正则 - 验证更多 email,username,password,telephone,mobile,number,ip,date,zip

    Public Function IsValid(ByRef act,ByRef strng)
        Dim Pattern
        isvalid = False
        '
        Select Case act
            Case "email"
                Pattern = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
            Case "username"
                Pattern = "^[a-zA-Z][a-zA-Z0-9_]{3,15}$"
            Case "password"
                Pattern = "^[a-zA-Z0-9_]{4,15}$"
            Case "telephone"
                Pattern = "^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$"
            Case "mobile"
                Pattern = "^(13|15)[0-9]{9}$"
            Case "number"
                Pattern = "^[0-9]*[1-9][0-9]*$"
            Case "ip"
                Pattern = "^((25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)\.){3}(25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)$"
            Case "date" '0000-00-00
                Pattern = "^((\d{2}(([02468][048])|([13579][26]))[\-\/\s]?((((0?[13578])|(1[02]))[\-\/\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\-\/\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\-\/\s]?((0?[1-9])|([1-2][0-9])))))|(\d{2}(([02468][1235679])|([13579][01345789]))[\-\/\s]?((((0?[13578])|(1[02]))[\-\/\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\-\/\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\-\/\s]?((0?[1-9])|(1[0-9])|(2[0-8]))))))(\s(((0?[1-9])|(1[0-2]))\:([0-5][0-9])((\s)|(\:([0-5][0-9])\s))([AM|PM|am|pm]{2,2})))?$"
            Case "zip"
                Pattern = "^\d{6}$"
            Case Else
                Exit Function
        End Select
        IsValid = CheckExp(strng,Pattern) '外部函数
    End Function
	
	'**************************************************
	
	'1E+20  表示100000000000000000000(即1后面20个零)
	'CInt   整形 Integer  -32,768 至 32,767，小数部分四舍五入
	'CLng   长整形 Long     -2,147,483,648 至 2,147,483,647，小数部分四舍五入
	'CSng() 单精度 精度7位，负值 -1.18E - 38 到 -3.40E + 38，正值 1.18E - 38 到 3.40E + 38，也可以取 0。
	'CDbl() 双精度 精度15位。负值 - 2.23E - 308 到 -1.79E + 308，正值 2.23E - 308 到 1.79E + 308，也可以为 0。
	'CCur() 货币 -922,337,203,685,477.5808 至922,337,203,685,477.5807
	
	'@StrClng(ByRef str): 数字 - 将数字串转为长整型数值,Fix向零方向取整
	
	Public Function StrClng(ByRef str)
		If IsNumeric(str) Then
			StrClng = Fix(CDbl(str))
		Else
			StrClng = 0
		End If
	End Function
	
	'@StrClngf(ByRef str,ByRef n): 数字 - 将数字串转为长整型数值,Round四舍无入
	
	Public Function StrClngf(ByRef str,ByRef n)
		If IsNumeric(Str) Then
			StrClngf = round(CDbl(str), n)
		Else
			StrClngf = 0
		End If
	End Function
	
	'@FormatNum(ByVal num,ByRef n): 数字 - 取小数点后几位,小于1整数位自动添零
	
	Public Function FormatNum(ByVal num,ByRef n)
		If num<1 Then
			num = "0"&CStr(FormatNumber(num, n))
		Else
			num = CStr(FormatNumber(num, n))
		End If
		FormatNum = Replace(num, ",", "")
	End Function
	
	'**************************************************
	
    '@FormatDate(ByRef dateAndTime,ByRef para): 日期 - 格式化

    Public Function FormatDate(ByRef dateAndTime,ByRef para)
        Dim y, m, d, h, mi, s, strDateTime
        FormatDate = dateAndTime
        If Not IsNumeric(para) Then Exit Function
        If Not IsDate(dateAndTime) Then Exit Function
        '
        y = CStr(Year(dateAndTime))
        m = CStr(Month(dateAndTime))
        If Len(m) = 1 Then m = "0" & m
        d = CStr(Day(dateAndTime))
        If Len(d) = 1 Then d = "0" & d
        h = CStr(Hour(dateAndTime))
        If Len(h) = 1 Then h = "0" & h
        mi = CStr(Minute(dateAndTime))
        If Len(mi) = 1 Then mi = "0" & mi
        s = CStr(Second(dateAndTime))
        If Len(s) = 1 Then s = "0" & s
        '
        Select Case para
            Case "1"
                strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
            Case "2"
                strDateTime = y & "-" & m & "-" & d
            Case "3"
                strDateTime = y & "/" & m & "/" & d
            Case "4"
                strDateTime = m & "-" & d & " " & h & ":" & mi
            Case "5"
                strDateTime = Right(y, 2) & "-" &m & "-" & d
            Case "6"
                strDateTime = m & "-" & d
            Case Else
                strDateTime = dateAndTime
        End Select
        FormatDate = strDateTime
    End Function

	'**************************************************
	
	'@OpenURL(ByRef act,ByRef url,ByRef param,ByRef res): 文档 - 读取远端文件 Ajax,Get,Post
	
	Public Function OpenURL(ByRef act,ByRef url,ByRef param,ByRef res)
		On Error Resume Next
		Dim http
		Set http = server.CreateObject("Msxml2.ServerXMLHTTP")
		With http
			If act = "POST" Then
				.Open "POST", url, false , "" , ""
				.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				.Send(param)
			Else
				If param = "" Then
					.Open "GET", url, false
				Else
					.Open "GET", url&"?"&param, false
				End If
				.setRequestHeader "If-Modified-Since", "0"
				.send()
			End If
			'
			If .Readystate = 4 Then
				If .Status = 200 Then
					Select Case res
						Case "head"
							OpenURL = .getAllResponseHeaders
						Case "body"
							OpenURL = .ResponseBody '二进制
						Case "xml"
							OpenURL = .ResponseXML
						Case Else
							OpenURL = .ResponseText
					End Select
				End If
			End If
		End With
		Set http = Nothing
	End Function
	
    '@DownLoad(path): 文档 - 读服务器文件，缓存流输出到浏览器

    Public Function DownLoad(ByRef path)
        Dim fso, f, strFileName, intFileLength
		Download = False
		'
        Set fso = server.CreateObject("scripting.filesystemobject")
        If Not fso.FileExists(path) Then
            Exit Function
        Else
            Set f = fso.GetFile(path)
            intFileLength = f.Size '获取文件大小
            strFileName = f.Name
            Set f = Nothing
        End If
        Set fso = Nothing
        '
        response.Clear
        response.buffer = true
        response.addheader "content-disposition", "attachment;filename=" & strfilename
        response.addheader "content-length" , intfilelength
        response.contenttype = "application/octet-stream" 
	   'response.addheader "content-type","application/x-msdownload"
		'
        Dim s
        Set s = server.CreateObject("adodb.stream")
			s.Open
			s.Type = 1
			s.LoadFromFile(path)
			While Not s.eos
			   response.binarywrite s.Read(1024 * 64)
			   response.flush
			Wend
			s.Close
        Set s = Nothing
        '
        Download = True
    End Function
	
	'@XmlTrans(ByRef strXml,ByRef strXsl): 文档 - 合并XML XSLT
	
	Public Function XmlTrans(ByRef strXml,ByVal strXsl)
		On Error Resume Next
		dim xml,xsl
		if  strXml = "" or strXsl = "" Then Exit Function
		'
		set xml = Server.CreateObject("Msxml2.DOMDocument.6.0")
		xml.async = false
		xml.loadxml(strXml)
		'
		set xsl = Server.CreateObject("Msxml2.DOMDocument.6.0")
		xsl.async = false
		xsl.loadxml(strXsl)
		'
		XmlTrans = xml.TransformNode(xsl)
	   'XmlTrans = Replace(XmlTrans,"><",">"&chr(10)&"<")
	    Set xsl = Nothing
	    Set xml = Nothing
	End Function
	
	'**************************************************
	
	'@IsObjInstalled(ByRef strClass): 检查组件是否已经安装
	Function IsObjInstalled(ByRef strClass)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClass)
		If 0 = Err Then 
		   IsObjInstalled = True
		   Set xTestObj = Nothing
		   Err = 0
		Else
		   Err.Clear
		End If
	End Function
	
	'@IsEmptys(ByRef arg): 调试 - 判断是否为空,包括数值型和引用型
	
	Function IsEmptys(ByRef arg)	
		if isArray(arg) Then
			if ubound(arg)<0 Then IsEmptys = true : exit Function
		end if
		if isNumeric(arg) Then
			if arg = 0 Then IsEmptys = true : exit Function
		End If
		IsEmptys = false
		Select Case Typename(arg)
			Case "Empty","Null","Nothing"
				IsEmptys = true
			Case "Dictionary","IVariantDictionary","IRequestDictionary"
				IsEmptys = (0 = arg.count)
			Case "Recordset"
				IsEmptys = (0 = arg.RecordCount)
			Case "ISessionObject"
				IsEmptys = (0 = arg.Contents.Count)
			Case "String"
				IsEmptys = ("" = arg)
			Case "Boolean"
				IsEmptys = (false = arg)
			Case else
				'其它情况还未发现 
		End Select
	End Function
	
    '获取打印集合字符串

    Private Function ToStr(ByRef sets)
	    On Error Resume Next
        Dim x, y ,na ,str
		t=Typename(sets)
		Select Case t
		Case "Byte","Integer","Long","Double","Single","Currency","Decimal","Boolean","Date","String" 'String
		    str=str & sets & chr(10)
		Case "IRequestDictionary" 'Request.Cookies
			For Each x in sets
				If sets(x).HasKeys Then
				    IF Err Then 'Requst.Form , Request.ServerVariables
					   Err.clear
					   str=str & x & " => " & sets(x) & chr(10)
					Else
					   For Each y in sets(x)
						   str=str & x & " => " & y & " => " & sets(x)(y) &chr(10)
					   Next
					End If
				Else
					str=str & x & " => " & sets(x) & chr(10)
				End If
			Next
	    Case "IStringList"
		    str=str&sets
		Case "Variant()" 'array
		    For x = 0 to Ubound(sets)
				str=str & x & " => " & sets(x) & chr(10)
			Next
		Case  "Dictionary","IVariantDictionary" 'Dictionary, Session.Contents
			For Each x in sets
			    If TypeName(sets(x)) = "String" or TypeName(sets(x)) = "Integer"  Then
				   str=str &  x & " => " & sets(x) & chr(10)
				Else
				   str=str &  x & " => " & TypeName(sets(x)) & chr(10)
				End If
			Next
		Case  "IMatchCollection2"  'Matches
			For Each x in sets
			    str=str &  x & chr(10)
				For Each y in x.SubMatches
					str=str &  " => " & y & chr(10)
				Next
			Next
		Case "Byte()"
			For i = 0 to ubound(sets)
			    b = MidB(sets,i+1,1)
			    t = AscB(b)
				Select Case t
				   Case 9
				      t1 = "vbTab" '制表符TAB按键
				   Case 10
				      t1 = "vbLF" '换行\n 在新的一行光标位置不变
				   Case 13
				      t1 = "vbCR" '回车\r 将光标移动到所在行的开始
				   Case 0
				      t1 = "Null" '空字符
				   Case Else
				      t1 = chrW(t)
				End Select
				str=str & i& " => "& t &" => "& Hex(t) & " => " & t1 &Chr(10)
			Next
		Case Else
		    str=str &  t & chr(10)
		End Select
		ToStr = str
	End Function
	
	'@Print(ByRef str): 调试 - 输出变量
	
	Public Sub Print(ByRef str)
		response.charset = "utf-8"
		response.write ToStr(str)
	End Sub

End Class
%>