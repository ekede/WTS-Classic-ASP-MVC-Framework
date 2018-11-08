<%
'@title: Class_Template
'@author: ekede.com
'@date: 2018-08-11
'@description: 模板类

Class Class_Template
    '
    Private isDebug_
    Private regEx_
    Private tempdata_
    Private loader_
    Private path_tpl_
    Private pathD_tpl_
	
    '@path_tpl: 模板根地址

    Public Property Let path_tpl(Value)
        path_tpl_ = LCase(Value)
    End Property
	
    '@pathD_tpl: 模板默认根地址
	
    Public Property Let pathD_tpl(Value)
        pathD_tpl_ = LCase(Value)
    End Property
	
    '@loader: loader对象依赖

    Public Property Let loader(Value)
        Set loader_ = Value
    End Property
	
    '@tempdata: 模板标签存放字典

    Public Property Let tempdata(Value) '替换字典
        If VarType(Value) = 9 Then Set tempdata_ = Value
    End Property

    Private Sub Class_Initialize()
        If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
		'
        Set regEx_ = New RegExp
        regEx_.IgnoreCase = True
        regEx_.Global = True
        '
        Set tempdata_ = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
        Set tempdata_ = Nothing
        Set regEx_ = Nothing
    End Sub

    '@SetVal(keys, values): set value

    Public Function SetVal(keys, values)
        Dim tempArray, k
        SetVal = true
        If InStr(keys, "/")>0 Then
            tempArray = Split(keys, "/")
            If tempArray(0) = "" Then
                SetVal = false
                Exit Function
            Else
                k = tempArray(0)
                If tempdata_.Exists(k) Then
                    If IsNumeric(tempdata_(k)) Then
                        tempdata_(keys&"/"&tempdata_(k)) = values
                    Else
                        tempdata_(k) = 0
                        tempdata_(keys&"/"&tempdata_(k)) = values
                    End If
                Else
                    tempdata_(k) = 0
                    tempdata_(keys&"/"&tempdata_(k)) = values
                End If
            End If
        Else
            tempdata_(keys) = values
        End If
    End Function

    '@GetVal(keys): get value

    Public Function GetVal(keys)
        GetVal = tempdata_(keys)
    End Function

    '@SetVali(keys, i, values): set value table

    Public Function SetVali(keys, i, values)
        If i>= 0 Then
            tempdata_(keys&"/"&i) = values
            SetVali = true
        Else
            SetVali = False
        End If
    End Function

    '@GetVali(keys, i): get value table

    Public Function GetVali(keys, i)
        GetVali = tempdata_(keys&"/"&i)
    End Function

    '@UpdVal(keys): update value

    Public Function UpdVal(keys)
        UpdVal = true
        If tempdata_.Exists(keys) Then
            If IsNumeric(tempdata_(keys)) Then
                tempdata_(keys) = tempdata_(keys) + 1
            Else
                UpdVal = false
            End If
        Else
            UpdVal = false
        End If
    End Function
	
    '@Fetch(mb_name): Fetch Template

    Public Function Fetch(mb_name)
        Dim str
        str = ReadTpl(mb_name)
        str = ReplaceIF(str)
        str = ReplaceExt(str)
        str = ReplaceLoop(str)
        str = ReplaceTag(str, 0)
        str = ReplaceArray(str)
        str = ReplaceSpaceLine(str)
        Fetch = str
    End Function

    'Read Tpl

    Private Function ReadTpl(ByVal mb_name)
        On Error Resume Next
        Dim str
        str = loader_.LoadFile(path_tpl_&mb_name)
        If str = -1 Then '加载默认模板
		    If  pathD_tpl <> "" And path_tpl_<> pathD_tpl_ Then
				str = loader_.LoadFile(pathD_tpl_&mb_name)
				If str = -1 Then
					Exit Function
				Else
					ReadTpl = ReplaceInclude(Str)
				End If
			Else
			    Exit Function
			End If
        Else
            ReadTpl = ReplaceInclude(Str)
        End If
		If Err Then OutErr("Template:ReadTpl "&Err.Description)
    End Function

    'ReplaceInclude 合并模板文件

    Private Function ReplaceInclude(ByVal str)
        On Error Resume Next
        Dim matches,m,tempValue,tempStr
        If str = "" Then Exit Function
        '
        regEx_.Pattern = "{include ([\.\w_/]*)}" '匹配出<!--XXX-->
        Set matches = regEx_.Execute(str)
        For Each m In matches
            tempStr = m.SubMatches(0)
            tempValue = loader_.LoadFile(path_tpl_&tempStr)
			If tempValue<>"" And tempValue<> -1 Then 
			   str=Replace(str,m.value,tempValue)
			Else '加载默认模板
			   If  pathD_tpl <> "" And path_tpl_<> pathD_tpl_ Then
			       tempValue = loader_.LoadFile(pathD_tpl_&tempStr)
				   If tempValue<>"" And tempValue<> -1 Then str=Replace(str,m.value,tempValue)
			   End If
			End If
        Next
		Set matches = Nothing
        '
        ReplaceInclude = str
		If Err Then OutErr("Template:ReplaceInclude "&Err.Description)
    End Function
	
    'ReplaceExt 替换模块

    Private Function ReplaceExt(ByVal str)
        On Error Resume Next
        Dim matches,m,tempValue,tempStr,tempArray
        If str = "" Then Exit Function
        '
        regEx_.Pattern = "{ext ([:\w_/]*)}" '匹配出<!--XXX-->
        Set matches = regEx_.Execute(str)
        For Each m In matches
            tempStr = m.SubMatches(0)
			tempArray = Split(tempStr, ":")
			If Ubound(tempArray)=1 Then
               tempValue = loader_.LoadControlTag(tempArray(0),tempArray(1))
			Else
               tempValue = ""
			End If
			str=Replace(str,m.value,tempValue)
        Next
		Set matches = Nothing
        '
        ReplaceExt = str
		If Err Then OutErr("Template:ReplaceExt "&Err.Description)
    End Function

    'ReplaceIF 替换IF块,不能嵌套

    Private Function ReplaceIF(ByVal str)
        On Error Resume Next
		Dim matches,m,tempStr,tempValue,tempArray,thenValue
        If str = "" Then Exit Function
        '
        regEx_.Pattern = "{if ([\w_=]*)}([\s\S]*?){end if}" '匹配:<!--if XXX-->YYY<!--end if-->
        Set matches = regEx_.Execute(str)
        For Each m In matches
            tempStr = m.SubMatches(0)
			thenValue = m.SubMatches(1)
            If InStr(tempStr, "=")>0 Then
                tempArray = Split(tempStr, "=")
                If CStr(tempdata_(tempArray(0))) = CStr(tempArray(1)) Then
                    tempValue = "ok"
                Else
                    tempValue = ""
                End If
            Else
                tempValue = tempdata_(tempStr)
                If VarType(tempValue)>8000 Then tempValue = "ok"
            End If
            '
            If InStr(thenValue, "{else}")>0 Then
			    tempArray=split(thenValue,"{else}")
                If tempValue = "" Or IsNull(tempValue) Then
					str= Replace(str,m.value,tempArray(1))
                Else
					str= Replace(str,m.value,tempArray(0))
                End If
            Else
                If tempValue = "" Or IsNull(tempValue) Then
				    str= Replace(str,m.value,"")
                Else
                    str= Replace(str,m.value,thenValue)
                End If
            End If
        Next
		Set matches = Nothing
        '
        ReplaceIF = str
		If Err Then OutErr("Template:ReplaceIF "&Err.Description)
    End Function
	
    'ReplaceLoopIF 替换LoopIF块,不能嵌套

    Private Function ReplaceLoopIF(ByVal str,ByVal i)
        On Error Resume Next
		Dim matches,m,tempStr,tempValue,tempArray,thenValue
        If str = "" Then Exit Function
        '
        regEx_.Pattern = "{if loop ([\w_=/]*)}([\s\S]*?){end loop if}" '匹配:<!--if XXX-->YYY<!--end if-->
        Set matches = regEx_.Execute(str)
        For Each m In matches
            tempStr = m.SubMatches(0)
			thenValue = m.SubMatches(1)
            If InStr(tempStr, "=")>0 Then
                tempArray = Split(tempStr, "=")
				If InStr(tempArray(0), "/")>0 Then
					If CStr(tempdata_(tempArray(0)&"/"&i)) = CStr(tempArray(1)) Then
						tempValue = "ok"
					Else
						tempValue = ""
					End If
                Else
					If CStr(tempdata_(tempArray(0))) = CStr(tempArray(1)) Then
						tempValue = "ok"
					Else
						tempValue = ""
					End If
                End If
            Else
                tempValue = tempdata_(tempStr)
                If VarType(tempValue)>8000 Then tempValue = "ok"
            End If
            '
            If InStr(thenValue, "{else loop}")>0 Then
			    tempArray=split(thenValue,"{else loop}")
                If tempValue = "" Or IsNull(tempValue) Then
					str= Replace(str,m.value,tempArray(1))
                Else
					str= Replace(str,m.value,tempArray(0))
                End If
            Else
                If tempValue = "" Or IsNull(tempValue) Then
				    str= Replace(str,m.value,"")
                Else
                    str= Replace(str,m.value,thenValue)
                End If
            End If
        Next
		Set matches = Nothing
        '
        ReplaceLoopIF = str
		If Err Then OutErr("Template:ReplaceLoopIF "&Err.Description)
    End Function

    'ReplaceLoop 替换循环

    Private Function ReplaceLoop(ByVal str)
	    On Error Resume Next
		Dim matches,m,loopName,tempHeader,tempSquare,tempFooter,tempStr,i,tempBlock
        If str = "" Then Exit Function
        '
        regEx_.Pattern = "{loop ([\w_]*)}([\s\S]*?){loop_body start}([\s\S]*?){loop_body end}([\s\S]*?){end loop}"
        Set matches = regEx_.Execute(str)
        For Each m In matches
            '
            loopName = m.SubMatches(0)
			If  tempdata_.Exists(LoopName) and IsNumeric(tempdata_(LoopName)) Then
                tempHeader = m.SubMatches(1)
                tempSquare = m.SubMatches(2)
                tempFooter = m.SubMatches(3)
                tempStr = ""
                For i = 0 To tempdata_(loopName) -1
                    tempStr = tempStr&ReplaceTag(ReplaceLoopIF(tempSquare, i), i) '替换标签
                Next
                tempBlock = tempHeader&tempStr&tempFooter
            Else
                tempBlock = ""
            End If
			str=Replace(str,m.value,tempBlock)
        Next
		Set matches = Nothing
        '
        ReplaceLoop = str
		If Err Then OutErr("Template:ReplaceLoop "&Err.Description)
    End Function

    'ReplaceTag 替换语言包,字典

    Private Function ReplaceTag(ByVal str, ByVal i)
        On Error Resume Next
		Dim matches,m,tempStr,tempValue
        If str = "" Then Exit Function
        '
        regEx_.Pattern = "{([\w_/:]*)}" '匹配出<!--XXX-->
        Set matches = regEx_.Execute(str)
        For Each m In matches
            tempStr = m.SubMatches(0)
            If InStr(tempStr, ":")>0 Then
                tempArray = Split(tempStr, ":")
                If InStr(tempArray(0), "/")>0 Then
                    tempValue = FormatTag(tempdata_(tempArray(0)&"/"&i), tempArray)
                Else
                    tempValue = FormatTag(tempdata_(tempArray(0)), tempArray)
                End If
            Else
                If InStr(tempStr, "/")>0 Then
                    tempValue = tempdata_(tempStr&"/"&i)
                Else
                    tempValue = tempdata_(tempStr)
                End If
            End If
			str=Replace(str,m.value,tempValue)
        Next
		Set matches = Nothing
        '
        ReplaceTag = str
		If Err Then OutErr("Template:ReplaceTag "&Err.Description)
    End Function

    'ReplaceArray 数组值

    Private Function ReplaceArray(ByVal str)
        On Error Resume Next
		Dim matches,m,tempArray,tempI,tempFormat,tempValue
        If str = "" Then Exit Function
        '
        regEx_.Pattern = "{([\w_]*)\.([0-9]+)(|:[\w_:]*)}" '匹配出<!--XXX-->
        Set matches = regEx_.Execute(str)
        For Each m In matches
            tempArray = m.SubMatches(0)
            tempI = m.SubMatches(1)
            tempFormat = m.SubMatches(2)
            '
            tempValue = ""
            If VarType(tempdata_(tempArray))>8000 Then
                tempNum = UBound(Temp_data(tempArray))
                If CInt(tempI)>= 0 And CInt(tempI)<= tempNum Then
                    tempValue = tempdata_(tempArray)(tempI)
                    If tempValue<>"" And tempFormat<>"" Then tempValue = FormatTag(tempValue, Split(tempFormat, ":")) '格式化标签
                End If
            End If
            '
			If IsNull(tempValue) = False Then str=Replace(str,m.value,tempValue)
        Next
		Set matches = Nothing
        '
        ReplaceArray = str
		If Err Then OutErr("Template:ReplaceArray "&Err.Description)
    End Function
	
    'ReplaceSpaceLine 替换空行,美化代码 

    Private Function ReplaceSpaceLine(ByVal str)
         If str = "" Then Exit Function
		 regEx_.Pattern = "\n[ ]{0,}\r"
		 ReplaceSpaceLine = regEx_.Replace(str, "")
    End Function

    'FormatTag 格式化标签

    Private Function FormatTag(ByVal tag, ByVal arr)
        On Error Resume Next
        Select Case arr(1)
            Case "date"
                FormatTag = wts.fun.FormatDate(tag, arr(2))
            Case "str"
                FormatTag = wts.fun.strleft(tag, arr(2), arr(3))
            Case "br"
                FormatTag = wts.fun.strbr(tag)
            Case "strip"
                FormatTag = wts.fun.StripTags(tag)
            Case Else
                FormatTag = tag
        End Select
		If Err Then OutErr("Template:FormatTag "&Err.Description)
    End Function
	
	'错误提示

	Public Sub OutErr(ErrMsg)
		If isDebug_ = true Then
			Response.charset = "utf-8"
			Response.Write ErrMsg
			Response.End
		End If
	End Sub

End Class
%>