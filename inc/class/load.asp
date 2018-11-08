<%
'@title: Class_Load
'@author: ekede.com
'@date: 2018-08-05
'@description: 动态加载,包含,实例化对象

Class Class_Load

    Private isDebug_
    Private frameworkPath_
    Private classPath_
	Private controlPath_
	Private modelPath_
	Private languagePath_
	Private languageDefaultPath_ '默认语言包地址
	Private templatePath_
	Private templateDefaultPath_ '默认模板地址
	Private includeCount_

	'@isDebug: 开启调试

	Public Property Let isDebug(Value)
		isDebug_ = Value
	End Property
	
	'@frameworkPath: 框架根路径

	Public Property Let frameworkPath(Value)
		frameworkPath_ = Value
	End Property
	
	'@classPath: 类库根路径

	Public Property Let classPath(Value)
		classPath_ = Value
	End Property
	
	'@modelPath: 模型根路径

	Public Property Let modelPath(Value)
		modelPath_ = Value
	End Property
	
	'@controlPath: 控制器根路径

	Public Property Let controlPath(Value)
		controlPath_ = Value
	End Property
	
	'@languageDefaultPath: 语言包默认根路径
	
	Public Property Let languageDefaultPath(Value)
		languageDefaultPath_ = Value
	End Property
	
	'@languagePath: 语言包根路径

	Public Property Let languagePath(Value)
		languagePath_ = Value
	End Property
	
	'@templateDefaultPath: 模板默认根路径

	Public Property Let templateDefaultPath(Value)
		templateDefaultPath_ = Value
	End Property
	
	'@templatePath: 模板根路径

	Public Property Let templatePath(Value)
		templatePath_ = Value
	End Property

	Private Sub Class_Initialize()
        If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
		'
		Set includeCount_ = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
		Set includeCount_ = nothing
	End Sub
	
	'@LoadFile(fileName): 加载文件并缓存

	Public Function LoadFile(fileName)
		Dim str, rpath
		'
		If PATH_APP <> "" Then
		   If Left(fileName,Len(PATH_INC))= PATH_INC Then
		      rpath_m = GetMapPath(PATH_ROOT&PATH_APP&Right(fileName,Len(fileName)-Len(PATH_INC)))
		   Else
		      rpath_m = GetMapPath(PATH_ROOT&PATH_APP&FileName)
		   End If
		Else
		   rpath_m = -1
		End If
		rpath = GetMapPath(PATH_ROOT&FileName)
		'
		If rpath_m <> -1 And Application("file_"&rpath_m)<>"" Then
			str = Application("file_"&rpath_m) 'Modify
		ElseIf Application("file_"&rpath)<>"" Then
			str = Application("file_"&rpath)   'Original
		Else
			'Modify
			If rpath_m = -1 Then
			   str = -1
			Else
				str = ReadUTF(rpath_m)
				If str <> -1 Then
					If Right(fileName, 4) = ".asp" Then str = IncludeReplace(str)
					If isDebug_ = false Then Application("file_"&rpath_m) = str '调试环境,不缓存
				End If
			End If
			'Original
			If str <> - 1 Then
			ElseIf rpath = -1 Then
				str = -1
			Else
				str = ReadUTF(rpath)
				If str<> -1 Then
					If Right(fileName, 4) = ".asp" Then str = IncludeReplace(str)
					If isDebug_ = false Then Application("file_"&rpath) = str '调试环境,不缓存
				End If
			End If
		End If
		LoadFile = str
	End Function
	
	'Include 代码块动态包含,包含文件路径-永远是相对根目录路径 PATH_ROOT:回根,filePath相对根目录+文件名
	
	'@Include(filePath): 包含并执行,变量全局, 函数,类局部

	Public Sub Include(filePath)
		On Error Resume Next
		Dim str
		str = LoadFile(filePath&".asp")
		If str<> -1 Then Execute str  '全局 变量; 类,函数,根据位置全局局部; 类相同位置出错; 函数覆盖静态;
		If Err Then OutErr("Include:"&filePath&":"&Err.Number&":"&Err.Description)
	End Sub
	
	'@IncludeG(filePath): 包含并执行,全局
	
	Public Sub IncludeG(filePath)
		On Error Resume Next
		Dim str
		If includeCount_(filePath)<>1 Then '避免全局类包含两次以上出错
			includeCount_(filePath)=1
			str = LoadFile(filePath&".asp")
			If str<> -1 Then ExecuteGlobal str  '全局 类,函数,变量
			If Err Then OutErr("IncludeG:"&filePath&":"&Err.Number&":"&Err.Description)
		End If
	End Sub
	
	'@IncludeS(filePath): 包含不执行,返回内容
	
	Public Function IncludeS(filePath)
		On Error Resume Next
		Dim str
		str = LoadFile(filePath&".asp")
		If str<> -1 Then IncludeS = str  '返回文件字符串
		If Err Then OutErr("IncludeS:"&filePath&":"&Err.Number&":"&Err.Description)
	End Function
	
	'@LoadFrameWork(filePath):加载框架对象

	Public Function LoadFramework(filePath)
		On Error Resume Next
		Dim Str, cname
		Str = LoadFile(frameworkPath_&filePath&".asp")
		If Str <> -1 Then
			Execute Str '使类局部,避免多次包含冲突
			Set LoadFramework = Eval("new framework_"&filePath)
		End If
		If Err Then OutErr("LoadFramework:"&filePath&":"&Err.Number&":"&Err.Description)
	End Function

	'@LoadClass(filePath):加载类库对象

	Public Function LoadClass(filePath)
		On Error Resume Next
		Dim Str, cname
		Str = LoadFile(classPath_&filePath&".asp")
		If Str <> -1 Then
			Execute Str '使类局部,避免多次包含冲突
			If InStr(filePath, "/")>0 Then
				cname = Replace(filePath, "/", "_")
			Else
				cname = filePath
			End If
			Set LoadClass = Eval("new class_"&cname)
		End If
		If Err Then OutErr("LoadClass:"&filePath&":"&Err.Number&":"&Err.Description)
	End Function

	'@LoadControl(filePath): 加载控制器对象

	Public Function LoadControl(filePath)
		On Error Resume Next
		Dim Str, cname
		Str = LoadFile(controlPath_&filePath&".asp")
		If Str <> -1 Then
			Execute Str
			If InStr(filePath, "/")>0 Then
				cname = Replace(filePath, "/", "_")
			Else
				cname = filePath
			End If
			Set LoadControl = Eval("new control_"&cname)
		End If
		If Err Then OutErr("LoadControl:"&filePath&":"&Err.Number&":"&Err.Description)
	End Function

	'@LoadModel(filePath): 加载模型对象

	Public Function LoadModel(filePath)
		On Error Resume Next
		Dim Str, cname
		Str = LoadFile(modelPath_&filePath&".asp")
		If Str <> -1 Then
			Execute Str
			If InStr(filePath, "/")>0 Then
				cname = Replace(filePath, "/", "_")
			Else
				cname = filePath
			End If
			Set LoadModel = Eval("new model_"&cname)
		End If
		If Err Then OutErr("LoadModel:"&filePath&":"&Err.Number&":"&Err.Description)
	End Function

	'@LoadLanguage(filePath): 语言包

	Public Function LoadLanguage(filePath)
		On Error Resume Next
		'
		Dim Str
		'默认语言包
		If  languageDefaultPath_<>"" Then
		    Str = LoadFile(languageDefaultPath_&filePath&".asp")
		    If Str<> -1 Then
			   ExecuteGlobal Str
		    End If
		End If
		'当前语言包-覆盖
		If  languagePath_<>"" and languagePath_<>languageDefaultPath_ Then
			Str = LoadFile(languagePath_&filePath&".asp")
			If Str<> -1 Then
				ExecuteGlobal Str
			End If
		End If
		If Err Then OutErr("LoadLanguage:"&filePath&":"&Err.Number&":"&Err.Description)
	End Function

	'@LoadView(mb_name, mb_data):渲染视图

	Public Function LoadView(mb_name, mb_data)
		Dim tem
		Set tem = LoadClass("template")
		tem.loader = Me
        tem.pathD_tpl = templateDefaultPath_
		tem.path_tpl = templatePath_
		tem.tempdata = mb_data
		LoadView = tem.fetch(mb_name)
		Set tem = Nothing
	End Function

	'@LoadControlAct(col, act): 实例化对象执行方法，可跳转 Dispatch

	Public Function LoadControlAct(col, act)
		On Error Resume Next
		LoadControlAct = true
		'返回实例对象
		Dim control
		Set control = LoadControl(col)
		'调用方法
		If act<>"" Then
			Eval("control."&act&"_action()")
		Else
			Eval("control.index_action()")
		End If
		If Err Then
			OutErr("LoadControlAct:"&col&"/"&act&":"&Err.Number&":"&Err.Description)
			LoadControlAct = False
		End If
		Set control = Nothing
	End Function

	'@LoadControlTag(col, act): 加载Tag

	Public Function LoadControlTag(col, act)
		On Error Resume Next
		'返回实例对象
		Dim control
		Set control = LoadControl(col)
		'调用方法,返回字符串
		If act = "" Then
			Set LoadControlTag = control
		Else
			LoadControlTag = Eval("control."&act)
		End If
		If Err Then
			OutErr("LoadControlTag:"&col&"/"&act&":"&Err.Number&":"&Err.Description)
			LoadControlTag = ""
		End If
		Set control = Nothing
	End Function

	'替换asp标记

	Private Function IncludeReplace(str)
		str = Replace(str, Chr(60)&Chr(37), "")
		str = Replace(str, Chr(37)&Chr(62), "")
		IncludeReplace = str
	End Function
	
	'@ReadUTF(fileName): 读文件 UTF-8
	
	Public Function ReadUTF(fileName)
		On Error Resume Next
		Set objStream = Server.CreateObject("ADODB.Stream")
		ObjStream.Type = 2 '1二进制, 2文本
		ObjStream.Mode = 3 '1读, 2写, 3读写
		ObjStream.Open
		ObjStream.LoadFromFile fileName
		ObjStream.Charset = "utf-8" 'utf-8, GB2312
		ObjStream.Position = 5 '5 utf-8加bom, 2为utf-8不加bom或ANSI
		ReadUTF = ObjStream.ReadText '读文本
		ObjStream.Close
		Set objStream = nothing
		If Err Then 
		   Err.clear
		   ReadUTF = -1
		End if
	End Function
	
	'@GetMapPath(path): 获取物理路径
	
	Public Function GetMapPath(path)
		If  StrCheck(path) Then
			GetMapPath = -1
		Else
			GetMapPath = server.mappath(path)
		End If
	End Function
	
	'判断是否包含路径非法字符
	Private Function StrCheck(str)
		Dim i, arrays
		StrCheck = False
		If IsNull(str) Or Trim(str) = Empty Then Exit Function
		'
	    arrays = Split(":,*,?,"",<,>,|" , ",")
		For i = 0 To UBound(arrays)
			If InStr(str, arrays(i)) > 0 Then
				StrCheck = True
				Exit Function
			End If
		Next
	End Function

	'错误提示

	Public Sub OutErr(ErrMsg)
		If isDebug_ = true Then
			Response.charset = "utf-8"
			Response.Write ErrMsg
			Response.End
		End If
	End Sub

	'@ViewApp(): Application缓存文件查看

	Public Function ViewApp()
		Dim str
		Dim i
		i = 0
		For Each a In Application.Contents
			i = i + 1
			str = str & i&":"&a&Chr(10) 'Application(a)
		Next
		ViewApp = str
	End Function
	
	'@ClearApp(): Application缓存文件清除

	Public Function ClearApp()
		For Each objItem in Application.Contents
			If Left(objItem, 5) = "file_" Then
				Application.Contents.Remove(objItem)
			End If
		Next
	End Function

End Class
%>