<%
'@title: Class_Route
'@author: ekede.com
'@date: 2018-06-09
'@description: 网站路由,通过操作Request对象中的url，解析出模块,控制器,方法,参数

Class Class_Route

	Private isDebug_
    Private loader_
    Private requests_
	Private fun_
	Private s_
	
	
    '@loader: loader对象依赖

    Public Property Let loader(Value)
        Set loader_ = Value
    End Property
	
    Public Property Get loader
        Set loader = loader_
    End Property
	
    '@fun: fun对象依赖

    Public Property Let fun(Value)
        Set fun_ = Value
    End Property
	
    Public Property Get fun
        Set fun = fun_
    End Property
	
    '@requests: requests对象依赖

    Public Property Let requests(Value)
        Set requests_ = Value
    End Property
	
    Public Property Get requests
        Set requests = requests_
    End Property
	
	'@routers: 路由集合

    Public Property Let routers(Value)
		Dim arr
		'其他路由
		arr=Split(cstr(Value),",")
		For i = 0 to UBOUND(arr)
		    If  arr(i)<>"" Then
				Set s_(arr(i)) = loader_.LoadControl("start/route/"&arr(i))
					s_(arr(i)).route = Me
			End If
			If Err Then OutErr("路由加载错误:"&arr(i)&":"&Err.Description)
		Next
		'加载斜线路由
		Set s_("slash") = loader_.LoadClass("route/slash")
			s_("slash").route = Me
		If Err Then OutErr("路由加载错误:slash:"&Err.Description)
    End Property
	
    Public Property Get routers
	    Dim r,s
		For Each r in s_
		    If s = "" Then
			   s=r
			Else
			   s=s&","&r
			End If
		Next
		routers = s
    End Property
	
    '@s: 单个路由,默认属性
	
    Public Default Property Get s(k)
		If  s_.Exists(k) Then Set s = s_(k)
	End Property
	
	'@baseAddr: 默认根目录,script所在目录
	
	Dim baseAddr
	
	'@routeAddr: 路由地址
	
	Dim routeAddr
	
	'@basePicAddr: 图片根目录,网站根地址
	
	Dim basePicAddr
	
	'@routePicAddr: 图片路由地址
	
	Dim routePicAddr
	
	'@modules: 已开启模块
	
    Dim modules
	
    '@module: 当前模块
	
    Dim module
	
    '@control: 当前控制器
	
    Dim control
	
    '@action: 当前方法
	
    Dim action
	
    '@rewrite_on: 开启url重写
	
    Dim rewrite_on
	
    '@dewrite_on: url解码是否成功
	
    Dim dewrite_on

    Private Sub Class_Initialize()
		If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
		'
        modules = "default"
        Set s_ = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
		For Each r in s_
			Set s_(r) = Nothing
		Next
        Set s_ = Nothing
    End Sub
	
    '--------------------------------------0 初始化路由
	
	'@Start(): 路由启动,需要预先设置loader,request外部依赖属性
	
    Public Sub Start()
	    On Error Resume Next
	    Dim arr,i
		'
        rewrite_on = false
        dewrite_on = false
        module = "default"
	    '
	    SetBaseAddr requests_.baseAddr
	    SetBasePicAddr requests_.basePicAddr
    End Sub
	
    '--------------------------------------1 首先获得模块信息然后Start对应模块
	
	'@GetModule(): 默认获取模块,特殊模块劫持修改此方法
	
    Public Sub GetModule()
        Dim temp_path
        '
        If requests_.Status404 Then
            temp_path = RouteAddr
        Else
            temp_path = requests_.querystr("route")
        End If
        '
        If  temp_path <> "" Then
            temp_array = Split(temp_path, "/")
            temp_path = temp_array(0)
            If fun_.StrEqual(temp_path,modules,",") Then module = temp_array(0)
        End If
    End Sub
	
	'--------------------------------------2 设置根目录,计算路由地址
	
	'@SetBaseAddr(str): 指定网页baseAddr,得到routeAddr路由
	
    Public Sub SetBaseAddr(str)
	    baseAddr = str
		'
	    dim tmp_standardAddr,tmp_baseAddr
		tmp_standardAddr=GetBieUrl(requests_.standardAddr)
		tmp_baseAddr=GetBieUrl(str)
	    If Instr(tmp_standardAddr,tmp_baseAddr)>0 Then
           routeAddr = Replace(tmp_standardAddr, tmp_baseAddr, "")
		Else
		   routeAddr = "error/e404"
		End if
    End Sub
	
	'@SetBasePicAddr(str): 指定图片basePicAddr,得到routePicAddr路由
	
    Public Sub SetBasePicAddr(str)
	    basePicAddr = str
	    '
	    dim tmp_standardAddr,tmp_basePicAddr
		tmp_standardAddr=GetBieUrl(requests_.standardAddr)
		tmp_basePicAddr=GetBieUrl(str)
	    If Instr(tmp_standardAddr,tmp_basePicAddr)>0 Then
           routePicAddr = Replace(tmp_standardAddr, tmp_basePicAddr, "")
		End if
    End Sub
	
	'-------------------------------------- 路由对象
	
	' 路由对 1,2,3
	
    ' 路由对象0

    Private Sub DeWrite_Ask(r_path,p_path)
        If InStr(r_path, "?")>0 Then
            temp_array = Split(r_path, "?")
            r_path = temp_array(0)
            Add_Query temp_array(1)
        End If
		'
        If InStr(p_path, "?")>0 Then
            temp_array = Split(p_path, "?")
            p_path = temp_array(0)
            Add_Query temp_array(1)
        End If
    End Sub
	
    Private Function Add_Query(byval Web_Query)
        Dim i, j, arr, arr_j
        arr = Split(Web_Query, "&")
        For i = 0 To UBound(arr)
            If arr(i)<> "" Then
                arr_j = Split(arr(i), "=")
                If UBound(arr_j) = 1 Then
                    If arr_j(0)<>"" And arr_j(1)<>"" Then
                        requests_.querystr(arr_j(0)) = fun_.urldecodes(arr_j(1)) '++query
                    End If
                End If
            End If
        Next
    End Function

    ' 路由对象4

    Private Sub DeWrite_404()
        c_path = PATH_MODULE&module&"/"&PATH_CONTROL
        c = "error"
        If loader_.LoadFile(c_path&c&".asp")<> -1 Then
            control = c
            action = "e404"
            dewrite_on = true
        End If
    End Sub
	
    '-------------------------------------- 编码，解码

    '@ReWrite(base, r_path): 路由编码

    Public Function ReWrite(base, r_path)
	    On Error Resume Next
        Dim str
		
        'status
        If rewrite_on = false Then
            ReWrite = base&r_path
            Exit Function
        End If

        '遍历路由
		For Each r in s_
            str = s_(r).ReWrite(r_path)
			If str <> "" Then
				ReWrite = base&str
				Exit Function
			End If
			If Err Then OutErr("路由编码错误:"&r&":"&r_path&Err.Description)
		Next

        'default
        ReWrite = base&"#"

    End Function
	
    '@DeWrite(): 路由解码

    Public Sub DeWrite()
	    On Error Resume Next
		
        If requests_.Status404 Then
           r_path = RouteAddr
		   p_path = RoutePicAddr
        Else
           r_path = requests_.querystr("route")
		   If s_.Exists("slash") Then s_("slash").DeWrite r_path
		   If dewrite_on Then Exit Sub
        End If
		
		'?
        DeWrite_Ask r_path,p_path
		
        '遍历路由
		For Each r in s_
		    If r = "pic" Then
			   s_(r).DeWrite p_path
            Else
			   s_(r).DeWrite r_path
			End If
			If dewrite_on Then Exit Sub
            If Err Then OutErr("路由解码错误:"&r&Err.Description)
		Next

        '404
        DeWrite_404
        If dewrite_on Then Exit Sub

        'no control
        OutErr("no control")
    End Sub
	
	'@GetBieUrl(url): 去https,http,端口，用于对比网址
	
	Public Function GetBieUrl(url)
		dim tmp
		tmp=split(url,"://")
		if instr(tmp(1),":")>0 then
		   GetBieUrl=left(tmp(1),instr(tmp(1),":")-1) + right(tmp(1),len(tmp(1))-instr(tmp(1),"/")+1)
		else
		   GetBieUrl=tmp(1)
		end if
	End Function
	
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