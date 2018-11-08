<%
'@title: Control_Start_Site
'@author: ekede.com
'@date: 2018-02-01
'@description: module启动入口

Class Control_Start_Site

    '@config:   配置数据存储,便于不同对象间交换数据
    '@tempdata: 临时数据存储,便于不同对象间交换数据
    Dim config
    Dim tempdata
    '@langId: 站点语言id
    Dim langId
	'@langDefaultPath: 默认语言文件夹
	'@langPath: 当前语言文件夹
	Dim langDefaultPath,langPath
	'@tplDefaultPath: 默认模板文件夹
	'@tplPath: 当前模板文件夹
	Dim tplDefaultPath,tplPath

    Private Sub Class_Initialize()
	    '判断数据安装,正式上线可删除
		If wts.fso.ReportFolderStatus(wts.fso.GetMapPath(PATH_ROOT&PATH_DATA)) = -1 Then
		   wts.responses.Direct(wts.route.basePicAddr&"index.asp?route=help/install")
		End If

        '全局容器
        Set config = Server.CreateObject("Scripting.Dictionary")
        Set tempdata = Server.CreateObject("Scripting.Dictionary")
		'
        Set wts.template = loader.LoadClass("Template")
            wts.template.loader = loader
            wts.template.tempdata = tempdata
        Set wts.cache = loader.LoadClass("Cache")
            wts.cache.fso = wts.fso
        Set wts.db = loader.LoadClass("DB")
            If DB_TYPE > 0 Then wts.db.OpenConn DB_TYPE, DB_PATH, DB_NAME, DB_USER, DB_PASS
            DB_USER = ""
            DB_PASS = ""
		'
        langId = 0
		langDefaultPath = "en"
		langPath = langDefaultPath
		tplDefaultPath = "default"
		tplPath = tplDefaultPath
    End Sub

    Private Sub Class_Terminate()
        Set wts.db = Nothing
        Set wts.cache = Nothing
        Set wts.template = Nothing
        '释放容器
        Set tempdata = Nothing
        Set config = Nothing
    End Sub
	
	'@Start(): 启动模块配置

    Public Function Start()
	
        '路由配置 正则,图片,关键词 reg,pic,key
		wts.route.routers = "pic,key"
        wts.route.rewrite_on = True '开启地址重写

        '初始化站点langId
		'SetBaseAddr() '自定义路由多网址功能
		'
        loader.languageDefaultPath = PATH_MODULE&wts.route.module&"/"&PATH_LANGUAGE&langDefaultPath&"/"
        loader.languagePath = PATH_MODULE&wts.route.module&"/"&PATH_LANGUAGE&langPath&"/"
		'
		templateDefaultPath = PATH_MODULE&wts.route.module&"/"&PATH_VIEW&tplDefaultPath&"/"
        loader.templateDefaultPath = templateDefaultPath&"tpl/"
        wts.template.pathD_tpl = templateDefaultPath&"tpl/"
		'
		templatePath = PATH_MODULE&wts.route.module&"/"&PATH_VIEW&tplPath&"/"
        loader.templatePath = templatePath&"tpl/"
        wts.template.path_tpl = templatePath&"tpl/"
		'
        wts.cache.datapath = PATH_DATA&"cache/"&langId&"/"
		
        '路由分析
	    'SetUrlkey() '自定义路由SEO Url功能
        wts.route.DeWrite()
        '
        config("base_url") = wts.route.baseAddr
        config("base_pic_url") = wts.route.basePicAddr
		config("base_static_url") = wts.route.basePicAddr&PATH_STATIC&wts.route.module&"/"&tplPath&"/"
        '
        loader.LoadControlAct wts.route.control, wts.route.action
        '
        wts.responses.outputs
    End Function
    
	'取得虚拟根目录及其参数
	
    Private Sub SetBaseAddr()
        '
        If wts.route.rewrite_on = FALSE Then Exit Sub
	    dim i,arr,bie_location,bie_base,BaseA
		'
		siteKeys = wts.cache.GetValue("siteKeys")
        If IsArray(siteKeys)=False Then '读数据库-数组演示
           ReDim siteKeys(3)
		   siteKeys(0) = Array(1,"http://localhost/en/","","en","default")
		   siteKeys(1) = Array(2,"http://localhost/cn/","","en","default")
		   siteKeys(2) = Array(3,"http://localhost/de/","","en","default")
		   siteKeys(3) = Array(4,"http://localhost/","","cn","new")
           wts.cache.SetValue "siteKeys", siteKeys
		End If
		'
		bie_location=wts.route.GetBieUrl(wts.requests.standardAddr)
		For i = 0 to Ubound(siteKeys)
		    bie_base=wts.route.GetBieUrl(siteKeys(i)(1))
			If Instr(bie_location,bie_base)>0 then
			   baseA=siteKeys(i)
			   Exit For
			End If
		Next
		'
		If IsArray(baseA) Then
		   langId=baseA(0)
		   langPath = baseA(3)
		   tplPath = baseA(4)
	       wts.route.SetBaseAddr baseA(1)
		   If baseA(2) <> "" Then wts.route.SetBasePicAddr baseA(2)
		Else
		   wts.responses.Die("Invalid Site")
		End if
    End Sub

    '缓存urlKeys+id
		
    Private Sub SetUrlkey()
	
        '是否开启重写
        If wts.route.rewrite_on = False Then Exit Sub
		
        'urlkey路由设置
		If  IsObject(wts.route("key")) Then
			urlKeys = wts.cache.GetValue("urlKeys")
			urlDKeys = wts.cache.GetValue("urlDKeys")
			If IsArray(urlKeys) and IsArray(urlDKeys) Then '读取
				For i = 0 To UBound(urlKeys)
					wts.route("key").SetUrlKey urlKeys(i)(0), urlKeys(i)(1) 
					wts.route("key").SetDUrlKey urlDKeys(i)(0), urlDKeys(i)(1)
				Next
			Else '读数据库-数组演示
				ReDim urlKeys(10),urlDKeys(10)
				For i = 0 To UBound(urlKeys)
					k = "hello"&i&".html"
					v = "hello/detail/id/"&i
					urlKeys(i) = Array(k, v)
					urlDKeys(i) = Array(v, k)
					'
					wts.route("key").SetUrlKey k, v
					wts.route("key").SetDUrlKey v, k
				Next
				wts.cache.SetValue "urlKeys", urlKeys
				wts.cache.SetValue "urlDKeys", urlDKeys
			End If
			wts.route("key").SetUrlKey "index.html", "index/index"
			wts.route("key").SetDUrlKey "index/index", "index.html"
		End If
		
		'正则路由设置
		If  IsObject(wts.route("reg")) Then
			wts.route("reg").SetRegKey "^test-([0-9]+)\.do$","hello/detail/id/$1"
			wts.route("reg").SetRegKey "^robots\.txt$","hello/index/page/2"
		End If
		
    End Sub

End Class
%>