<%
'@title: Framework_Wts
'@author: ekede.com
'@date: 2018-06-09
'@description: WTS框架

Class Framework_Wts

    '@fun: 对象
    '@fso: 对象
    '@errs: 对象
    '@valid: 对象 
    '@logs: 对象 
    '@cache: 对象
    '@cookie: 对象
    '@sessions: 对象 
    '@requests: 对象 
    '@responses: 对象
    '@route: 对象
    '@db: 对象
    '@template: 对象
    '@site: 对象

    Dim fun
    Dim fso
    Dim errs
    Dim valid
    Dim logs
    Dim cache
    Dim cookie
    Dim sessions
    Dim requests
    Dim responses
    Dim route
    Dim db
    Dim template
    Dim site
	
	Private attr_
    Private zone_
    Private times_
	
    '@version: 版本
	
    Public Property Get version
        Version = "1.0.0"
    End Property
	
    '@zone: 时区
	
    Public Property Get zone
        zone = zone_
    End Property
	
    '@times: 时间

    Public Property Get times
        times = times_
    End Property

    '@attr: 自定义属性 attr(k)=v

    Public Property Let attr(k, v)
		If  IsObject(v) Then
			Set attr_(k) = v
		Else
			attr_(k) = v
		End If
    End Property

    Public Default Property Get attr(k)
		If  IsObject(attr_(k)) Then
			Set attr = attr_(k)
		Else
			attr = attr_(k)
		End If
	End Property

    Private Sub Class_Initialize()
        zone_ = 8
        times_ = Now()
        Set attr_ = Server.CreateObject("Scripting.Dictionary")
    End Sub
    Private Sub Class_Terminate()
        Set attr_ = Nothing
    End Sub

    '@Start():启动框架

    Public Sub Start()

        'loader 类库路径配置
        loader.classPath = PATH_CLASS
		
        '全局对象
        Set fun = loader.LoadClass("Function")
        Set fso = loader.LoadClass("Fso")

        Set logs = loader.LoadClass("Log")
            logs.fso = fso
            logs.LogPath = PATH_DATA&"logs/"

        Set errs = loader.LoadClass("Error")
            errs.loader = loader

        Set valid = loader.LoadClass("Valid")
            valid.errs = errs
			
        Set cookie = loader.LoadClass("Cookie")
        Set sessions = Loader.LoadClass("Session")
        Set requests = loader.LoadClass("Request")
        Set responses = loader.LoadClass("Response")

        Set route = loader.LoadClass("Route")
            route.fun = fun
            route.loader = loader
            route.requests = requests
            route.modules = MODULES
			route.start()
			route.GetModule()
			
        'loader MVCL默认路径配置
        loader.controlPath = PATH_MODULE&route.module&"/"&PATH_CONTROL
        loader.modelPath = PATH_MODULE&route.module&"/"&PATH_MODEL
        loader.languagePath = PATH_MODULE&route.module&"/"&PATH_LANGUAGE '&language_path
        loader.templatePath = PATH_MODULE&route.module&"/"&PATH_VIEW '&view_path
		
        '交接start
        Set site = loader.loadControl("Start/Site")
            site.start()
        Set site = Nothing
		
    End Sub
	
	'@Finish():释放框架
	
    Public Sub Finish()
        '释放对象
        Set route = Nothing
        Set responses = Nothing
        Set requests = Nothing
        Set sessions = Nothing
        Set cookie = Nothing
        Set valid = Nothing
        Set errs = Nothing
        Set logs = Nothing
        Set fso = Nothing
        Set fun = Nothing
    End Sub
	
End Class
%>