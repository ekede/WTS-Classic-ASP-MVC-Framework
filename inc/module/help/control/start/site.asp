<%
'@title: Control_Start_Site
'@author: ekede.com
'@date: 2018-02-01
'@description: Start

Class Control_Start_Site

    '@config:   配置数据存储,便于不同对象间交换数据

    Dim config

    Private Sub Class_Initialize()
        Set config = Server.CreateObject("Scripting.Dictionary")
    End Sub
    Private Sub Class_Terminate()
        Set config = Nothing
    End Sub

    '@Start(): 启动模块

    Public Function Start()
        '路由
		wts.route.routers = "" '必须
        wts.route.rewrite_on = True
        wts.route.DeWrite()
		'静态根路径
		config("base_url") = wts.route.baseAddr
        '
        loader.LoadControlAct wts.route.control, wts.route.action
        '
        wts.responses.Outputs
		
    End Function

End Class
%>