<%
'@title: Class_Route_Module
'@author: ekede.com
'@date: 2018-12-06
'@description: 路由前特殊模块判断

Class Class_Route_Module

    Private route_
	
    '@route: route对象依赖

    Public Property Let route(Value)
        Set route_ = Value
    End Property
	
    '@GetModule(r): 根据Requests特殊对象模块

    Public Function GetModule(r)
        GetModule = False
       'route_.module = "default"
    End Function

End Class
%>