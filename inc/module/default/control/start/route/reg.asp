<%
'@title: Control_Start_Route_Reg
'@author: ekede.com
'@date: 2018-08-12
'@description: 正则路由

Class Control_Start_Route_Reg

    Private route_
    Private tempKeys_
    Private regEx_
	Private isRule
	private rPath
	private nPath

    '@route: route对象依赖

    Public Property Let route(Value)
        Set route_ = Value
    End Property
	
    Private Sub Class_Initialize()
	    isRule = False
        Set regEx_ = New RegExp
        Set tempKeys_ = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
        Set tempKeys_ = Nothing
        Set regEx_ = Nothing
    End Sub

    '@DeWrite(ByVal r_path): 解码

    Public Sub DeWrite(ByVal r_path)
	    rPath = r_path
		'
		For Each k in tempKeys_
		    Rule k,tempKeys_(k)
		Next
		'正则解析成斜线路由处理
		If isRule Then 
		   If IsObject(route_("slash")) Then route_("slash").Dewrite npath 
		End If
    End Sub
	
	'解串
	Private Sub rule(ByVal sPattern,ByVal sContent)
		If isRule Then exit sub
		regEx_.Pattern = sPattern ' 设置模式。
		If regEx_.Test(rPath) Then
		   nPath = regEx_.Replace(rPath, sContent)
		   isRule = True
		End If
	End Sub
	
    '@ReWrite(ByVal str): 编码

    Public Function ReWrite(ByVal str)
    End Function
	
    '@SetRegKey(keys, values): 设置正则
	
    Public Sub SetRegKey(keys, values)
        tempKeys_(keys) = values
    End Sub

End Class
%>