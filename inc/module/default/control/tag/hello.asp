<%
'@title: Control_Tag_Hello
'@author: ekede.com
'@date: 2018-06-09
'@description: 模块制作演示

Class Control_Tag_Hello
    '
    Dim temp_data

    Private Sub Class_Initialize()
        Set temp_data = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
        Set temp_data = Nothing
    End Sub
	
	'@Hello(): 返回模块信息

    Function Hello()
        temp_data("tag_para") = "this is a para"
        hello = loader.loadView("tag/hello.htm", temp_data)
    End Function

End Class
%>