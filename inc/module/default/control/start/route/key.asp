<%
'@title: Control_Start_Route_Key
'@author: ekede.com
'@date: 2018-06-09
'@description: Keyword路由

Class Control_Start_Route_Key

    Private route_
    Private tempKeys_
    Private tempDKeys_

    '@route: route对象依赖

    Public Property Let route(Value)
        Set route_ = Value
    End Property
	
    Public Property Let tempkeys(Value)
        Set tempkeys_ = Value
    End Property
	
    Private Sub Class_Initialize()
        Set tempKeys_ = Server.CreateObject("Scripting.Dictionary")
        Set tempDKeys_ = Server.CreateObject("Scripting.Dictionary")
    End Sub
	
    Private Sub Class_Terminate()
        Set tempDKeys_ = Nothing
        Set tempKeys_ = Nothing
    End Sub
	
    '@DeWrite(ByVal r_path): 解码
	
    Public Sub DeWrite(ByVal r_path)
        If tempkeys_.Exists(r_path) Then
            r_path = tempkeys_(r_path)
			'调用SLASH路由
			If IsObject(route_("slash")) Then route_("slash").Dewrite r_path
        End If
    End Sub

    '@ReWrite(ByVal str): 编码 keyword只编码module/id/1,其余部分仍然动态参数

    Public Function ReWrite(ByVal str)
        Dim i, arr, arr_j
        Dim str_route, str_id, str_para
        Dim r_path
        '
        If tempkeys_.Count = 0 Then Exit Function
        '去index.asp
        If InStr(str, "?")>0 Then
            str = Right(str, Len(str) - InStr(str, "?"))
        End If
        '拆参数
        arr = Split(str, "&")
        For i = 0 To UBound(arr)
            If arr(i)<> "" Then
                arr_j = Split(arr(i), "=")
                If UBound(arr_j) = 1 Then
                    If arr_j(0)<>"" And arr_j(1)<>"" Then
                        If arr_j(0) = "route" Then
                            str_route = arr_j(1)
                        ElseIf arr_j(0) = "id" Then
                            str_id = arr_j(1)
                        Else
                            If str_para = "" Then
                                str_para = arr_j(0)&"="&arr_j(1)
                            Else
                                str_para = str_para&"&"&arr_j(0)&"="&arr_j(1)
                            End If
                        End If
                    End If
                End If
            End If
        Next
        '路由名+id
        If str_route<>"" Then r_path = str_route
        If str_id<>"" Then r_path = r_path&"/id/"&str_id
        '重写+参数
        If tempDKeys_.Exists(r_path) Then
            r_path = tempDKeys_(r_path)
            If str_para<>"" Then r_path = r_path&"?"&str_para
            ReWrite = r_path
        End If
    End Function

    '@SetUrlKey(Keys, values): 设置编码键值

    Public Sub SetUrlKey(Byref keys,Byref values)
        tempKeys_(keys) = values
    End Sub
	
    '@SetDUrlKey(Byref Keys,Byref values): 设置解码键值
	
    Public Sub SetDUrlKey(Byref Keys,Byref values)
        tempDKeys_(keys) = values
    End Sub

End Class
%>