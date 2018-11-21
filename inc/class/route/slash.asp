<%
'@title: Class_Route_Slash
'@author: ekede.com
'@date: 2018-06-09
'@description: 斜线路由

Class Class_Route_Slash

    Private route_
	
    '@route: route对象依赖

    Public Property Let route(Value)
        Set route_ = Value
    End Property

    '@DeWrite(ByVal r_path): 解码

    Public Sub DeWrite(ByVal r_path)
        Dim i, c, c_path, j, k
        Dim temp_array
        '
        If r_path = "" Then
            'route_.module = "default" '默认路由为当前路由
            c_path = PATH_MODULE&route_.module&"/"&PATH_CONTROL
            c = "index"
            If route_.loader.LoadFile(c_path&c&".asp")<> -1 Then route_.control = c '--loader
            route_.dewrite_on = true
            Exit Sub
        End If
        '
        temp_array = Split(r_path, "/")
        For i = 0 To UBound(temp_array)
		    If  temp_array(i) = "" Then
                '空斜线跳过
			Else
				If i = 0 Then
					r_path = temp_array(0)
					If route_.fun.StrEqual(r_path, route_.modules,",") Then
						route_.module = temp_array(0)
						c_path = PATH_MODULE&route_.module&"/"&PATH_CONTROL
					Else
					   'route_.module = "default" '默认路由为当前路由
						c_path = PATH_MODULE&route_.module&"/"&PATH_CONTROL
						c = temp_array(0)
						If route_.loader.LoadFile(c_path&c&".asp")<> -1 Then route_.control = c '--loader
						k = 1
					End If
				Else
					If route_.control = "" Then
						If c = "" Then
							c = temp_array(i)
						Else
							c = c&"/"&temp_array(i)
						End If
						If route_.loader.LoadFile(c_path&c&".asp")<> -1 Then route_.control = c '--loader
						k = 1 
					Else
						If route_.action = "" Then
							route_.action = temp_array(i)
							j = 0
						Else
							j = j + 1
							If j = 2 Then
								route_.requests.querystr(temp_array(i -1)) = temp_array(i) '++query
								j = 0
							End If
						End If
					End If
				End If
			End If
        Next
        '未做控制器判断,查看默认控制器是否存在
		If route_.control="" And k <> 1 Then
            c = "index"
            If route_.loader.LoadFile(c_path&c&".asp")<> -1 Then route_.control = c '--loader
		End If
		'
        If route_.control<>"" Then route_.dewrite_on = true
    End Sub
	
    '@ReWrite(ByVal str): 编码

    Public Function ReWrite(ByVal str)
        If InStr(str, "?")>0 Then
            ReWrite = Add_Slash(Right(str, Len(str) - InStr(str, "?")))
        Else
            ReWrite = str
        End If
    End Function

    Private Function Add_Slash(byval Web_Query)
        Dim i, j, arr, arr_j, str, str_route
        arr = Split(Web_Query, "&")
        For i = 0 To UBound(arr)
            If arr(i)<> "" Then
                arr_j = Split(arr(i), "=")
                If UBound(arr_j) = 1 Then
                    If arr_j(0)<>"" And arr_j(1)<>"" Then
                        If arr_j(0) = "route" Then
                            str_route = arr_j(1)
                        ElseIf arr_j(0) = "urlkey" Then
                            '排除
                        Else
                            If str = "" Then
                                str = arr_j(0)&"/"&arr_j(1)
                            Else
                                str = str&"/"&arr_j(0)&"/"&arr_j(1)
                            End If
                        End If
                    End If
                End If
            End If
        Next
		if str="" then
           Add_Slash = str_route
		else
           Add_Slash = str_route&"/"&str
		end if
    End Function

End Class
%>