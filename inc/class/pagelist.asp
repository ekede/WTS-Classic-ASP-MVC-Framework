<%
'@title: Class_PageList
'@author: ekede.com
'@date: 2017-02-13
'@description: 分页,翻页

Class Class_PageList

    Private pageNum_
    Private pageKey_
	Private tempdata_

    '@currentPage: 当前页
	
    Dim currentPage
	
    '@maxPerPage: 每页显示数
	
    Dim maxPerPage
	
    '@tempdata: 内容存放容器

    Public Property Let tempdata(Values)
        If VarType(Values) = 9 Then Set tempdata_ = Values
    End Property

    Private Sub Class_Initialize
        currentPage = 1
        maxPerPage = 10
        pageNum_ = 1
        Set tempdata_ = Server.CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate()
        Set tempdata_ = Nothing
    End Sub

    '@List(ByRef keys,ByRef Rs): 计算分页,将rs二维表抽象为一维存入dictionary对象,返回当前页条数

    Public Function List(ByRef keys,ByRef Rs)
        Dim n
        '
        If keys = "" Then Exit Function
		pageKey_ = keys
        '
        If Rs.EOF And Rs.bof Then
            n = 0
        Else
            Rs.Pagesize = MaxperPage
            pageNum_ = Rs.PageCount
            If currentPage>pageNum_ Then
                currentPage = pageNum_
                n = 0
            Else
                Rs.Move (currentPage -1) * MaxperPage
                n = 0
                Do While Not rs.EOF
                    For Each field In rs.fields
                        tempdata_(keys&"/"&field.Name&"/"&n) = field.Value
                    Next
                    n = n + 1
                    If n>= MaxPerPage Then Exit Do
                    rs.movenext
                Loop
            End If
            tempdata_(keys) = n
        End If
        '
        list = n
    End Function

    '@Plist(ByRef route, ByRef base, ByRef url): 生成翻页链接,保存在dictionary对象中,供模板loop读取

    Public Function Plist(ByRef route, ByRef base, ByRef url)
        Dim i,n
		n = 0
        '
        PageUrl route, base, url, currentPage, Currentpage&"/"&pageNum_, 0 ,n
		n = n + 1

        '计算当前开始结束页
        naviLength = 5
        startPage = (currentPage \ naviLength) * naviLength + 1
        If (currentPage Mod naviLength) = 0 Then startPage = startPage - naviLength
        endPage = startPage + naviLength - 1
        If endPage>pageNum_ Then endPage = pageNum_
		
        '前移分页
        If startPage>1 Then
            i = currentPage - (currentPage Mod naviLength) - naviLength + 1
            PageUrl route, base, url, i, "&lt;&lt;", 0 ,n
			n = n + 1
        End If
        '前移一页
        If currentPage <> 1 Then
            i = currentPage -1
            PageUrl route, base, url, i, "&lt;", 0 ,n
			n = n + 1
        End If
        '当前分页
        For i = startPage To endPage
            If Currentpage = i Then
               PageUrl route, base, url, i, i, 1 ,n
            Else
               PageUrl route, base, url, i, i, 0 ,n
            End If
			n = n + 1
        Next
        '后移一页
        If currentPage <> pageNum_ Then
            i = currentPage + 1
            PageUrl route, base, url, i, "&gt;", 0 ,n
			n = n + 1
        End If
        '后移分页
        If endPage<pageNum_ Then
            i = currentPage - (currentPage Mod naviLength) + naviLength + 1
			If i > pageNum_ Then i = pageNum_
            PageUrl route, base, url, i, "&gt;&gt;", 0 ,n
			n = n + 1
        End If
        '
		tempdata_(pageKey_&"_page") = n
    End Function

    'link

    Private Function PageUrl(ByRef route,ByRef base,ByRef url,ByRef i,ByRef navi,ByRef selected,ByRef n)
        link = route.ReWrite(base,url&"&page="&i)
		tempdata_(pageKey_&"_page/link/"&n)=link
		tempdata_(pageKey_&"_page/num/"&n)=navi
		tempdata_(pageKey_&"_page/selected/"&n)=selected
    End Function

End Class
%>