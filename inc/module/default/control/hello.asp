<%
'@title: Control_Hello
'@author: ekede.com
'@date: 2018-06-09
'@description: Hello WTS

Class Control_Hello

    Private Sub Class_Initialize()
        '加载语言包
        loader.LoadLanguage "hello"
		'加css,js
	    wts.template.SetVal "script/src", wts.site.config("base_static_url")&"js/hello.js"
	    wts.template.UpdVal "script"
	    wts.template.SetVal "style/href", wts.site.config("base_static_url")&"css/hello.css"
	    wts.template.UpdVal "style"
    End Sub

    Private Sub Class_Terminate()
    End Sub
	
	'@Index_Action(): 控制器内部转向
	
    Sub Index_Action()
	     Call List_Action()
	End Sub
	
	'@List_Action(): 查 列表,翻页

    Sub List_Action()
        '#列表分页演示:
        '接收参数
        page = wts.valid.IntNum(wts.requests.querystr("page"), 1, 500, "")
		
		'调用模型
        Set mHello = loader.LoadModel("Hello")
        Set rs = mHello.GetAll() '获取数据集合
        Set oPage = loader.LoadClass("PageList") '调用分页对象
        oPage.tempdata = wts.site.tempdata '用全局临时数据存储器取代默认临时数据存储器
        oPage.CurrentPage = page '设置当前页
        oPage.MaxPerPage = 5 '设置每页显示条数
        n = oPage.List("news", rs) '将数据集合导入分页对象,并返回当前页集合条数
        If n = 0 Then
            wts.errs.AddMsg "no page"
            wts.errs.Out 404
        Else
            '遍历将链接等将后生成数据更新到分页集合中
            For i = 0 To n
                link_str = "index.asp?route=hello/detail&id="&wts.template.GetVali("news/id", i)
                wts.template.setVali "news/link", i, wts.route.ReWrite(wts.site.config("base_url"), link_str)
				'
                link_str = "index.asp?route=hello/edit&id="&wts.template.GetVali("news/id", i)
                wts.template.setVali "news/link_edit", i, wts.route.ReWrite(wts.site.config("base_url"), link_str)
				'
                link_str = "index.asp?route=hello/del&id="&wts.template.GetVali("news/id", i)
                wts.template.setVali "news/link_del", i, wts.route.ReWrite(wts.site.config("base_url"), link_str)
                '
                link_img = wts.route("pic").ReWritePic(wts.site.config("base_pic_url"), "image/no.gif", 50, 50, "")
                wts.template.setVali "news/pic", i, link_img
                '
                wts.template.setVali "news/time", i, wts.fun.FormatDate(wts.times,1)
            Next
            '生成分页"news_page"
            oPage.Plist wts.route,wts.site.config("base_url"),"index.asp?route=hello/list"
        End If
        Set oPage = Nothing
        rs.Close
        Set rs = Nothing
        Set mHello = Nothing
		
		'添加链接
		wts.template.setVal "tag_addlink", wts.route.ReWrite(wts.site.config("base_url"), "index.asp?route=hello/add")
		
		'设置标题
		wts.template.SetVal "title","WTS列表,翻页演示"
		
		'渲染模板
        moban = wts.template.Fetch("hello_list.htm")
		
        '输出内容
        wts.responses.SetOutput moban
		'##
    End Sub
	
	'@Detail_Action(): 查 详情

    Sub Detail_Action()
        '#详情页演示:
        '接收参数
        id = wts.valid.IntNum(wts.requests.querystr("id"), 0, 0, "")
        wts.template.SetVal "tag_id", id
		
		'调用模型
        Set mHello = loader.LoadModel("hello")
        Set rs = mHello.GetNameById(id)
        If rs.recordcount>0 Then
            wts.template.SetVal "tag_name", rs("name")
		    wts.template.SetVal "title",rs("name")
			'meta
		    wts.template.SetVal "meta/name","description"
		    wts.template.SetVal "meta/content",rs("name")
	        wts.template.UpdVal "meta"
        Else
            wts.template.SetVal "tag_name", "no name"
        End If
        rs.Close
        Set rs = Nothing
        Set mHello = Nothing
		
		'渲染模板
        moban = wts.template.Fetch("hello_detail.htm")
		
		'输出内容
        wts.responses.SetOutput moban
		'##
	End Sub
	
	'@Add_Action(): 增 表单

    Sub Add_Action()
	    submit_url=wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=hello/addsave")
        wts.template.SetVal "submit_url",submit_url
		
        '渲染模板
        moban = wts.template.Fetch("hello_form.htm")
		
        '输出内容
        wts.responses.SetOutput moban
	End Sub
	
	'@AddSave_Action():  增 保存

    Sub AddSave_Action()
	    name = wts.valid.text(wts.requests.forms("name"), 1, 50, "")
        Set mHello = loader.LoadModel("Hello")
        id=mHello.Add(name)
        Set mHello = Nothing
		'
		wts.responses.Direct wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=hello/edit&id="&id)
	End Sub
	
	'@Edit_Action(): 改 表单

    Sub Edit_Action()
        'querystr
        id = wts.valid.IntNum(wts.requests.querystr("id"), 0, 0, "")
        wts.template.SetVal "tag_id", id

        '调用模型
        Set mHello = loader.LoadModel("hello")
        Set rs = mHello.GetNameById(id)
        If rs.recordcount>0 Then
		    wts.template.SetVal "name",rs("name")
		    wts.template.SetVal "id",rs("id")
        End If
        rs.Close
        Set rs = Nothing
        Set mHello = Nothing
	
	    '提交链接
	    submit_url=wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=hello/editsave")
        wts.template.SetVal "submit_url",submit_url
		
        '渲染模板
        moban = wts.template.Fetch("hello_form.htm")
		
        '输出内容
        wts.responses.SetOutput moban
	End Sub
	
	'@EditSave_Action(): 改 保存

    Sub EditSave_Action()
	    '接收并验证数据
	    name = wts.valid.text(wts.requests.forms("name"), 1, 50, "Invalid Name")
		id = wts.valid.intNum(wts.requests.forms("id"), 1, 0, "Invalid Id")
		If wts.errs.foundErr Then wts.errs.OutMsg
		'判断id是否存在
		Set data = Server.CreateObject("Scripting.Dictionary")
        Set mHello = loader.LoadModel("Hello")
        Set rs = mHello.GetNameById(id)
        If rs.recordcount>0 Then
		    data("name")=name
			data("id")=id
		    mHello.Edit data '添加数据
		Else
		    wts.errs.AddMsg "Invalid id"
		    wts.errs.OutMsg
        End If
        rs.Close
        Set rs = Nothing
        Set mHello = Nothing
		Set data = Nothing
		'跳转
		wts.responses.Direct wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=hello/list")
	End Sub
	
	'@Del_Action(): 删

    Sub Del_Action()
	    '接收并验证数据
		id = wts.valid.intNum(wts.requests.querystr("id"), 1, 0, "Invalid Id")
		If wts.errs.foundErr Then wts.errs.OutMsg
		'删除
		Set mHello = loader.LoadModel("Hello")
		    mHello.Del id
		Set mHello = Nothing
		'跳转
		wts.responses.Direct wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=hello/list")
	End Sub

End Class
%>