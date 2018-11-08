<%
'@title: Control_Index
'@author: ekede.com
'@date: 2018-10-30
'@description: 查看帮助文件

Class Control_Index

	Private template_

    Private Sub Class_Initialize()
        Set template_ = loader.LoadClass("Template")
            template_.loader = loader
            template_.path_tpl = PATH_MODULE&wts.route.module&"/"&PATH_VIEW
    End Sub

    Private Sub Class_Terminate()
        Set template_ = Nothing
    End Sub
	
	'@Index_Action(): 

    Sub Index_Action()
        Call View_Action()
    End Sub
	
	'@View_Action(): 查看
	
    Public Sub View_Action()
	    '加css,js
		Dim cdn
		cdn="https://cdnjs.cloudflare.com/ajax/libs/" 'https://cdn.bootcss.com/
	    template_.SetVal "script/src", cdn&"SyntaxHighlighter/3.0.83/scripts/shCore.js"
	    template_.UpdVal "script"
	    template_.SetVal "script/src", cdn&"SyntaxHighlighter/3.0.83/scripts/shBrushVb.js"
	    template_.UpdVal "script"
	    template_.SetVal "style/href", cdn&"SyntaxHighlighter/3.0.83/styles/shCoreMidnight.css"
	    template_.UpdVal "style"

		'接收文件
        filename=wts.requests.querystr("f")
        'c = wts.cache.GetCache(filename)
		c=-1
        If c <> -1 Then
		    moban=c
		Else
			f=wts.fso.Reads(wts.fso.getMapPath("./")&"\"&replace(filename,"_","\")&".asp","UTF-8")
			If f = -1 Then
			   tag_frame = "WTS ASP FRAME"
			   template_.SetVal "title",tag_frame
			Else
			   tag_help = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=help")
			   template_.SetVal "tag_help",tag_help
			   GetREM(f)
			End IF
			GetList(filename)
			'
			template_.SetVal "tag_frame",tag_frame
			moban = template_.Fetch("help.htm")
			'wts.cache.SetCache filename, moban
        End If
		'
		wts.responses.SetOutput moban
	End Sub
	
	'取列表
	
	Private Sub GetList(f)

		p_inc=wts.fso.getmappath("./")&"\"
		Set d = Server.CreateObject("Scripting.Dictionary")
			LoadData wts.fso.GetMapPath("./"),d
			For Each k in d
			    t=d(k)
				t=replace(t,p_inc,"")
				t=replace(t,"\","_")
				t=replace(t,".asp","")
				d(k)=t
			Next
            '
			For Each k in d
			   link=wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=help/index/view&f="&d(k))
			   template_.SetVal "list/name", replace(d(k),"_","/")
			   template_.SetVal "list/link", link
			   If f=d(k) Then template_.SetVal "list/select", "&lt;--"
			   template_.UpdVal "list"
			Next
	    Set d = Nothing
		'
	End Sub
	
	'取详细内容
	
	Private Sub GetREM(str)
	
	    If str="" Then Exit Sub
		'属性方法
		Set matches = wts.fun.MatchesExp(str,"'"&""&"@(.*):(.*)\r") '增加了个无用空格，避免被当作注释显示到前端
		For Each x in matches
			names = x.SubMatches(0)
			content = x.SubMatches(1)
			If names="title" or names= "author"  or names = "date" or names="description" Then
			    If names="title" Then template_.SetVal "title",wts.fun.TrimVBcrlf(content)
			    If names="description" Then template_.SetVal "description",wts.fun.TrimVBcrlf(content)
				'手工单独添加行
				link = wts.route.ReWrite(wts.site.config("base_url"), "index.asp?route=detail/index") '无id命名
				template_.SetVal "head/name", names
				template_.SetVal "head/content", replace(content,";","<br/>")
				template_.UpdVal "head"
			ElseIf instr(names,"(") Then
				template_.SetVal "func/name", names
				template_.SetVal "func/content", replace(content,";","<br/>")
				template_.UpdVal "func"
			Else
				template_.SetVal "proper/name", names
				template_.SetVal "proper/content", replace(content,";","<br/>")
				template_.UpdVal "proper"
			End If
		Next
		Set matches = nothing
		'例子
		Set matches = wts.fun.MatchesExp(str,"'"&"#(.*):([\s\S]*?)'##") '增加了个无用空格，避免被当作注释显示到前端
		For Each x in matches
			names = x.SubMatches(0)
			content = x.SubMatches(1)
			template_.SetVal "example/name", names
			template_.SetVal "example/content", content
			template_.UpdVal "example"
		Next
		Set matches = nothing
		
	End Sub
	
	'递归文件
	
    Private Sub LoadData(dirPath,data)
	
        Dim fso
		Dim objFolder
		Dim objFiles,objSubFolders
		
		Set fso = server.CreateObject("scripting.filesystemobject")
        Set objFolder = fso.GetFolder(DirPath)
		
		'文件列表集合
        Set objFiles = objFolder.Files
        For Each objFile in objFiles
            fpathname = DirPath &"\"& objFile.Name
			If wts.fun.getext(fpathname) = ".asp" Then
			   data(fpathname)=fpathname
			End If
        Next
		Set objFiles = nothing

        '子文件夹集合
        Set objSubFolders = objFolder.SubFolders
        For Each objSubFolder in objSubFolders
            pathname = DirPath &"\"& objSubFolder.Name
            Call LoadData(pathname,data)
        Next
        Set objSubFolders = Nothing
		
		Set objFolder = nothing
		Set fso = nothing
		
    End Sub

End Class
%>