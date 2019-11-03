<%
'@title: Control_Index
'@author: ekede.com
'@date: 2018-06-09
'@description: 首页访问

Class Control_Index

    Private Sub Class_Initialize()
        loader.LoadLanguage "hello"
    End Sub

    Private Sub Class_Terminate()
    End Sub
	
	'@Index_Action(): 默认首页

    Public Sub Index_Action()

	   '标题
	   wts.template.SetVal "title",wts.site.tempdata("Lan_Hello")
	   '
	   link = wts.route.ReWrite(wts.route.basePicAddr,"index.asp?route=help/index")
       wts.template.setVal "tag_link", link
	   '
	   tag_list = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=hello/index")
	   wts.template.SetVal "tag_list",tag_list
	   '
	   tag_upload = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=upload/index")
	   wts.template.SetVal "tag_upload",tag_upload
	   '
	   tag_json = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=json/index")
	   wts.template.SetVal "tag_json",tag_json
	   '
	   tag_crypt = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=crypt/index")
	   wts.template.SetVal "tag_crypt",tag_crypt
	   '
	   tag_spider = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=index/spider")
	   wts.template.SetVal "tag_spider",tag_spider
	   '
	   tag_stat = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=index/stat")
	   wts.template.SetVal "tag_stat",tag_stat
	   '
	   tag_verify = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=index/verify")
	   wts.template.SetVal "tag_verify",tag_verify
	   '
	   tag_mail = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=index/mail")
	   wts.template.SetVal "tag_mail",tag_mail
	   '
	   tag_zip = wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=index/zip")
	   wts.template.SetVal "tag_zip",tag_zip
	   
       '渲染模板
       moban = wts.template.Fetch("index.htm")
		
       '输出内容
       wts.responses.SetOutput moban
		
    End Sub
	
	'@Verify_Action(): 获取验证码图片

    Public Sub Verify_Action()
	   '#验证码演示:
	   Set verify = loader.LoadClass("Ext/Verify")
	       verify.output
	   Set verify = Nothing
	   '##
    End Sub
	
	'@Mail_Action(): 发邮件测试

    Public Sub Mail_Action()
	   '#发邮件演示:
	   Set mail = loader.LoadClass("Ext/Mail")
	       mail.Setting "smtp.exmail.qq.com",465,1,"message@ekede.com","xxx123"
		   i = mail.Send("3002823478@qq.com", "Frank", "test subject", "test content", "ekede", "message@ekede.com", 1)
	   Set mail = Nothing
	   '##
       wts.responses.SetOutput i
    End Sub
	
	'@Spider_Action(): 发送HTTP请求

    Public Sub Spider_Action()
	    '#Http请求演示:
		url = wts.route.ReWrite(wts.site.config("base_url"), "index.asp?route=index/http")
		Set http = loader.LoadClass("Ext/Http")
		With http
		     .cookie_on = True
			 .cache = wts.cache
			'.SetHeader "Accept-Encoding", "gzip, deflate"'设定gzip压缩
			 .SetHeader "Referer", url '设定页面来源
			 .SetHeader "Accept-Language", "zh-cn" '设定语言
			 .SetHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" '设定浏览器
			 .SetHeader "Accept", "*/*" '文档类型
			 .SetHeader "aaa", "bbb" '自定义头
			 .SetHeader "If-Modified-Since", "0" '避免跳转错误
		     .SetHeader "Content-Type", "application/x-www-form-urlencoded" '"multipart/form-data"
			 .AddItem "username","nom"
			 .AddItem "password","pass"
			'.items = "你好"
			 .Send "POST",url
			 '
			 str="status:"&.rStatus&"<br/>"
			 str=str&"header<textarea style='width:100%;height:200px'>"&.rHeader&"</textarea><br/>"
			 str=str&"cookie old<textarea  style='width:100%;height:100px'>"&.rCookie&"</textarea><br/>"
			 str=str&"body<textarea  style='width:100%;height:200px'>"&.rText&"</textarea>"
		End With
		Set http = Nothing
		'##
        wts.responses.SetOutput str
    End Sub
	
    Public Sub Http_Action()
	    'wts.requests.bytes
		'wts.requests.servers
		'request.querystring
        wts.responses.SetOutput request.form
    End Sub
	
	'@Stat_Action(): 网站统计

    Public Sub Stat_Action()
	    Dim stat,ip
		'#网站统计演示:
        Set stat = loader.LoadClass("Ext/WebStat")
		    stat.fun = wts.fun
		    ip=stat.GetIp()
			wts.responses.SetOutput stat.GetSys
			'wts.responses.SetOutput stat.GetBrowser()
			'wts.responses.SetOutput stat.GetLanguage()
		Set stat = Nothing
		'##
    End Sub
	
	'@IP_Action(): 读取IP地址

    Public Sub IP_Action()
	    Dim wry,ip
		'#IP库查询演示:
		ip="202.202.202.202"
		Set wry = loader.LoadClass("Ext/Tqqwry")
			wry.data=wts.fso.GetMapPath(PATH_ROOT&"data/db/qqwry.dat")
			ipType = wry.QQWry(ip)
			ipContry = wry.Country
		Set wry = Nothing
		wts.responses.SetOutput ipContry
		'##
    End Sub
	
	'@xml_Action(): xslt渲染xml
	
    Public Sub xml_Action()

		'xml+xslt渲染
		str1="<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
		str1=str1&"<title>This is Title</title>"
		'
		str2="<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
		str2=str2&"<html xsl:version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" xmlns=""http://www.w3.org/1999/xhtml"">"
		str2=str2&"<xsl:value-of select=""title""/>"
		str2=str2&"</html>"
        '
        str=wts.fun.XmlTrans(str1,str2)
        wts.responses.SetContentType="text/xml"
		
		'输出
        wts.responses.SetOutput str
    End Sub

	'@dom_Action(): xml dom操作
	
    Public Sub dom_Action()
	    '#dom操作演示:
		Set xml = loader.LoadClass("Ext/Xml")
		    '增加根节点
		    Set root = xml.CreateNew("root")
			'增加子节点
			Set first_ = xml.AddNode("first_",root)
			    xml.AddAttribute "id","1",first_
			'增加孙节点2
			Set second_ = xml.AddNode("second_",first_)
			    xml.AddAttribute "id","2",second_   '节点增加属性
			    xml.AddText "Cd*asdf&",True,second_ '节点添加文本
			'增加孙节点3
			Set third_ = xml.AddNode("third_",first_)
			    xml.AddAttribute "id","3",third_
				xml.AddText "3333",False,third_
			'查找节点返回数组
			Set findnode=xml.FindNodes("second_")
			    s = "length:"& findnode.length&chr(10)
			    If findnode.length > 0 Then
				   s = s & "id:"&xml.GetAtt("id",findnode(0))&chr(10)    '节点属性读取
				   s = s & "name:"& xml.GetNodeName(findnode(0))&chr(10) '节点名称读取
				   s = s & "text:"& xml.GetNodeText(findnode(0))&chr(10) '节点文本读取
				   s = s & "type:"& xml.GetNodeType(findnode(0))&chr(10) '节点类型读取
				End If
            '替换节点内容
			xml.ReplaceNode "root/first_/third_","444",True
			'删除节点
			xml.DelNode "root/first_/third_"
			'保存为xml
			xml.SaveAsXML wts.fso.GetMapPath(PATH_ROOT&PATH_DATA&"test.xml")
		Set xml = Nothing
		'##
		'输出
        wts.responses.SetOutput s
    End Sub

	'@Arr_Action(): 数组操作
	
    Public Sub Arr_Action()
	    '#数组排序演示:
        'a=Array(25,20,31,33,4,5,5,7,9,2,10)
        a=Array("zoo","a","big","som")
		Set Arr = loader.LoadClass("Ext/Array")
            Arr.Sorts a,"desc"
		Set Arr = Nothing
		For i = 0 To UBound(a)
		    s=s& a(i)&Chr(10)
		Next
		'##
		'输出
        wts.responses.SetOutput s
    End Sub
	
    '@Zip_Action(): ZIP打包解包
	
    Public Sub Zip_Action()
	    Dim zFolder,zFile,uFolder
	    '#zip打包演示:
		zFolder = wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&PATH_PIC_IMAGES)
		zFile = wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&"image.zip")
		uFolder = wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&"unzip") 

		Set fzip = loader.LoadClass("Ext/Zip")
		    fzip.Zip zFolder, zFile
			fzip.UnZip zFile,uFolder
        Set fzip = Nothing
		'##
		wts.responses.SetOutput "Zip OK"
    End Sub
	
	'@Pack_Action(): Pack打包解包
	
    Public Sub Pack_Action()
	    Dim zFolder,zFile,uFolder
	    '#pack打包演示:
		zFolder = wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&PATH_PIC_IMAGES)
		zFile = wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&"images.pack")
		uFolder = wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&"unpack")
		
		Set p = loader.LoadClass("Ext/Pack")
		    p.Pack zFolder, zFile
		    p.UnPack zFile,uFolder
		Set pack = Nothing
		'##
        wts.responses.SetOutput "Pack Ok"
    End Sub
	
	'@Cart_Action(): 购物车演示
	
    Public Sub Cart_Action()
	    
	    '#购物车:
		Set cart = loader.LoadClass("Ext/Cart")
			cart.cartId = wts.site.langId
			cart.add 1,1
			cart.add 2,1
			cart.add 3,1
			cart.Remove 1
            n = cart.HasNum
		Set cart = Nothing
		'##

        '输出内容
        wts.responses.SetOutput cstr(n)
	End Sub

End Class
%>