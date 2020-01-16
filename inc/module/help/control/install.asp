<%
'@title: Control_Install
'@author: ekede.com
'@date: 2018-02-01
'@description: 安装目录数据

Class Control_Install

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub

    '@Index_Action(): 安装数据目录,数据库

    Public Sub Index_Action()

	    '安装目录
	    Install_Folder()

	    '安装Hello数据库
	    Install_Access()

		'输出
        wts.responses.SetOutput "Install OK"

	End Sub

	'安装目录

    Private Sub Install_Folder()

	    '判断数据路径是否存在并创建
		wts.fso.CreateFolders wts.fso.GetMapPath(PATH_ROOT&PATH_DATA)

		'判断图片路径是否存在并创建
		wts.fso.CreateFolders wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&PATH_PIC_IMAGES)

		'判断默认图片是否存在并拷贝
		d_pic="no.gif"
        If wts.fso.GetRealPath(PATH_ROOT&PATH_PIC&PATH_PIC_IMAGES&d_pic)= -1 Then
		   If  wts.fso.GetRealPath(PATH_MODULE&wts.route.module&"/"&PATH_VIEW&PATH_PIC_IMAGES&d_pic) <> -1 Then
			   wts.fso.CopyAFile _
			   wts.fso.GetMapPath(PATH_MODULE&wts.route.module&"/"&PATH_VIEW&PATH_PIC_IMAGES&d_pic),_
			   wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&PATH_PIC_IMAGES&d_pic)
		   End If
		End If

	    '判断APP路径是否存在并创建
		If PATH_APP <> "" Then wts.fso.CreateFolders wts.fso.GetMapPath(PATH_ROOT&PATH_APP)

    End Sub

	'安装数据库

    Private Sub Install_Access()

	    '#安装压缩Access演示:

	    d_name = DB_NAME
		c_name = "backup.mdb"
      
        '判断数据库是否存在
        If DB_TYPE<>1 or wts.fso.GetRealPath(PATH_ROOT&DB_PATH&d_name)<> -1 Then Exit Sub
		
		'判断数据库路径是否存在并创建
		wts.fso.CreateFolders wts.fso.GetMapPath(PATH_ROOT&DB_PATH)

	    Set db = loader.loadClass("db")

		    '创建数据库
		    db.CreateAccess DB_PATH&d_name

			'连接数据库
            db.OpenConn 1,DB_VERSION, DB_PATH, d_name, "", ""

			'创建表Hello
			sql="create table wts_hello ( "&_
			"id integer IDENTITY(1,1) primary key, "&_
			"name varchar(50), "&_
			"times date Default now() "&_
			")"
			db.SqlExecute(sql)

			'插入一条记录
			sql="insert into wts_hello (id,name) values (1,'example name')"
			db.SqlExecute(sql)

			'创建表mytable
			'm_id 自动编号字段并制作主键
			'm_class 文本型，长度50，非空，默认值：AAA
			'm_int 数字，长整型，非空
			'm_number 数字,小数，精度6，数值范围2
			'm_money 0.00货币，必添字段（非空）,默认0
			'm_memo text备注
			'm_date 日期/时间，date()默认当前日期（年月日）, datetime数据类型则对应 now()
			sql="create table mytable ( "&_
			"m_id integer IDENTITY(1,1) primary key, "&_
			"m_class varchar(50) NOT NULL Default 'AAA', "&_
			"m_int integer NOT NULL, "&_
			"m_numeric NUMERIC(6,2), "&_
			"m_money money NOT NULL Default 0.00, "&_
			"m_memo text, "&_
			"m_date date Default date(), "&_
			"m_boolean bit Default yes, "&_
			"m_blob OLEObject, "&_
			"m_double double, "&_
			"m_float real "&_
			")"
			'db.SqlExecute(sql)

			'增加字段
			sql="alter table mytable add column address varchar(200)"
			'db.SqlExecute(sql)

			'修改字段
			sql="alter table mytable Alter column address varchar(50)"
			'db.SqlExecute(sql)

			'删除字段
			sql="alter table mytable drop address"
			'db.SqlExecute(sql)

			'删除表
			sql="Drop table mytable"
			'db.SqlExecute(sql)

			'关闭数据库连接
			db.CloseConn

			'压缩备份数据库
		    db.CompactAccess DB_PATH&d_name,DB_PATH&c_name

		Set db = Nothing

		'##

    End Sub

	'@Space_Action(): 服务器组件
	
    Public Sub Space_Action()

	    '#常用组件:
		Dim theInstalledObjects(30)
		'危险
		theInstalledObjects(0) = array("WScript.Shell","wshom.ocx","danger")
		theInstalledObjects(1) = array("WScript.Network","wshom.ocx","danger")
		theInstalledObjects(2) = array("Shell.Application","shell32.dll","danger")
        '内置
		theInstalledObjects(3) = array("MSWC.AdRotator","adrot.dll","")
		theInstalledObjects(4) = array("MSWC.BrowserType","Browsercap.dll","")
		theInstalledObjects(5) = array("MSWC.NextLink","mswc.dll","")
		theInstalledObjects(6) = array("MSWC.Tools","tools.dll","")
		theInstalledObjects(7) = array("MSWC.Status","status.dll","")
		theInstalledObjects(8) = array("MSWC.Counters","counters.dll","")
		theInstalledObjects(9) = array("MSWC.PermissionChecker","permchk.dll","")
		'必要
		theInstalledObjects(10) = array("ADOX.Catalog","msadox.dll","")
		theInstalledObjects(11)= array("JRO.JetEngine","msjro.dll","")
		theInstalledObjects(12)= array("ADODB.Connection","msado15.dll","")
		theInstalledObjects(13)= array("ADODB.Stream","scrrun.dll","")
		theInstalledObjects(14)= array("Scripting.FileSystemObject","scrrun.dll","")
		theInstalledObjects(15)= array("Scripting.Dictionary","scrrun.dll","")
		'邮件
		theInstalledObjects(16)= array("CDO.Message","cdosys.dll","")
		theInstalledObjects(17)= array("JMail.Message","jmail.dll","")
		'图片
		theInstalledObjects(18)= array("WIA.ImageFile","wiaaut.dll","")
		theInstalledObjects(19)= array("Persits.Jpeg"," aspjpeg.dll","")
		'压缩
		theInstalledObjects(20)= array("Dyy.Zipsvr","dyy.dll","")
		'XML
		theInstalledObjects(21)= array("Microsoft.XMLDOM","msxml.dll","")
		theInstalledObjects(22)= array("MSXML2.DOMDocument","","")
		theInstalledObjects(23)= array("MSXML2.DOMDocument.3.0","msxml3.dll","")
		theInstalledObjects(24)= array("MSXML2.DOMDocument.4.0","","")
		theInstalledObjects(25)= array("MSXML2.DOMDocument.5.0","","")
		theInstalledObjects(26)= array("MSXML2.DOMDocument.6.0","msxml6.dll","")
		'HTTP
		theInstalledObjects(27)= array("MSXML2.ServerXMLHTTP","msxml2.dll","")
        '引擎
		theInstalledObjects(28)= array("MSScriptControl.ScriptControl","","")
		'应用
		theInstalledObjects(29)= array("InternetExplorer.Application","","")
		theInstalledObjects(30)= array("Excel.Application","","")
		'##
		'生成表格
		str="<table border=1>"
		str=str&"<tr><td>组件名称</td><td>支持</td><td>版本</td><td>DLL</td><td>说明</td></tr>"
		For i=0 to ubound(theInstalledObjects)
			If theInstalledObjects(i)(0)<>"" then
				str=str&"<tr class=tr_southidc>" 
				str=str&"<td>"& theInstalledObjects(i)(0) & "</td>"
				version=IsObjInstalled(theInstalledObjects(i)(0))
				If version = "" Then
				   str=str&"<td><b>×</b></td>"
				   str=str&"<td></td>"
				Else
				   str=str&"<td><b>√</b></td>" 
				   str=str&"<td>"&version&"</td>"
				End If
				str=str&"<td>" & theInstalledObjects(i)(1) & "</td>"
				str=str&"<td>"& theInstalledObjects(i)(2) & "</td>"
				str=str&"</tr>" & vbCrLf
			End If
		Next
		str=str&"</table>"
		
        '输出内容
        wts.responses.SetOutput str
		
    End Sub

	Private Function IsObjInstalled(strClass)
		On Error Resume Next
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClass)
		If Err Then
		   Err.Clear
		Else
		   IsObjInstalled = xTestObj.Version
		   If Err Then 
		      IsObjInstalled = "-"
		      Err.Clear
		   End If
		   Set xTestObj = Nothing
		End If
	End Function

End Class
%>