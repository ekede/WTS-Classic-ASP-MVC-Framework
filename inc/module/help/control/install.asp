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
            db.OpenConn 1, DB_PATH, d_name, "", ""

			'创建表Hello
			sql="create table wts_hello ( "&_
			"id integer IDENTITY(1,1) primary key, "&_
			"name varchar(50) "&_
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
		theInstalledObjects(0) = "WScript.Shell"               'wshom.ocx
		theInstalledObjects(1) = "Shell.Application"           'shell32.dll
        '内置
		theInstalledObjects(2) = "MSWC.AdRotator"              'adrot.dll
		theInstalledObjects(3) = "MSWC.BrowserType"            'Browsercap.dll
		theInstalledObjects(4) = "MSWC.NextLink"               'mswc.dll
		theInstalledObjects(5) = "MSWC.Tools"                  'tools.dll
		theInstalledObjects(6) = "MSWC.Status"                 'status.dll
		theInstalledObjects(7) = "MSWC.Counters"               'counters.dll 
		theInstalledObjects(8) = "MSWC.PermissionChecker"      'permchk.dll
		'必要
		theInstalledObjects(9) = "ADOX.Catalog"
		theInstalledObjects(10)= "JRO.JetEngine"
		theInstalledObjects(11)= "ADODB.Connection"            'msado15.dll 
		theInstalledObjects(12)= "ADODB.Stream"                'scrrun.dll
		theInstalledObjects(13)= "Scripting.FileSystemObject"  'scrrun.dll
		theInstalledObjects(14)= "Scripting.Dictionary"        'scrrun.dll
		'邮件
		theInstalledObjects(15)= "CDO.Message"                 'cdosys.dll
		theInstalledObjects(16)= "JMail.Message"               'jmail.dll
		'图片
		theInstalledObjects(17)= "WIA.ImageFile"               'wiaaut.dll
		theInstalledObjects(18)= "Persits.Jpeg"
		'压缩
		theInstalledObjects(19)= "Dyy.Zipsvr"                  'dyy.dll
		'XML
		theInstalledObjects(20)= "Microsoft.XMLDOM"            'msxml.dll
		theInstalledObjects(21)= "MSXML2.DOMDocument"
		theInstalledObjects(22)= "MSXML2.DOMDocument.3.0"      'msxml3.dll 
		theInstalledObjects(23)= "MSXML2.DOMDocument.4.0"
		theInstalledObjects(24)= "MSXML2.DOMDocument.5.0"
		theInstalledObjects(25)= "MSXML2.DOMDocument.6.0"      'msxml6.dll
		'HTTP
		theInstalledObjects(26)= "MSXML2.ServerXMLHTTP"        'msxml2.dll
        '引擎
		theInstalledObjects(27)= "MSScriptControl.ScriptControl"
		'应用
		theInstalledObjects(28)= "InternetExplorer.Application"
		theInstalledObjects(29)= "Excel.Application"
		'##
		'生成表格
		str="<table border=1>"
		str=str&"<tr><td>组件名称</td><td>支持及版本</td></tr>"
		For i=0 to ubound(theInstalledObjects)
			If theInstalledObjects(i)<>"" then
				str=str&"<TR class=tr_southidc><TD>" & theInstalledObjects(i) & "</td><td>"
				If Not wts.fun.IsObjInstalled(theInstalledObjects(i)) Then
				   str=str&"<font class='red'><b>×</b></font>"
				Else
				   str=str&"<b>√</b> " 
				End If
				str=str&"</td></TR>" & vbCrLf
			End If
		Next
		str=str&"</table>"
		
        '输出内容
        wts.responses.SetOutput str
		
    End Sub

End Class
%>