<%
'@title: Model_Hello
'@author: ekede.com
'@date: 2017-02-23
'@description: 模型演示

'#模型演示:
Class Model_Hello

	private t_fields
	private t_join
	private t_where
	private t_order

    '@tfields: 主要字段
	
	Public Property Get tfields
		tfields=t_fields
	End Property
	
    '@tjoin: 表连接

	Public Property Get tjoin
		tjoin=t_join
	End Property
	
    '@twhere: 条件

	Public Property Get twhere
		twhere=t_where
	End Property
	
    '@torder: 排序

	Public Property Get torder
		torder=t_order
	End Property
	
	Private sub Class_Initialize
	   t_fields= ""
	   t_join= DB_PRE&"hello"
	   t_where= ""
	   t_order= ""
	End Sub
	Private Sub Class_Terminate() 
	End Sub 
	
	'@getHello(): 取所有
	
	Public Function GetAll()
	    Set getAll=wts.db.Query(t_join,"","","","",1,1)
	End Function
	
	'@getNameById(id): 取单条
	
	Public Function GetNameById(id)
	    Set getNameById=wts.db.Query(t_join,"","id="&id,"","",1,1)
	End Function
	
	'@Add(name): 添加
	
	Public Function Add(name)
	    Add=wts.db.Add(t_join, "name", "'"&name&"'")
	End Function
	
	'@Edit(data): 修改
	
	Public Function Edit(data)
	    Edit=wts.db.Edit(t_join, "name", "'"&data("name")&"'","id="&data("id"))
	End Function
	
	'@Del(id): 删除
	
	Public Function Del(id)
	    Del=wts.db.Del(t_join,"id="&id)
	End Function

End Class
'##
%>