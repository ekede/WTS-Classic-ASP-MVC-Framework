<%
'@title: Control_Json
'@author: ekede.com
'@date: 2018-06-09
'@description: json演示

Class Control_Json

    Private Sub Class_Initialize()
        loader.IncludeG PATH_CLASS&"Ext/Json"
    End Sub

    Private Sub Class_Terminate()
    End Sub
	
	'@Index_Action(): 

    Public Sub Index_Action()
        Call Test2()
    End Sub
	
    '转Json串
		
    Private Sub Test1()
        '#生成json串:
        Set jj = New Class_Ext_Json
			jj.setKind="object"
			jj(null)="a"
			jj(null)="b"
			jj(null)="c"
			jj(null)="d"
			jj("b")="g"
		Set jj("a")= New Class_Ext_Json
			jj("a").setKind="array"
			jj("a")(null)="e"
			jj("a")(null)="f"
		    str=jj.ToString
		Set jj = nothing
		'##
        wts.responses.SetContentType="application/json"
        wts.responses.SetOutput str

    End Sub
	
	'rs转json串
	
	Public Sub Test2()
        '#rs转json串:
        Set mHello = loader.LoadModel("Hello")
        Set rs = mHello.getAll
		Set jsa = New Class_Ext_Json
		jsa.setKind = "array"
		While Not (rs.EOF Or rs.BOF)
			Set jsa(Null) = New Class_Ext_Json
			jsa(Null).setKind = "object"
			For Each col In rs.Fields
				jsa(Null)(col.Name) = col.Value
			Next
			rs.MoveNext
		Wend
		str=jsa.ToString
		Set jsa=nothing
		rs.close
		set rs = nothing
		set mHello = nothing
		'##
        wts.responses.SetContentType="application/json"
        wts.responses.SetOutput str
		
	End Sub
	
	'可增强
	
    Private Sub Test4()
        '#json串解析:
		str="{""a"":""1"",""b"":""2"",""c"":""3"",d:[5,{a1:{a1:11,a2:22222,a3:33}},7,8,9],e:{f:10,g:11}}"
		'
		Set jt = loader.LoadClass("Ext/JsonT")
		Set jo = jt.getJSONObject(str)
		    wts.responses.SetOutput jt.getJSArrayItem(jo.d,1).a1.a2
		Set jo = Nothing
		Set jt = Nothing
		'##
    End Sub
	
End Class
%>