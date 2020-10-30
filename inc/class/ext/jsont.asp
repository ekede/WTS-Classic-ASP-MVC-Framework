<%
'@title: Class_Ext_JsonT
'@author: ekede.com
'@date: 2017-11-29
'@description: Str To Json Object

Class Class_Ext_JsonT

	Dim sc4Json
	
    Private Sub Class_Initialize()
		Set sc4Json = Server.CreateObject("MSScriptControl.ScriptControl")    
		sc4Json.Language = "JavaScript"    
		sc4Json.AddCode "var itemTemp=null;" 
		sc4Json.AddCode "function getJSArray(arr, index){itemTemp=arr[index];}" 
    End Sub
	
    Private Sub Class_Terminate()
	    Set sc4Json = nothing
    End Sub
	
	'@GetJSONObject(ByRef strJSON): Json字符串转对象
	
	Function GetJSONObject(ByRef strJSON)
		sc4Json.AddCode "var jsonObject = " & strJSON    
		Set getJSONObject = sc4Json.CodeObject.jsonObject    
	End Function 
	
	'@GetJSArrayItem(ByRef objJSArray,ByRef indexs): 数组对象索引取值
	
	Function GetJSArrayItem(ByRef objJSArray,ByRef indexs)
		On Error Resume Next    
		sc4Json.Run "getJSArray",objJSArray, indexs
		Set GetJSArrayItem = sc4Json.CodeObject.itemTemp    
		If Err.number=0 Then Exit Function    
		GetJSArrayItem = sc4Json.CodeObject.itemTemp    
	End Function 
	
End Class
%>