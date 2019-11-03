<%
'@title: Control_Clear
'@author: ekede.com
'@date: 2018-02-01
'@description: 清除缓存

Class Control_Clear

	'@Index_Action(): 
	
    Public Sub Index_Action()
        loader.ClearApp()
        wts.cache.ClearValue()
        wts.responses.SetOutput "clear application"
    End Sub
	
	'@View_Action(): 

    Public Sub View_Action()
        wts.responses.SetOutput loader.ViewApp()
    End Sub

End Class
%>