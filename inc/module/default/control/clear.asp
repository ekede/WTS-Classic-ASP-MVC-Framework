<%
'@title: Control_Clear
'@author: ekede.com
'@date: 2018-02-01
'@description: 清除缓存

Class Control_Clear

	'@Index_Action(): 
	
    Public Function Index_Action()
        loader.ClearApp()
        wts.cache.ClearValue()
        wts.responses.SetOutput "clear application"
    End Function
	
	'@View_Action(): 

    Public Function View_Action()
        wts.responses.SetOutput loader.ViewApp()
    End Function

End Class
%>