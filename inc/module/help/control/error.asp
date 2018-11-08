<%
'@title: Control_Error
'@author: ekede.com
'@date: 2018-02-01
'@description: Error

Class Control_Error

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
        wts.responses.outputs
        wts.responses.Die("")
    End Sub
	
    '@E404_Action(): 404错误

    Public Sub E404_Action()
        wts.responses.setStatus = wts.responses.getStatus(404)
        wts.responses.SetOutput "this is 404"
    End Sub
	
    '@E405_Action(): 405错误

    Public Sub E405_Action()
        wts.responses.SetOutput "this is 405"
    End Sub
	
    '@E500_Action(): 500错误

    Public Sub E500_Action()
        wts.responses.setStatus = wts.responses.getStatus(500)
        wts.responses.SetOutput "this is 500"
    End Sub

End Class
%>