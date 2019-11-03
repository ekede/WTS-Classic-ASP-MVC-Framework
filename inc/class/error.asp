<%
'@title: Class_Error
'@author: ekede.com
'@date: 2018-06-18
'@description: 错误信息类

Class Class_Error

    '@foundErr: 是否有错误
	
    Dim foundErr
    Private errMsg_
    Private loader_
	
    '@loader: Loader对象依赖

    Public Property Let loader(Value)
        Set loader_ = Value
    End Property

    Private Sub Class_Initialize()
	    foundErr = False
    End Sub

    Private Sub class_terminate()
    End Sub

    '@AddMsg(msg): 添加错误

    Public Sub AddMsg(msg)
	    foundErr = True
		'
        If errMsg_ = "" Then
            errMsg_ = msg
        Else
            errMsg_ = errMsg_ & "|-|"&msg
        End If
    End Sub

    '@OutMsg(): 查看错误

    Public Sub OutMsg()
		Response.charset = "utf-8"
		Response.Write errMsg_
		Response.End
    End Sub

    '@GetMsg(): 返回错误

    Public Function GetMsg()
        Dim temp, aMsg, aNum, i
        aMsg = Split(errMsg_, "|-|")
        aNum = UBound(aMsg)
        For i = 0 To aNum
            If i > 0 Then Temp = Temp&","
            temp = temp&""""&aMsg(i)&""""
        Next
        '
        GetMsg = temp
    End Function

    '@Out(e): 转向404页

    Public Sub Out(e)
        loader_.LoadControlAction "Error", "E404"
    End Sub

End Class
%>