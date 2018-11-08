<%
'@title: Class_Ext_Mail
'@author: ekede.com
'@date: 2017-02-13
'@description: 邮件类 CDO.Message

Class Class_Ext_Mail

    Private isDebug_
    Private mailObject_
    Private objConfig_
    Private fields_
	
    '@isDebug: 是否设置为调试模式
	
    Public Property Let isDebug(Value) 
        isDebug_ = Value
    End Property

    Private Sub Class_Initialize
        On Error Resume Next
		If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
		'
        Set mailObject_ = Server.CreateObject("CDO.Message")
        Set objConfig_ = Server.CreateObject ("CDO.Configuration")
        Set fields_ = objConfig_.fields
		'
        If Err Then OutErr("No CDO.Message:"&Err.Description)
    End Sub

    Private Sub Class_Terminate
        Set fields_ = Nothing 
        Set objConfig_ = Nothing
        Set mailObject_ = Nothing
    End Sub
	
    '@Setting(mServer,mPort,mSSL,mUserName,mPassword): 配置服务器

    Public Function Setting(mServer,mPort,mSSL,mUserName,mPassword)
		With fields_
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '使用网络服务器还是本地服务
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mServer '服务器地址
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = mPort '465谷歌端口,正常25器
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 '服务器认证方式
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = mSSL '是否使用SSL 1或true为启用
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = mUserName '发件人邮箱
			.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mPassword '发人邮箱密码
			.Item("http://schemas.microsoft.com/cdo/configuration/languagecode") = "UTF-8"
			.Update()
		End With 
		Set mailObject_.Configuration = objConfig_
    End Function

    '@Send(toMail, toName, subject, body, fromName, fromMail, priority): 发送邮件

    Public Function Send(toMail, toName, subject, body, fromName, fromMail, priority)
        On Error Resume Next
		Send = True
        If fromName <> "" Then
		   fm = """" & fromName & """ <" & Trim(fromMail) & ">"
		Else
		   fm = fromMail
		End If
		'
		With mailObject_
			.Subject = subject
			.From = fm
			.To = toMail
			.HTMLBody = body     'HTML 網頁格式信件
		   '.CC = strYouEmail   '副本
		   '.BCC = strYouEmail  '密件副本
		   '.TextBody = "信件內容" '文字格式信件內容
		   '.AddAttachment(http://xxxxxx/xxxx.xxx) '或者其他任何正确的url,包括http,ftp,file等等。
		    .Send
		End With 
        '
        If  Err Then
            Send = False
            OutErr("Send Mail Fail:"&Err.Description)
        End If
    End Function

    'Err

    Private Sub OutErr(str)
		Err.clear
        If IsDebug_ = true Then
            Response.charset = "utf-8"
            Response.Write str
            Response.End
        End If
    End Sub

End Class
%>