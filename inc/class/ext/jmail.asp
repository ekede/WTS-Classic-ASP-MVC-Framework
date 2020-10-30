<%
'@title: Class_Ext_JMail
'@author: ekede.com
'@date: 2017-02-13
'@description: 邮件类 JMAIL.Message

Class Class_Ext_JMail

    Private isDebug_
    Private mailObject_
	Private mailServer_
	
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
        Set mailObject_ = Server.CreateObject("JMAIL.Message")
		'
        If Err Then OutErr("No JMAIL.Message:"&Err.Description)
    End Sub

    Private Sub Class_Terminate
        Set mailObject_ = Nothing
    End Sub
	
    '@Setting(ByRef mServer,ByRef mPort,ByRef mSSL,ByRef mUserName,ByRef mPassword): 配置服务器

    Public Function Setting(ByRef mServer,ByRef mPort,ByRef mSSL,ByRef mUserName,ByRef mPassword)
		mailObject_.Charset="utf-8"                          '邮件编码
		mailObject_.silent=true
		mailObject_.ContentType = "text/html"                '邮件正文格式

        mailServer_ = mServer
	   'mailObject_.ServerAddress= mServer                   '用来发送邮件的SMTP服务器
		mailObject_.MailServerUserName = mUserName           '登录用户名
		mailObject_.MailServerPassWord = mPassword           '登录密码
		mailObject_.MailDomain = MailDomain                  '域名（如果用“name@domain.com”这样的用户名登录时，请指明domain.com
    End Function

    '@Send(ByRef toMail,ByRef toName,ByRef subject,ByRef body,ByRef fromName,ByRef fromMail,ByRef priority): 发送邮件

    Public Function Send(ByRef toMail,ByRef toName,ByRef subject,ByRef body,ByRef fromName,ByRef fromMail,ByRef priority)
        On Error Resume Next
		Dim er
		Send = True
		'
		With mailObject_			
			.AddRecipient toMail,toName    '收信人
			.Subject=subject               '主题
		   '.HMTLBody=body                '邮件正文（HTML格式）
			.Body=body                     '邮件正文（纯文本格式）
			.FromName=fromName             '发信人姓名
			.From = fromMail               '发信人Email
			.Priority=priority             '邮件等级，1为加急，3为普通，5为低级
			.Send(mailServer_)
			er =.ErrorMessage
		End With 
        '
        If  er <> "" Then
            Send = False
            OutErr("Send Mail Fail:"&er)
        End If
    End Function

    'Err

    Private Sub OutErr(ByRef str)
		Err.clear
        If IsDebug_ = true Then
            Response.charset = "utf-8"
            Response.Write str
            Response.End
        End If
    End Sub

End Class
%>