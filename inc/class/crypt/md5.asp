<%
'@title: Class_Crypt_Md5
'@author: ekede.com
'@date: 2017-02-13
'@description: MD5加密支持中文

Class Class_Crypt_Md5

    Private TAsc

    Private Sub Class_Initialize()
        Set TAsc = Server.CreateObject("System.Text.UTF8Encoding")
    End Sub

    Private Sub Class_Terminate()
	    Set TAsc = Nothing
    End Sub
	
	'@MD5(ByVal Str): MD5
	
	Public Function MD5(ByRef Str)
		Dim Enc,Bytes,objXML,objXMLNode,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
		'Convert the string to a byte array and hash it
		Bytes = TAsc.GetBytes_4(Str)
		MD5 = Enc.ComputeHash_2((Bytes))
		Set Enc = Nothing
	End Function
	
	'@HMACMD5(ByVal Str,ByVal Key): HMACMD5
	
	Public Function HMACMD5(ByRef Str,ByRef Key)
		Dim Enc,Bytes
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.HMACMD5")
		'Convert the string to a byte array and hash it
		Enc.Key = TAsc.GetBytes_4(Key)
		Bytes = TAsc.GetBytes_4(Str)
		HMACMD5 = Enc.ComputeHash_2((Bytes))
		Set Enc = Nothing
	End Function

End Class
%>