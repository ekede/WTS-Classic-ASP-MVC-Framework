<%
'@title: Class_Crypt_Sha
'@author: ekede.com
'@date: 2017-02-13
'@description: SHA,HMACSHA加密

Class Class_Crypt_Sha

    Private TAsc

    Private Sub Class_Initialize()
        Set TAsc = Server.CreateObject("System.Text.UTF8Encoding")
    End Sub

    Private Sub Class_Terminate()
	    Set TAsc = Nothing
    End Sub
	
	'@SHA1(ByRef Str): SHA1

	Function SHA1(ByRef Str)
		Dim Enc,Bytes,objXML,objXMLNode,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
			'Convert the string to a byte array and hash it
			Bytes = TAsc.GetBytes_4(Str)
			SHA1 = Enc.ComputeHash_2((Bytes))
		Set Enc = Nothing
	End Function
	
	'@SHA256(ByRef Str): SHA256
	
	Function SHA256(ByRef Str)
		Dim Enc,Bytes,objXML,objXMLNode,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.SHA256Managed")
			'Convert the string to a byte array and hash it
			Bytes = TAsc.GetBytes_4(Str)
			SHA256 = Enc.ComputeHash_2((Bytes))
		Set Enc = Nothing
	End Function
	
	'@SHA512(ByRef Str): SHA512
	
	Function SHA512(ByRef Str)
		Dim Enc,Bytes,objXML,objXMLNode,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.SHA512Managed")
			'Convert the string to a byte array and hash it
			Bytes = TAsc.GetBytes_4(Str)
			SHA512 = Enc.ComputeHash_2((Bytes))
		Set Enc = Nothing
	End Function
	
	'@HMACSHA1(ByRef Str,ByRef Key): HMACSHA1
	
	Function HMACSHA1(ByRef Str,ByRef Key)
		Dim Enc,Bytes
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.HMACSHA1")
			'Convert the string to a byte array and hash it
			Enc.Key = TAsc.GetBytes_4(Key)
			Bytes = TAsc.GetBytes_4(Str)
			HMACSHA1 = Enc.ComputeHash_2((Bytes))
		Set Enc = Nothing
	End Function
	
	'@HMACSHA256(ByRef Str,ByRef Key): HMACSHA256
	
	Function HMACSHA256(ByRef Str,ByRef Key)
		Dim Enc,Bytes
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.HMACSHA256")
			'Convert the string to a byte array and hash it
			Enc.Key = TAsc.GetBytes_4(Key)
			Bytes = TAsc.GetBytes_4(Str)
			HMACSHA256 = Enc.ComputeHash_2((Bytes))
		Set Enc = Nothing
	End Function
	
	'@HMACSHA512(ByRef Str,ByRef Key): HMACSHA512
	
	Function HMACSHA512(ByRef Str,ByRef Key)
		Dim Enc,Bytes
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.HMACSHA512")
			'Convert the string to a byte array and hash it
			Enc.Key = TAsc.GetBytes_4(Key)
			Bytes = TAsc.GetBytes_4(Str)
			HMACSHA512 = Enc.ComputeHash_2((Bytes))
		Set Enc = Nothing
	End Function
	
End Class
%>