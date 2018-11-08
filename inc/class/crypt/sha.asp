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
	
	'@SHA1(ByVal Str): SHA1

	Function SHA1(ByVal Str)
		Dim Enc,Bytes,objXML,objXMLNode,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
			'Convert the string to a byte array and hash it
			Bytes = TAsc.GetBytes_4(Str)
			Bytes = Enc.ComputeHash_2((Bytes))
			SHA1 = Bytes
		Set Enc = Nothing
	End Function
	
	'@SHA256(ByVal Str): SHA256
	
	Function SHA256(ByVal Str)
		Dim Enc,Bytes,objXML,objXMLNode,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.SHA256Managed")
			'Convert the string to a byte array and hash it
			Bytes = TAsc.GetBytes_4(Str)
			Bytes = Enc.ComputeHash_2((Bytes))
			SHA256 = Bytes
		Set Enc = Nothing
	End Function
	
	'@SHA512(ByVal Str): SHA512
	
	Function SHA512(ByVal Str)
		Dim Enc,Bytes,objXML,objXMLNode,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.SHA512Managed")
			'Convert the string to a byte array and hash it
			Bytes = TAsc.GetBytes_4(Str)
			Bytes = Enc.ComputeHash_2((Bytes))
			SHA512 = Bytes
		Set Enc = Nothing
	End Function
	
	'@HMACSHA1(ByVal Str,ByVal Key): HMACSHA1
	
	Function HMACSHA1(ByVal Str,ByVal Key)
		Dim Enc,Bytes
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.HMACSHA1")
			'Convert the string to a byte array and hash it
			Enc.Key = TAsc.GetBytes_4(Key)
			Bytes = TAsc.GetBytes_4(Str)
			Bytes = Enc.ComputeHash_2((Bytes))
			HMACSHA1 = Bytes
		Set Enc = Nothing
	End Function
	
	'@HMACSHA256(ByVal Str,ByVal Key): HMACSHA256
	
	Function HMACSHA256(ByVal Str,ByVal Key)
		Dim Enc,Bytes
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.HMACSHA256")
			'Convert the string to a byte array and hash it
			Enc.Key = TAsc.GetBytes_4(Key)
			Bytes = TAsc.GetBytes_4(Str)
			Bytes = Enc.ComputeHash_2((Bytes))
			HMACSHA256 = Bytes
		Set Enc = Nothing
	End Function
	
	'@HMACSHA512(ByVal Str,ByVal Key): HMACSHA512
	
	Function HMACSHA512(ByVal Str,ByVal Key)
		Dim Enc,Bytes
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.HMACSHA512")
			'Convert the string to a byte array and hash it
			Enc.Key = TAsc.GetBytes_4(Key)
			Bytes = TAsc.GetBytes_4(Str)
			Bytes = Enc.ComputeHash_2((Bytes))
			HMACSHA512 = Bytes
		Set Enc = Nothing
	End Function
	
End Class
%>