<%
'@title: Class_Crypt_Aes
'@author: ekede.com
'@date: 2017-02-13
'@description: Aes加密解密

Class Class_Crypt_Aes

    Private TAsc

    Private Sub Class_Initialize()
        Set TAsc = Server.CreateObject("System.Text.UTF8Encoding")
    End Sub

    Private Sub Class_Terminate()
	    Set TAsc = Nothing
    End Sub

    '@AESEncrypt(ByVal Str,ByVal Key): 加密

	'Mode 1 : cbc , 2 : ecb , 4 : cfb
	'Padding 2 : pkcs5 , 4 : ansix923
	Public Function AESEncrypt(ByVal Str,ByVal Key)
		Dim Enc,BytesText,Bytes,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.RijndaelManaged")
		'Convert the string to a byte array and hash it
		Enc.Mode = 2
		Enc.Padding = 2
		Enc.IV = TAsc.GetBytes_4(Key)
		Enc.Key = TAsc.GetBytes_4(Key)
		BytesText = TAsc.GetBytes_4(Str)
		Bytes = Enc.CreateEncryptor().TransformFinalBlock((BytesText),0,Lenb(BytesText))
		'Convert the byte array to a hex or bsae64 string
		AESEncrypt = Bytes
		Set Enc = Nothing
	End Function

    '@AESDecrypt(ByVal Bytes,ByVal Key): 解密

	'Mode 1 : cbc , 2 : ecb , 4 : cfb
	'Padding 2 : pkcs5 , 4 : ansix923
	Public Function AESDecrypt(ByVal Bytes,ByVal Key)
		Dim Enc,BytesText,Outstr
		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set Enc = Server.CreateObject("System.Security.Cryptography.RijndaelManaged")
		'Convert the string to a byte array and hash it
		Enc.Mode = 2
		Enc.Padding = 2
		Enc.IV = TAsc.GetBytes_4(Key)
		Enc.Key = TAsc.GetBytes_4(Key)
		'Convert the byte array to a hex or bsae64 string
		Outstr = Enc.CreateDecryptor().TransformFinalBlock((Bytes),0,Lenb(Bytes))
		AESDecrypt = TAsc.GetString((Outstr))
		Set Enc = Nothing
	End Function

End Class
%>