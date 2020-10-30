<% 
'@title: Class_Crypt_Rsa
'@author: ekede.com
'@date: 2020-10-28
'@description: RSA 公钥加密->私钥解密 , 私钥签名->公钥验签

Class Class_Crypt_Rsa

		Private TAsc,objRsa
		Private PrivateKey_,PublicKey_

        '@PrivateKey: Your personal private key.  Keep this hidden. Need C# format.

		Public Property Get PrivateKey
			PrivateKey = PrivateKey_
		End Property

		Public Property Let PrivateKey(Value)
			PrivateKey_ = Value
			objRsa.FromXmlString (PrivateKey_)
            PublicKey_ = objRsa.ToXmlString(False)
		End Property

        '@PublicKey: Key for others to encrypt data with.

		Public Property Get PublicKey
			PublicKey = PublicKey_
		End Property

		Public Property Let PublicKey(Value)
			PublicKey_ = Value
			objRsa.FromXmlString (PublicKey_)
		End Property

		Private Sub Class_Initialize()
			Set TAsc = Server.CreateObject("System.Text.UTF8Encoding")
			Set objRsa = Server.CreateObject("System.Security.Cryptography.RSACryptoServiceProvider")
			CreateKey()
		End Sub

		Private Sub Class_Terminate()
			Set objRsa = Nothing
			Set TAsc = Nothing
		End Sub

		Public Sub CreateKey()
			PrivateKey_ = objRsa.ToXmlString(True)
			PublicKey_ = objRsa.ToXmlString(False)
		End Sub

		'@Encrypt(ByRef Str): 公钥加密

		Public Function Encrypt(ByRef Str)
			Dim Bytes
			Bytes = TAsc.GetBytes_4(Str)
			Encrypt = Bytes2Base64(RsaEncrypt((Bytes)))
		End Function

		Private Function RsaEncrypt(ByRef Bytes)
			RsaEncrypt = objRsa.Encrypt((Bytes),False)
		End Function

		'@Decrypt(ByRef Bytes): 私钥解密

		Public Function Decrypt(ByRef Str)
			Dim Bytes
			Bytes=RsaDecrypt(Base642Bytes(Str))
			Decrypt = TAsc.GetString((Bytes))
		End Function

		Private Function RsaDecrypt(ByRef Bytes)
			 RsaDecrypt = objRsa.Decrypt((Bytes), False)
		End Function

		'@SignData(ByRef Str,ByRef Hash): 私钥签名 MD5 SHA1 SHA256

		Public Function SignData(ByRef Str,ByRef Hash)
			Dim Bytes
			Bytes = TAsc.GetBytes_4(Str)
			SignData = Bytes2Base64(SignHash(Bytes, Hash))
		End Function

		Private Function SignHash(ByRef Bytes,ByRef Hash)
			Dim MapNameToOID
			If Hash="MD5" Then
				MapNameToOID = "1.2.840.113549.2.5"
				Bytes = Md5(Bytes)
				SignHash = objRsa.SignHash((Bytes),MapNameToOID)
			End If
			If Hash="SHA1" Then
				MapNameToOID = "1.3.14.3.2.26"
				Bytes = SHA1(Bytes)
				SignHash = objRsa.SignHash((Bytes),MapNameToOID)
			End If
			If Hash="SHA256" Then
				MapNameToOID = "2.16.840.1.101.3.4.2.1"
				Bytes = SHA256(Bytes)
				SignHash = objRsa.SignHash((Bytes),MapNameToOID)
			End If
		End Function

		'@VerifyData(ByRef str,ByRef Hash,ByRef StrSign): 公钥验签

		Public Function VerifyData(ByRef str,ByRef Hash,ByRef StrSign)
			Dim Bytes,BytesSign
			Bytes = TAsc.GetBytes_4(Str)
			BytesSign = Base642Bytes(StrSign)
			VerifyData = objRsa.VerifyData((Bytes),Hash,(BytesSign))
		End Function

		'Hash

		Public Function Md5(ByRef Bytes)
			Dim En
			Set En = Server.CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
			Md5 = En.ComputeHash_2((Bytes))
			Set En = Nothing
		End Function

		Public Function SHA1(ByRef Bytes)
			Dim En
			Set En = Server.CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
			SHA1 = En.ComputeHash_2((Bytes))
			Set En = Nothing
		End Function

		Public Function SHA256(ByRef Bytes)
			Dim En
			Set En = Server.CreateObject("System.Security.Cryptography.SHA256CryptoServiceProvider")
			SHA256 = En.ComputeHash_2((Bytes))
			Set En = Nothing
		End Function

		'Base64

		Public Function Base642Bytes(str)
			Dim objXML, objXMLNode
			Set objXML = Server.CreateObject("msxml2.domdocument")
			Set objXMLNode = objXML.createelement("b64")
				objXMLNode.datatype = "bin.base64"
				objXMLNode.text = str
				Base642Bytes = objXMLNode.nodetypedvalue
			Set objXMLNode = Nothing
			Set objXML = Nothing
		End Function

		Public Function Bytes2Base64(bytes)
			Dim objXML, objXMLNode
			Set objXML = Server.CreateObject("msxml2.domdocument")
			Set objXMLNode = objXML.createelement("b64")
				objXMLNode.datatype = "bin.base64"
				objXMLNode.nodetypedvalue = bytes
				Bytes2Base64 = objXMLNode.text
			Set objXMLNode = Nothing
			Set objXML = Nothing
		End Function

End Class
%> 