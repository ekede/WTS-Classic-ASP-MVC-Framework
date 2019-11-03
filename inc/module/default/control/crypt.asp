<%
'@title: Control_Crypt
'@author: ekede.com
'@date: 2018-06-09
'@description: 加密解密

Class Control_Crypt

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub

    '@Index_Action(): 
    
	Sub Index_Action()
        Call Test5
    End Sub
	
	'散列,哈西
	
    Private Sub Test0()
	    '#md5演示:
		set c = loader.loadClass("Ext/Md5")
		    wts.responses.SetOutput "md5(""你好"") : "& c.MD5("你好",32) '中文不一致的问题
		set c = nothing
		'##
    End Sub
	
    Private Sub Test1()
	    '#HMACMD5演示:
	    set h = loader.loadClass("Crypt/Hex")
        set c = loader.loadClass("Crypt/Md5")
			wts.responses.SetOutput h.Bytes2Hex(c.MD5("你好"))
			wts.responses.SetOutput h.Bytes2Hex(c.HMACMD5("你好","123"))
		set c = nothing
		set h = nothing
		'##
    End Sub
	
    Private Sub Test2()
	    '#HMACSHA1演示:
	    set h = loader.loadClass("Crypt/Hex")
        set c = loader.loadClass("Crypt/Sha")
			wts.responses.SetOutput h.Bytes2Hex(c.SHA1("你好"))
			wts.responses.SetOutput h.Bytes2Hex(c.HMACSHA1("你好","123"))
		set c = nothing
		set h = nothing
		'##
    End Sub
	
	'Base64,转码
	
    Private Sub Test3()
	    '#Base64演示:
        set c = loader.loadClass("Crypt/Base64")
			x = c.Bytes2Base64(wts.fso.Str2Bytes("Str,二进制,Base64转换","utf-8"))
			wts.responses.SetOutput wts.fso.Bytes2Str(c.Base642Bytes(x),"utf-8")
		set c = nothing
		'##
    End Sub
	
    Private Sub Test4()
        set c = loader.loadClass("Crypt/Escape")
			x = c.Escape("Escape,UnEscape函数")
			wts.responses.SetOutput c.UnEscape(x)
		set c = nothing
    End Sub
	
    Private Sub Test5()
        set c = loader.loadClass("Crypt/A2U")
			x = c.Encode("ASCII,UNICODE转换")
			wts.responses.SetOutput c.Decode(x)
		set c = nothing
    End Sub
	
    Private Sub Test6()
        set c = loader.loadClass("Crypt/Num")
			wts.responses.SetOutput c.DcH(30)
		set c = nothing
    End Sub
	
	Private Sub Test7()
		Set c = loader.loadClass("Crypt/UrlDecode")
			c.UrlDecode(server.URLEncode("url解码"))
		Set c = nothing
	End Sub
	
	
	'加密/解密-对称

    Private Sub Test8()
        set c = loader.loadClass("Crypt/Des")
			x = c.DESEncrypt("DES加密,解密","12345678")
			wts.responses.SetOutput c.DESDecrypt(x,"12345678")
		set c = nothing
    End Sub
	
    Private Sub Test9()
        set c = loader.loadClass("Crypt/Aes")
			x = c.AESEncrypt("AES加密,解密","12345678ABCDEFGH")
			wts.responses.SetOutput c.AESDecrypt(x,"12345678ABCDEFGH")
		set c = nothing
    End Sub
	
	'加密/解密-非对称
	
	Private Sub Test10()
	
		Dim LngKeyE 
		Dim LngKeyD 
		Dim LngKeyN 
		Dim ObjRSA 
        
		str1="RSA,123"
		
		Set ObjRSA = loader.loadClass("Crypt/Rsa")
		
		   'Generate Keys
			ObjRSA.GenKey() 
			LngKeyE = ObjRSA.PublicKey 
			LngKeyD = ObjRSA.PrivateKey 
			LngKeyN = ObjRSA.Modulus 
			
			'Encrypt
			ObjRSA.PublicKey = LngKeyE 
			ObjRSA.Modulus = LngKeyN 
			x = ObjRSA.Encode(str1)
			
			'Decrypt 
			ObjRSA.PrivateKey = LngKeyD 
			ObjRSA.Modulus = LngKeyN 
			str2 = ObjRSA.Decode(x) 
			
		Set ObjRSA = nothing
		
		wts.responses.SetOutput str2

	End Sub

	
End Class
%>