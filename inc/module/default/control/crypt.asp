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

		'#RSA 演示:
		'Example Key Format : PEM PKCS1 -> PEM PKCS8 -> C# Private key -> C# PublicKey
		privatekey_pem="MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC7wJL6+hjchXCbZCMkWDi7JThNf8f9wK4b7xSAavSqy9fo5SfEq7xgbXyEoYH/vL+cHCVzSZgfw9KQOZPZIZHnQszPZ7+vvIGS9rOXKS6XTOpVPAkcA6u1cNUrxBkwf+8XNE9mu+vVJvAIw+snlkZG198+ZP/sxWTzJZCj1eUCrYAHPrZrHA7UULpNlzc+mYT7ymvXz4CfMRqLRMKOrc1oLbxBavcsYkM798a1P9tbPS4Gtg+EIDF5jJawOiSJfQ1TSLFQHY+6Uaq+zO2OvFlhA+xkgiBnWUzP/jeut8caSu2Y54JLbn5T9uXN9lJEXpYNUiFPAErVwVFvi1WB3XTvAgMBAAECggEBALEYV0tmxjaTo4DfNoqcsH5OAEqRohnHOjNdEvCCcl/8QJ8ML7PB7cDi5RXRpeaMqgvdPLH/E/+6XQ3vUXb4xD/n8XodOWDRJUNzcji9/pV2Vn6pT0peaAOP93L92Gi389TmYZLc5Pk8biNGcbP4ejdufcPDzucN1kfHAiSXqBkZ4G/Fp8ImaG9EY2KdZ65cDUHPbx786oI4U/UYcQ2BYd5tjH0A2WVXna1Ok6Qz51gS5h8pen2ga24FZn5IuGgm9ZXVRjJXH16bmLz9Bxj7qHVmkyAUUwNQelvGnmpF0JKfPvs5yhyavMPqEnAxwwv3pGkNYnYFbvX7z387mWy/9jECgYEA9C9bAbM87egP+dLaA8IQA2lJEBG6b5pJidw40lm0E2Ey1v2NMCEhd3stEUSrS0QfVN9S0N0aPZmsBNkbm7P7nSviq53n2Qo/mEP06dhx7+MI1nKlmTrgqH0HvCYK7+55vxojZsuvj9E9Q7tE5KXJsKZ8syLrWbLPvrPPF4TVwjcCgYEAxNY0H25RbMJGFUaZ/a4b6/yPXECdFX7LeFxFAiJ4ds+zavenRMacm4MNjY91m90t7p0UZYvytk39YeX2/J6x0C2U/gQE3VS5ER+NAOqrl9UgBxZeb0e5Cz3TcU4w/zT+sQedYqG4p/ldT4UnBKXleI/+l6H86Qnix1O9Xae25wkCgYB3jpgohPHYKj9oOmy0Wlgs02gKjiOScSCAd2r60yDwPC8ARLTUU+Rm89BlHBIikAAnNhD+YsNuVcd7uDFkUwNnOQ2KqY3THsl0bBGGTYu7wJWbKhcap1FILa+T16yTPVgu0UV0F1amO/SbLR3WNbZC38E+lGJXUM2WucMz6L4gkQKBgQCa4OzsWlJpYEfiz8W1LP09Z2GqNhEj67vP/dIyxsrAudcz8J/F5v0tBCZy35GrzZIpsaFt8XtN5PndwSPhTEEfS+5zHNhzCwn/pjK9qOjRtFnaGci+iNHaPZCVE/BLrvhEdXhqNlPkn7rDKkM0ThDMF4k86LHm7+dn7cUP3zp0eQKBgCj5Tcne23U8C93ifTM0mzlx3VEilL41lbS3pIiABiV+Cjk/e9YqYmEdkwwCk8g2mLmBYmzRCnCTCJbEOaLu4YPI2v1qbgo3WcTpodt2x7XskAPC8i4Kb7I9b1kMvFXQlxMlLGY4uz7JKSzp58ja5dFV2b4r1KlEd6x4ILF9OT4E"
		privatekey_csharp="<RSAKeyValue><Modulus>u8CS+voY3IVwm2QjJFg4uyU4TX/H/cCuG+8UgGr0qsvX6OUnxKu8YG18hKGB/7y/nBwlc0mYH8PSkDmT2SGR50LMz2e/r7yBkvazlykul0zqVTwJHAOrtXDVK8QZMH/vFzRPZrvr1SbwCMPrJ5ZGRtffPmT/7MVk8yWQo9XlAq2ABz62axwO1FC6TZc3PpmE+8pr18+AnzEai0TCjq3NaC28QWr3LGJDO/fGtT/bWz0uBrYPhCAxeYyWsDokiX0NU0ixUB2PulGqvsztjrxZYQPsZIIgZ1lMz/43rrfHGkrtmOeCS25+U/blzfZSRF6WDVIhTwBK1cFRb4tVgd107w==</Modulus><Exponent>AQAB</Exponent><P>9C9bAbM87egP+dLaA8IQA2lJEBG6b5pJidw40lm0E2Ey1v2NMCEhd3stEUSrS0QfVN9S0N0aPZmsBNkbm7P7nSviq53n2Qo/mEP06dhx7+MI1nKlmTrgqH0HvCYK7+55vxojZsuvj9E9Q7tE5KXJsKZ8syLrWbLPvrPPF4TVwjc=</P><Q>xNY0H25RbMJGFUaZ/a4b6/yPXECdFX7LeFxFAiJ4ds+zavenRMacm4MNjY91m90t7p0UZYvytk39YeX2/J6x0C2U/gQE3VS5ER+NAOqrl9UgBxZeb0e5Cz3TcU4w/zT+sQedYqG4p/ldT4UnBKXleI/+l6H86Qnix1O9Xae25wk=</Q><DP>d46YKITx2Co/aDpstFpYLNNoCo4jknEggHdq+tMg8DwvAES01FPkZvPQZRwSIpAAJzYQ/mLDblXHe7gxZFMDZzkNiqmN0x7JdGwRhk2Lu8CVmyoXGqdRSC2vk9eskz1YLtFFdBdWpjv0my0d1jW2Qt/BPpRiV1DNlrnDM+i+IJE=</DP><DQ>muDs7FpSaWBH4s/FtSz9PWdhqjYRI+u7z/3SMsbKwLnXM/Cfxeb9LQQmct+Rq82SKbGhbfF7TeT53cEj4UxBH0vucxzYcwsJ/6Yyvajo0bRZ2hnIvojR2j2QlRPwS674RHV4ajZT5J+6wypDNE4QzBeJPOix5u/nZ+3FD986dHk=</DQ><InverseQ>KPlNyd7bdTwL3eJ9MzSbOXHdUSKUvjWVtLekiIAGJX4KOT971ipiYR2TDAKTyDaYuYFibNEKcJMIlsQ5ou7hg8ja/WpuCjdZxOmh23bHteyQA8LyLgpvsj1vWQy8VdCXEyUsZji7PskpLOnnyNrl0VXZvivUqUR3rHggsX05PgQ=</InverseQ><D>sRhXS2bGNpOjgN82ipywfk4ASpGiGcc6M10S8IJyX/xAnwwvs8HtwOLlFdGl5oyqC908sf8T/7pdDe9RdvjEP+fxeh05YNElQ3NyOL3+lXZWfqlPSl5oA4/3cv3YaLfz1OZhktzk+TxuI0Zxs/h6N259w8PO5w3WR8cCJJeoGRngb8WnwiZob0RjYp1nrlwNQc9vHvzqgjhT9RhxDYFh3m2MfQDZZVedrU6TpDPnWBLmHyl6faBrbgVmfki4aCb1ldVGMlcfXpuYvP0HGPuodWaTIBRTA1B6W8aeakXQkp8++znKHJq8w+oScDHDC/ekaQ1idgVu9fvPfzuZbL/2MQ==</D></RSAKeyValue>"
		'
		Set r1= loader.loadClass("Crypt/Rsa")
		
			r1.Privatekey=privatekey_csharp
		
			a="Hello WTS"
			b=r1.Encrypt(a)
			c=r1.Decrypt(b)
			d=r1.SignData(a,"SHA1")
			e=r1.VerifyData(a,"SHA1",d)

		Set r1=Nothing
        '##
		
		s=""
		s=s+"a = "+a+Chr(10)+Chr(10)
		s=s+"b = Encrypt(a) : "+Chr(10)+b+Chr(10)+Chr(10)
		s=s+"c = Decrypt(b) :"+Chr(10)+c+Chr(10)+Chr(10)
		s=s+"d = SignData(a,""SHA1"") :"+Chr(10)+d+Chr(10)+Chr(10)
		s=s+"d = VerifyData(a,""SHA1"",d) :"+CStr(e)+Chr(10)+Chr(10)
		wts.responses.SetOutput s

	End Sub

End Class
%>