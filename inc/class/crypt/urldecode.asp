<%
'@title: Class_Crypt_UrlDecode
'@author: ekede.com
'@date: 2017-02-13
'@description: UrlDecode

Class Class_Crypt_UrlDecode

	'对String对象编码以便它们能在所有计算机上可读,所有空格、标点、重音符号以及其他非ASCII字符都用%xx 编码代替其中xx等于表示该字符的十六进制数
	'@URLDecode(ByVal urlcode): URLDecode
	
	Function URLDecode(ByVal urlcode)
		Dim start,final,length,char,i,butf8,pass
		Dim leftstr,rightstr,finalstr
		Dim b0,b1,bx,blength,position,u,utf8
		On Error Resume Next
		
		b0 = Array(192,224,240,248,252,254)
		urlcode = Replace(urlcode,"+"," ")
		pass = 0
		utf8 = -1
		
		length = Len(urlcode) : start = InStr(urlcode,"%") : final = InStrRev(urlcode,"%")
		If start = 0 Or length < 3 Then URLDecode = urlcode : Exit Function
		leftstr = Left(urlcode,start - 1) : rightstr = Right(urlcode,length - 2 - final)
		
		For i = start To final
		char = Mid(urlcode,i,1)
		If char = "%" Then
		bx = URLDecode_Hex(Mid(urlcode,i + 1,2))
		If bx > 31 And bx < 128 Then
		i = i + 2
		finalstr = finalstr & ChrW(bx)
		ElseIf bx > 127 Then
		i = i + 2
		If utf8 < 0 Then
		butf8 = 1 : blength = -1 : b1 = bx
		For position = 4 To 0 Step -1
		If b1 >= b0(position) And b1 < b0(position + 1) Then
		blength = position
		Exit For
		End If
		Next
		If blength > -1 Then
		For position = 0 To blength
		b1 = URLDecode_Hex(Mid(urlcode,i + position * 3 + 2,2))
		If b1 < 128 Or b1 > 191 Then butf8 = 0 : Exit For
		Next
		Else
		butf8 = 0
		End If
		If butf8 = 1 And blength = 0 Then butf8 = -2
		If butf8 > -1 And utf8 = -2 Then i = start - 1 : finalstr = "" : pass = 1
		utf8 = butf8
		End If
		If pass = 0 Then
		If utf8 = 1 Then
		b1 = bx : u = 0 : blength = -1
		For position = 4 To 0 Step -1
		If b1 >= b0(position) And b1 < b0(position + 1) Then
		blength = position
		b1 = (b1 xOr b0(position)) * 64 ^ (position + 1)
		Exit For
		End If
		Next
		If blength > -1 Then
		For position = 0 To blength
		bx = URLDecode_Hex(Mid(urlcode,i + 2,2)) : i = i + 3
		If bx < 128 Or bx > 191 Then u = 0 : Exit For
		u = u + (bx And 63) * 64 ^ (blength - position)
		Next
		If u > 0 Then finalstr = finalstr & ChrW(b1 + u)
		End If
		Else
		b1 = bx * &h100 : u = 0
		bx = URLDecode_Hex(Mid(urlcode,i + 2,2))
		If bx > 0 Then
		u = b1 + bx
		i = i + 3
		Else
		If Left(urlcode,1) = "%" Then
		u = b1 + Asc(Mid(urlcode,i + 3,1))
		i = i + 2
		Else
		u = b1 + Asc(Mid(urlcode,i + 1,1))
		i = i + 1
		End If
		End If
		finalstr = finalstr & Chr(u)
		End If
		Else
		pass = 0
		End If
		End If
		Else
		finalstr = finalstr & char
		End If
		Next
		URLDecode = leftstr & finalstr & rightstr
	End Function
	'
	Function URLDecode_Hex(ByVal h)
		On Error Resume Next
		h = "&h" & Trim(h) : URLDecode_Hex = -1
		If Len(h) <> 4 Then Exit Function
		If isNumeric(h) Then URLDecode_Hex = cInt(h)
	End Function

End Class
%>