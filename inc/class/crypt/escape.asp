<%
'@title: Class_Crypt_Escape
'@author: ekede.com
'@date: 2017-02-13
'@description: Escape

Class Class_Crypt_Escape

    '@Escape(ByVal Str): 编码

	Public Function Escape(ByVal Str)
		dim i,s,c,a
		s = ""
		For i = 1 To Len(Str)
			c = Mid(str,i,1)
			a = ASCW(c)
			If (a >= 48 And a <= 57) Or (a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Then
				s = s & c
			ElseIf InStr("@*_+-./",c) > 0 Then
				s = s & c
			ElseIf a > 0 And a < 16 Then
				s = s & "%0" & Hex(a)
			ElseIf a >= 16 And a < 256 Then
				s = s & "%" & Hex(a)
			Else
				s = s & "%u" & Hex(a)
			End If
		Next
		Escape = s
	End Function
	
    '@UnEscape(ByVal Str): 解码
	
	Public Function UnEscape(ByVal Str)
		dim i,s,c
		s = ""
		For i = 1 To Len(Str)
			c = Mid(Str,i,1)
			If Mid(Str,i,2) = "%u" And i <= Len(Str) - 5 Then
				If IsNumeric("&H" & Mid(Str,i + 2,4)) Then
					s = s & CHRW(CInt("&H" & Mid(Str,i + 2,4)))
					i = i + 5
				Else
					s = s & c
				End If
			ElseIf c = "%" And i <= Len(Str) - 2 Then
				If IsNumeric("&H" & Mid(Str,i + 1,2)) Then
					s = s & CHRW(CInt("&H" & Mid(Str,i + 1,2)))
					i = i + 2
				Else
					s = s & c
				End If
			Else
				s = s & c
			End If
		Next
		UnEscape = s
	End Function
	
End Class
%>