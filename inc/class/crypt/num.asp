<%
'@title: Class_Crypt_Num
'@author: ekede.com
'@date: 2017-02-13
'@description: 进制转换 binary二进制,decimal十进制,octal八进制,hexadecimal十六

Class Class_Crypt_Num

    '@CBit(num): 十进制转二进制

    Public Function CBit(num)
        cBitstr = ""
        If Len(num)>0 And IsNumeric(num) Then
            Do While Not num \ 2 < 1
                cBitstr = (num Mod 2) &cBitstr
                num = num \ 2
            Loop
        End If
        CBit = num&cBitstr
    End Function

    '@CDec(num): 二进制转十进制

    Public Function CDec(num)
        cDecstr = 0
        If Len(num)>0 And IsNumeric(num) Then
            For inum = 0 To Len(num) -1
                cDecstr = cDecstr + 2^inum * CInt(Mid(num, Len(num) - inum, 1))
            Next
        End If
        CDec = cDecstr
    End Function

    '@BcH(num): 二进制转十六进制

    Public Function BcH(num)
        BcH = Hex(cDec(num))
    End Function

    '@HcB(num): 十六进制转二进制

    Public Function HcB(num) '字符串
        If Len(num)>0 Then
            HcBstr = ""
            For i = 1 To Len(num)
                Select Case (Mid(num, i, 1))
                    Case "0" HcBstr = HcBstr&"0000"
                    Case "1" HcBstr = HcBstr&"0001"
                    Case "2" HcBstr = HcBstr&"0010"
                    Case "3" HcBstr = HcBstr&"0011"
                    Case "4" HcBstr = HcBstr&"0100"
                    Case "5" HcBstr = HcBstr&"0101"
                    Case "6" HcBstr = HcBstr&"0110"
                    Case "7" HcBstr = HcBstr&"0111"
                    Case "8" HcBstr = HcBstr&"1000"
                    Case "9" HcBstr = HcBstr&"1001"
                    Case "A" HcBstr = HcBstr&"1010"
                    Case "B" HcBstr = HcBstr&"1011"
                    Case "C" HcBstr = HcBstr&"1100"
                    Case "D" HcBstr = HcBstr&"1101"
                    Case "E" HcBstr = HcBstr&"1110"
                    Case "F" HcBstr = HcBstr&"1111"
                End Select
            Next
        End If
        HcB = HcBstr
    End Function

    '@OcB(num): 八进制转二进制

    Public Function OcB(num)
        OcBstr = ""
        If Len(num)>0 And IsNumeric(num) Then
            For i = 1 To Len(num)
                Select Case (Mid(num, i, 1))
                    Case "0" OcBstr = OcBstr&"000"
                    Case "1" OcBstr = OcBstr&"001"
                    Case "2" OcBstr = OcBstr&"010"
                    Case "3" OcBstr = OcBstr&"011"
                    Case "4" OcBstr = OcBstr&"100"
                    Case "5" OcBstr = OcBstr&"101"
                    Case "6" OcBstr = OcBstr&"110"
                    Case "7" OcBstr = OcBstr&"111"
                End Select
            Next
        End If
        OcB = OcBstr
    End Function

    '@BcO(num): 二进制转八进制

    Public Function BcO(num)
        BcO = Oct(cDec(num))
    End Function

    '@DcH(num): 十进制转十六进制

    Public Function DcH(num)
        DcH = Hex(num) 'system
    End Function

    '@HcD(num): 十六进制转十进制

    Public Function HcD(num) '字符串或者数字
        HcD = cDec(HcB(num))
    End Function

    '@DcO(num): 十进制转八进制

    Public Function DcO(num)
        DcO = Oct(num) 'system
    End Function

    '@OcD(num): 八进制转十进制

    Public Function OcD(num)
        OcD = cDec(OcB(num))
    End Function

    '@HcO(num): 十六进制转八进制

    Public Function HcO(num)
        HcO = Oct(HcD(num))
    End Function

    '@OcH(num): 八进制转十六进制

    Public Function OcH(num)
        OcH = Hex(OcD(num))
    End Function

End Class
%>