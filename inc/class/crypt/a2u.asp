<%
'@title: Class_Crypt_A2U
'@author: ekede.com
'@date: 2017-02-13
'@description: UNICODE字符串转换

'ANSI并不是某一种特定的字符编码，而是在不同的系统中，ANSI表示不同的编码。
'微软用一个叫“Windows code pages”（在命令行下执行chcp命令可以查看当前code page的值）的值来判断系统默认编码
'@CODEPAGE作用于所有静态的字符串, Response.CodePage,Session.CodePage作用于所有动态输出的字符串。

'SetLocale "zh-CN" 设定本地字符集
'其实在 GBK 集中，“轻”字对印的编码是二个字节: C7 E1 即是二进制的：11000111 11100001因为首位是1, 被当成有符号数了
'那怎么才能取得无符号数的值呢，加上65536便成。所以一般呢取得字符的本地编码值Asc("轻")+65536

'我们使用char来定义字符，占用一个字节，最多只能表示128个字符，也就是ASCII码中的字符. char可以表示所有的英文字符，在以英语为母语的国家完全没有问题。
'汉语、日语等有成千上万个字符，需要用多个字节来表示，称之为宽字符(Wide Character)
'Unicode 是宽字符编码的一种，已经被现代计算机指定为默认的编码方式

'Asc("轻") -14361 Asc 是按本地字符集取文字的编码数值
'AscB("轻") 123   作用于包含在字符串中的字节数据,返回第一个字节的字符代码，而非字符的字符代码
'AscW("轻") -28805 函数返回Unicode字符代码，若平台不支持Unicode，则与Asc函数功能相同

'Chr 返回与指定的 ANSI 字符代码相对应的字符。
'ChrB 不是返回一个或两个字节的字符，而总是返回单个字节的字符。chrB(ascB("轻"))
'ChrW 它的参数是一个Unicode(宽字符)的字符代码

Class Class_Crypt_A2U

    '@Encode(ByRef str): 将字符串中字符转UNICODE字符代码

    Public Function Encode(ByRef str) 'AscW()
        Dim a,s
        For i = 1 To Len(str)
            a = AscW(Mid(str, i, 1))
            If a<0 Then a = a + 65536
            s = s&"&#"&a&";"
        Next
        Encode = s
    End Function

    '@Decode(ByRef str): 将字符串中UNICODE字符代码转字符

    Public Function Decode(ByRef str) 'ChrW()
        If InStr(str, "&#")>0 Then
            Dim arr, s
            arr = Split(str, "&#")
            For i = 0 To UBound(arr)
                If arr(i)<>"" Then s = s&ChrW(Left(arr(i), Len(arr(i)) -1))
            Next
            Decode = s
        Else
            Decode = str
        End If
    End Function

End Class
%>