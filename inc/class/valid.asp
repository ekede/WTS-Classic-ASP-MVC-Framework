<%
'@title: Class_Valid
'@author: ekede.com
'@date: 2017-11-29
'@description: Valid类

Class Class_Valid

    Private errs_
	
	'@errs: 依赖errors对象

    Public Property Let errs(Value)
        Set errs_ = Value
    End Property

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub

    'SaveErr(message): 保存错误

    Private Sub SaveErr(message)
        errs_.AddMsg message
    End Sub

    '@Text(ByVal values, ByVal start, ByVal length, ByVal message): 验证字符串长度,并自动截取

    Public Function Text(ByVal values, ByVal start, ByVal length, ByVal message)
        If IsNull(values) Then values = ""
        If length<>0 And start<>0 Then '上限和下限
            If Len(values)>length Or Len(values)<start Then
                If message<>"" Then SaveErr message
            End If
            Text = Left(Trim(values), length)
        ElseIf length = 0 And start<>0 Then '下限和无上限
            If Len(values)<start Then
                If message<>"" Then SaveErr message
            End If
            Text = Trim(values)
        ElseIf length<>0 And start = 0 Then '上限和无下限
            If Len(values)>length Then
                If message<>"" Then SaveErr message
            End If
            Text = Left(Trim(values), length)
        End If
    End Function

    '@Num(ByVal values, ByVal start, ByVal length, ByVal message): 验证数字大小

    Public Function Num(ByVal values, ByVal start, ByVal length, ByVal message)
        If IsNumeric(values) = False Then
            Num = start
            If message<>"" Then SaveErr message
        Else
            Num = CDbl(values)
            If length<>0 And start<>0 Then '上限和下限
                If Num >length Or num<start Then
                    Num = start
                    If message<>"" Then SaveErr message
                End If
            ElseIf length = 0 And start<>0 Then '下限和无上限
                If Num<start Then
                    Num = start
                    If message<>"" Then SaveErr message
                End If
            ElseIf length<>0 And start = 0 Then '上限和无下限
                If Num>length Then
                    Num = length
                    If message<>"" Then SaveErr message
                End If
            End If
        End If
    End Function

    '@IntNum(ByVal values, ByVal start, ByVal length, ByVal message): 验证整形数字大小

    Public Function IntNum(ByVal values, ByVal start, ByVal length, ByVal message)
        IntNum = Fix(num(values, start, length, message))
    End Function

    '@Bool(ByVal values, ByVal message): 验证布尔值 0,1

    Public Function Bool(ByVal values, ByVal message)
        If IsNumeric(values) = False Then
            Bool = 0
            If message<>"" Then SaveErr message
        ElseIf CInt(values) = 0 Then
            Bool = 0
        Else
            Bool = 1
        End If
    End Function

    '@Email(ByVal values,ByVal message): 验证邮箱

    Public Function Email(ByVal values,ByVal message)
        If IsValidEmail (values) = False Then
            Email = ""
            If message<>"" Then SaveErr message
        Else
            Email = values
        End If
    End Function

    '@VerifyCode(ByVal values, ByVal message): 验证码

    Public Function VerifyCode(ByVal values, ByVal message)
        Dim ver
        ver = Session("verifycode") '全局变量
        '
        If Trim(values) <> CStr(ver) Then
            Session("verifycode") = ""
            If message<>"" Then SaveErr message
        Else
            VerifyCode = ver
        End If
    End Function

    '@Times(ByVal values, ByVal start, ByVal length, ByVal message): 验证时间

    Public Function Times(ByVal values,ByVal start, ByVal length, ByVal message)
        If values<>"" And IsDate(values) = true Then
            times = CDate(values)
        Else
            If message<>"" Then SaveErr message
            times = Now()
        End If
    End Function

    '@Safe(ByVal values): 验证单引号

    Public Function Safe(ByVal values)
        If values<>"" Then Safe = Replace(values, "'", "")
    End Function

    '验证邮箱

    Private Function IsValidEmail(email)
        Dim wname, Name, i, c
        IsValidEmail = true
        wname = Split(email, "@")
        If UBound(wname) <> 1 Then
            IsValidEmail = false
            Exit Function
        End If
        For Each Name in wname
            If Len(Name) <= 0 Then
                IsValidEmail = false
                Exit Function
            End If
            For i = 1 To Len(Name)
                c = LCase(Mid(Name, i, 1))
                If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                    IsValidEmail = false
                    Exit Function
                End If
            Next
            If Left(Name, 1) = "." Or Right(Name, 1) = "." Then
                IsValidEmail = false
                Exit Function
            End If
        Next
        If InStr(wname(1), ".") <= 0 Then
            IsValidEmail = false
            Exit Function
        End If
        i = Len(wname(1)) - InStrRev(wname(1), ".")
        If i <> 2 And i <> 3 Then
            IsValidEmail = false
            Exit Function
        End If
        If InStr(email, "..") > 0 Then
            IsValidEmail = false
        End If
    End Function

End Class
%>