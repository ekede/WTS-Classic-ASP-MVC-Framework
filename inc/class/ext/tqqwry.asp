<%
'@title: Class_Ext_Tqqwry
'@author: ekede.com
'@date: 2017-02-13
'@description: 纯真IP库查询

Class Class_Ext_Tqqwry

    ' 变量声名
    Private QQWryFile
    Private Stream, EndIPOff
    Private StartIP, EndIP, CountryFlag
    Dim FirstStartIP, LastStartIP, RecordCount
    Dim Country, LocalStr, Buf, OffSet
	
	'@Country: 国家信息
	
	'@data: IP库文件物理路径
	
    Public Property Let data(str)
        QQWryFile = str
    End Property

    Private Sub Class_Initialize
        QQWryFile = server.MapPath(PATH_ROOT&"data/db/qqwry.dat") 'IP库路径, 物理路径
		'
        Country = ""
        LocalStr = ""
        StartIP = 0
        EndIP = 0
        CountryFlag = 0
        FirstStartIP = 0
        LastStartIP = 0
        EndIPOff = 0
    End Sub

    Private Sub Class_Terminate
        Stream.Close
        Set Stream = Nothing
    End Sub

    ' IP地址转换成整数 ip

    Function IPToInt(IP)
        If InStr(IP, ":")>0 Then IP = "127.0.0.1" '当IP地址是::1这样的地址时返回本机地址
        Dim IPArray, i
        IPArray = Split(IP, ".", -1)
        For i = 0 To 3
            If Not IsNumeric(IPArray(i)) Then IPArray(i) = 0
            If CInt(IPArray(i)) < 0 Then IPArray(i) = Abs(CInt(IPArray(i)))
            If CInt(IPArray(i)) > 255 Then IPArray(i) = 255
        Next
        IPToInt = (CInt(IPArray(0)) * 256 * 256 * 256) + (CInt(IPArray(1)) * 256 * 256) + (CInt(IPArray(2)) * 256) + CInt(IPArray(3))
    End Function

    ' 整数逆转IP地址

    Function IntToIP(IntValue)
        p4 = IntValue - Fix(IntValue / 256) * 256
        IntValue = (IntValue - p4) / 256
        p3 = IntValue - Fix(IntValue / 256) * 256
        IntValue = (IntValue - p3) / 256
        p2 = IntValue - Fix(IntValue / 256) * 256
        IntValue = (IntValue - p2) / 256
        p1 = IntValue
        IntToIP = CStr(p1) & "." & CStr(p2) & "." & CStr(p3) & "." & CStr(p4)
    End Function

    ' 获取开始IP位置

    Private Function GetStartIP(RecNo)
        OffSet = FirstStartIP + RecNo * 7
        Stream.Position = OffSet
        Buf = Stream.Read(7)

        EndIPOff = AscB(MidB(Buf, 5, 1)) + (AscB(MidB(Buf, 6, 1)) * 256) + (AscB(MidB(Buf, 7, 1)) * 256 * 256)
        StartIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1)) * 256) + (AscB(MidB(Buf, 3, 1)) * 256 * 256) + (AscB(MidB(Buf, 4, 1)) * 256 * 256 * 256)
        GetStartIP = StartIP
    End Function

    ' 获取结束IP位置

    Private Function GetEndIP()
        Stream.Position = EndIPOff
        Buf = Stream.Read(5)
        EndIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1)) * 256) + (AscB(MidB(Buf, 3, 1)) * 256 * 256) + (AscB(MidB(Buf, 4, 1)) * 256 * 256 * 256)
        CountryFlag = AscB(MidB(Buf, 5, 1))
        GetEndIP = EndIP
    End Function

    ' 获取地域信息，包含国家和和省市

    Private Sub GetCountry(IP)
        If (CountryFlag = 1 Or CountryFlag = 2) Then
            Country = GetFlagStr(EndIPOff + 4)
            If CountryFlag = 1 Then
                LocalStr = GetFlagStr(Stream.Position)
                ' 以下用来获取数据库版本信息
                If IP >= IPToInt("255.255.255.0") And IP <= IPToInt("255.255.255.255") Then
                    LocalStr = GetFlagStr(EndIPOff + 21)
                    Country = GetFlagStr(EndIPOff + 12)
                End If
            Else
                LocalStr = GetFlagStr(EndIPOff + 8)
            End If
        Else
            Country = GetFlagStr(EndIPOff + 4)
            LocalStr = GetFlagStr(Stream.Position)
        End If
        ' 过滤数据库中的无用信息
        Country = Trim(Country)
        LocalStr = Trim(LocalStr)
        If InStr(Country, "CZ88.NET") Then Country = "本地/局域网"
        If InStr(LocalStr, "CZ88.NET") Then LocalStr = "本地/局域网"
    End Sub

    ' 获取IP地址标识符

    Private Function GetFlagStr(OffSet)
        Dim Flag
        Flag = 0
        Do While (True)
            Stream.Position = OffSet
            Flag = AscB(Stream.Read(1))
            If(Flag = 1 Or Flag = 2 ) Then
				Buf = Stream.Read(3)
				If (Flag = 2 ) Then
					CountryFlag = 2
					EndIPOff = OffSet - 4
				End If
				OffSet = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1)) * 256) + (AscB(MidB(Buf, 3, 1)) * 256 * 256)
			Else
				Exit Do
			End If
		Loop
	
		If (OffSet < 12 ) Then
			GetFlagStr = ""
		Else
			Stream.Position = OffSet
			GetFlagStr = GetStr()
		End If
	End Function
	
	' 获取字串信息 (www.viming.com) 这里获取代码最关键了 utf-8
	
	Private Function GetStr()
		Dim c
		getstr = ""
		Dim objstream
		Set objstream = server.CreateObject("adodb.stream")
		objstream.Type = 1
		objstream.mode = 3
		objstream.Open
		c = stream.Read(1)
		Do While (ascb(c)<>0 And Not stream.eos)
			objstream.Write c
			c = stream.Read(1)
		Loop
		objstream.position = 0
		objstream.Type = 2
		objstream.charset = "gb2312"
		getstr = objstream.readtext
		objstream.Close
		Set objstream = Nothing
	End Function
	
	'@QQWry(DotIP): 核心函数，执行IP搜索
	
	Public Function QQWry(DotIP)
	    On Error Resume Next
		
		Dim IP, nRet
		Dim RangB, RangE, RecNo
	
		IP = IPToInt (DotIP)
	
		Set Stream = CreateObject("ADodb.Stream")
		Stream.Mode = 3
		Stream.Type = 1
		Stream.Open
		Stream.LoadFromFile QQWryFile
		If Err.Number<>0 Then OutErr(err.description)
		Stream.Position = 0
		Buf = Stream.Read(8)
		FirstStartIP = AscB(MidB(Buf, 1, 1)) + (AscB(MidB(Buf, 2, 1)) * 256) + (AscB(MidB(Buf, 3, 1)) * 256 * 256) + (AscB(MidB(Buf, 4, 1)) * 256 * 256 * 256)
		LastStartIP = AscB(MidB(Buf, 5, 1)) + (AscB(MidB(Buf, 6, 1)) * 256) + (AscB(MidB(Buf, 7, 1)) * 256 * 256) + (AscB(MidB(Buf, 8, 1)) * 256 * 256 * 256)
		RecordCount = Int((LastStartIP - FirstStartIP) / 7)
		' 在数据库中找不到任何IP地址
		If (RecordCount <= 1) Then
			Country = "未知"
			QQWry = 2
			Exit Function
		End If
	
		RangB = 0
		RangE = RecordCount
	
		Do While (RangB < (RangE - 1))
			RecNo = Int((RangB + RangE) / 2)
			Call GetStartIP (RecNo)
			If (IP = StartIP) Then
				RangB = RecNo
				Exit Do
			End If
			If (IP > StartIP) Then
				RangB = RecNo
			Else
				RangE = RecNo
			End If
		Loop
	
		Call GetStartIP(RangB)
		Call GetEndIP()
	
		If (StartIP <= IP) And ( EndIP >= IP) Then
			' 没有找到
			nRet = 0
		Else
			' 正常
			nRet = 3
		End If
		Call GetCountry(IP)
	
		QQWry = nRet
	End Function
	
    '错误提示

    Public Sub OutErr(ErrMsg)
            Response.charset = "utf-8"
            Response.Write ErrMsg
            Response.End
    End Sub

End Class
%>