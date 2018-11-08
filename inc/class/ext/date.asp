<%
'@title: Class_Ext_Date
'@author: ekede.com
'@date: 2017-02-23
'@description: 时间类

Class Class_Ext_Date

    Private dateTime_
	Private zone_

    Private Sub Class_Initialize
	    dateTime_ = Now()
		zone_ = 8
    End Sub
	
    '@zone: 时区

    Public Property Get zone()
	    zone = zone_
    End Property

    Public Property Let zone(Value)
	    If IsNumeric(Value) Then zone = Value
    End Property

    '@times: 时间

    Public Property Get times()
	    times = dateTime_
    End Property

    Public Property Let times(Value)
	    If IsDate(Value) Then dataTime_ = CDate(Value)
    End Property

    '@unixTimes: 时间戳

    Public Property Get unixTimes()
	    unixTimes = ToUnixTime(LocalTime(zone_, 0, dateTime_)) '转换为0时区日期,日期转时间戳
    End Property

    Public Property Let unixTimes(Value)
	    If IsNumeric(Value) Then dateTime_ = LocalTime(0,zone_,FromUnixTime(Value))     '时间戳转0时区日期，0时区日期转当前时区日期
    End Property
	
    '@ToUnixTime(t): 0时区日期t 转 时间戳

    Public Function ToUnixTime(t)
        ToUnixTime = DateDiff("s", "1970-1-1 0:0:0", t)
    End Function

    '@FromUnixTime(t, z): 时间戳t 转 0时区日期

    Public Function FromUnixTime(t)
		FromUnixTime = DateAdd("s", t, "1970-1-1 0:0:0")
    End Function

    '@LocalTime(fz, tz, t): 转换时区时间 fz->tz

    Public Function LocalTime(fz, tz, t)
        LocalTime = DateAdd("h", (tz - fz), t) '时区相减
    End Function

    '@Week(d): 星期

    Public Function Week(d)
        temp = "Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday"
        temp = Split(temp, ",")
        Week = temp(Weekday(d) -1)
    End Function

    '@Zodiac(d): 生肖

    Public Function Zodiac(d)
        If IsDate(d) Then
            birthyear = Year(d)
            ZodiacList = Array("猴", "鸡", "狗", "猪", "鼠", "牛", "虎", "兔", "龙", "蛇", "马", "羊")
            Zodiac = ZodiacList(birthyear Mod 12)
        End If
    End Function

    '@Constellation(d): 星座

    Public Function Constellation(d)
        If IsDate(d) Then
            ConstellationMon = Month(d)
            ConstellationDay = Day(d)
            If Len(ConstellationMon)<2 Then ConstellationMon = "0"&ConstellationMon
            If Len(ConstellationDay)<2 Then ConstellationDay = "0"&ConstellationDay
            MyConstellation = ConstellationMon&ConstellationDay
            If MyConstellation < 0120 Then
                constellation = "魔羯座 Capricorn"
            ElseIf MyConstellation < 0219 Then
                constellation = "水瓶座 Aquarius"
            ElseIf MyConstellation < 0321 Then
                constellation = "双鱼座 Pisces"
            ElseIf MyConstellation < 0420 Then
                constellation = "白羊座 Aries"
            ElseIf MyConstellation < 0521 Then
                constellation = "金牛座 Taurus"
            ElseIf MyConstellation < 0622 Then
                constellation = "双子座 Gemini"
            ElseIf MyConstellation < 0723 Then
                constellation = "巨蟹座 Cancer"
            ElseIf MyConstellation < 0823 Then
                constellation = "狮子座 Leo"
            ElseIf MyConstellation < 0923 Then
                constellation = "处女座 Virgo"
            ElseIf MyConstellation < 1024 Then
                constellation = "天秤座 Libra"
            ElseIf MyConstellation < 1122 Then
                constellation = "天蝎座 Scorpio"
            ElseIf MyConstellation < 1222 Then
                constellation = "射手座 Sagittarius"
            ElseIf MyConstellation > 1221 Then
                constellation = "魔羯座 Capricorn"
            End If
        End If
    End Function

End Class
%>