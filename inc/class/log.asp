<%
'@title: Class_Log
'@author: ekede.com
'@date: 2017-12-7
'@description: 日志操作类

Class Class_Log
    '
    Private fso_
    Private logPath_
	
	'@fso: fso对象依赖

    Public Property Let fso(Value)
        Set fso_ = Value
    End Property
	
	'@logPath: 日志根路径

    Public Property Let logPath(Value)
        logPath_ = PATH_ROOT&Value
    End Property

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub

    '@GetLog(names): 取日志

    Public Function GetLog(names)
        Dim paths, str
        paths = LogPath_&names
        str = fso_.ReadTxtFile(fso_.GetMapPath(paths))
        If str = "" Then
            GetLog = -1
        Else
            GetLog = str
        End If
    End Function
	
    '@SetLog(names, content): 写日志

    Public Function SetLog(names, content)
        Dim paths
        paths = logPath_&names
        '
        fso_.CreateFolders fso_.GetMapPath(LogPath_)
        SetLog = fso_.WriteTxtFile(fso_.GetMapPath(paths), content, 3)
    End Function
	
    '@DelLog(names): 删日志

    Public Function DelLog(names)
        delLog = fso_.DeleteAFile(fso_.GetMapPath(logPath_&names))
    End Function

End Class
%>