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

    Public Property Let fso(Values)
        Set fso_ = Values
    End Property
	
	'@logPath: 日志根路径

    Public Property Let logPath(Values)
        logPath_ = PATH_ROOT&Values
    End Property

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub

    '@GetLog(ByRef names): 取日志

    Public Function GetLog(ByRef names)
        Dim paths, str
        paths = LogPath_&names
        str = fso_.ReadTxtFile(fso_.GetMapPath(paths))
        If str = "" Then
            GetLog = -1
        Else
            GetLog = str
        End If
    End Function
	
    '@SetLog(ByRef names,ByRef content): 写日志

    Public Function SetLog(ByRef names,ByRef content)
        Dim paths
        paths = logPath_&names
        '
        fso_.CreateFolders fso_.GetMapPath(LogPath_)
        SetLog = fso_.WriteTxtFile(fso_.GetMapPath(paths), content, 3)
    End Function
	
    '@DelLog(ByRef names): 删日志

    Public Function DelLog(ByRef names)
        delLog = fso_.DeleteAFile(fso_.GetMapPath(logPath_&names))
    End Function

End Class
%>