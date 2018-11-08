<%
'@title: Class_Session
'@author: ekede.com
'@date: 2017-12-28
'@description: Session操作类

Class Class_Session

    Private sid_
	
    '@session_id: 取得Session id

    Public Property Get session_id
        session_id = sid_
    End Property

    Private Sub Class_Initialize()
        sid_ = Session.SessionID
        Session.CodePage = 65001
        Session.Timeout = 30
    End Sub

    Private Sub Class_Terminate()
    End Sub

    '@SetS(k, v): 写

    Public Sub SetS(k, v)
        Session.Contents(k) = v
    End Sub

    '@GetS(k): 读

    Public Function GetS(k)
        GetS = Session(k)
    End Function

    '@GetAllS(k): 读集合

    Public Function GetAllS(k)
        GetS = Session.Contents
    End Function

    '@DelS(k):  删

    Public Sub DelS(k)
        Session.Contents.Remove(k)
    End Sub

    '@DelAllS(k): RemoveAll

    Public Sub DelAllS(k)
        Session.Contents.RemoveAll()
    End Sub

    '@CleanS(k): Abandon

    Public Sub CleanS(k)
        Session.Abandon()
    End Sub

End Class
%>