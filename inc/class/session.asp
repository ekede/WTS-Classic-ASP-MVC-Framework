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

    '@SetS(ByRef k,ByRef v): 写

    Public Sub SetS(ByRef k,ByRef v)
        Session.Contents(k) = v
    End Sub

    '@GetS(ByRef k): 读

    Public Function GetS(ByRef k)
        GetS = Session(k)
    End Function

    '@GetAllS(ByRef k): 读集合

    Public Function GetAllS(ByRef k)
        GetS = Session.Contents
    End Function

    '@DelS(ByRef k):  删

    Public Sub DelS(ByRef k)
        Session.Contents.Remove(k)
    End Sub

    '@DelAllS(ByRef k): RemoveAll

    Public Sub DelAllS(ByRef k)
        Session.Contents.RemoveAll()
    End Sub

    '@CleanS(ByRef k): Abandon

    Public Sub CleanS(ByRef k)
        Session.Abandon()
    End Sub

End Class
%>