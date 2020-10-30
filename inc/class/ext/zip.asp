<%
'@title: Class_Ext_zip
'@author: ekede.com
'@date: 2018-07-16
'@description: zip压缩,解压缩类

Class Class_Ext_zip

    Private isDebug_
    Dim fSvrZip
	
    '@isDebug: 是否设置为调试模式
	
    Public Property Let isDebug(Values)
        isDebug_ = Values
    End Property
	
    Private Sub Class_Initialize()
		If IsEmpty(DEBUGS) Then
		   isDebug_ = False
		Else
		   isDebug_ = DEBUGS
		End If
	    '
        On Error Resume Next
		Set fSvrZip=Server.Createobject("dyy.zipsvr")
        If Err.Number <> 0 Then OutErr("创建dyy.zipsvrg组件失败")
    End Sub

    Private Sub Class_Terminate()
		Set fSvrZip=nothing
    End Sub

    '@Zip(p, f): Place To File

    Public Function Zip(ByRef p,ByRef f)
	
		Set fzip=fSvrZip.ZipCom
			fzip.fileName = f
			fzip.AddFiles p&"\*.*"
			fzip.password = ""
			fzip.Process
		Set fzip=Nothing
		
    End Function
	
	'@UnZip(ByRef f,ByRef p): File To Place

    Public Function UnZip(ByRef f,ByRef p)
	
		Set funzip=fSvrZip.UnZipCom
			funzip.fileName = f
			funzip.objDir = p
			funzip.force2CreateObjDir = True
			funzip.Process
		Set funzip=Nothing
		
    End Function
	
    '错误提示

    Private Sub OutErr(ByRef str)
		Response.charSet = "utf-8"
		Response.Write str
		Response.End
    End Sub

End Class
%>