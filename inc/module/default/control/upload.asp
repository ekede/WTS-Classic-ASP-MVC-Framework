<%
'@title: Control_Upload
'@author: ekede.com
'@date: 2018-06-09
'@description: 上传演示

Class Control_Upload

    Private Sub Class_Initialize()
    End Sub

    Private Sub Class_Terminate()
    End Sub
	
    '@Index_Action(): 表单

    Sub Index_Action()
        'url
        wts.template.SetVal "form_url", wts.route.ReWrite(wts.site.config("base_url"),"index.asp?route=upload/save")
		'template
        moban = wts.template.Fetch("upload.htm")
        wts.responses.SetOutput moban
    End Sub
	
    '@Save_Action(): 保存
	
    Sub Save_Action()
	    '#上传保存演示:
		Dim upfile,i:i=0
		
		Set upFile=loader.LoadClass("Ext/UpFile")
	   'upFile.IsDebug = True
		upFile.NoAllowExt="asp;exe;htm;html;aspx;cs;vb;js;"
		upFile.GetData(1024*200)
		If upFile.isErr=0 Then  '如果出错
		for each formName in upFile.file '列出所有上传了的文件
		    set oFile=upfile.file(formname)
			if oFile.fileName <> "" then
		       upFile.SaveToFile formname,wts.fso.GetMapPath(PATH_ROOT&PATH_PIC&PATH_PIC_IMAGES&oFile.fileName)
			   i=i+1
			end if
			set oFile=nothing
		Next
		end if	
		set upFile = nothing
		'##
		wts.responses.SetOutput "Save "&i&" Files "
    End Sub

End Class
%>