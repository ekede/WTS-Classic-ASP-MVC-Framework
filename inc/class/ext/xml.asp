<%
'@title: Class_Ext_Xml
'@author: 做回自己
'@date: 2005-3-15
'@description: XmlDom操作类

Class Class_Ext_Xml

	Private objXml
	Private xmlDoc
	Private xmlPath

	Sub Class_initialize
		Set objXml = Server.CreateObject("MSXML2.DOMDocument")
		objXml.preserveWhiteSpace = true
		objXml.async = false
	End Sub
	
	Sub Class_Terminate
		Set objXml = Nothing
	End Sub

	'@CreateNew(ByRef sName): 建立一个新的XML文档
	
	Public Function CreateNew(ByRef sName)
		Set tmpNode = objXml.createElement(sName)
		objXml.appendChild(tmpNode)
		Set CreateNew = tmpNode
	End Function
	
	'@OpenXml(ByRef sPath): 从外部读入XML文档
	
	Public Function OpenXml(ByRef sPath)
		OpenXml=False
		sPath=Server.MapPath(sPath)
		'Response.Write(sPath)
		xmlPath = sPath
		If  objXml.load(sPath) Then
			Set xmlDoc = objXml.documentElement
			OpenXml=True
	    End If
	End Function
	
	'@LoadXml(ByRef sStr): 从外部读入XML字符串
	
	Public Sub LoadXml(ByRef sStr)
		objXml.loadXML(sStr)
		Set xmlDoc = objXml.documentElement
	End Sub
	
	'@InceptXml(ByRef xObj): 从外部读入XML对象
	
	Public Sub InceptXml(ByRef xObj)
		Set objXml = xObj
		Set xmlDoc = xObj.documentElement
	End Sub

	'@AddNode(ByRef sNode,ByRef rNode): 新增一个节点, sNode STRING 节点名称, rNode OBJECT 增加节点的上级节点引用 

	Public Function AddNode(ByRef sNode,ByRef rNode)
		Dim TmpNode
		Set TmpNode = objXml.createElement(sNode)
		rNode.appendChild TmpNode
		Set AddNode = TmpNode
	End Function
	
	'@AddAttribute(ByRef sName,ByRef sValue,ByRef oNode): sName STRING 属性名称, sValue STRING 属性值, oNode OBJECT 增加属性的对象
	
	Public Function AddAttribute(ByRef sName,ByRef sValue,ByRef oNode)
		oNode.setAttribute sName,sValue
	End Function
	
	'@AddText(ByRef FStr,ByRef cdBool,ByRef oNode): 新增节点内容
	
	Public Function AddText(ByRef FStr,ByRef cdBool,ByRef oNode)
		Dim tmpText
		If cdBool Then
		   Set tmpText = objXml.createCDataSection(FStr)
		Else
		   Set tmpText = objXml.createTextNode(FStr)
		End If
		oNode.appendChild tmpText
	End Function

	'@GetAtt(ByRef aName,ByRef oNode): 取得节点指定属性的值, aName STRING 属性名称, oNode OBJECT 节点引用
	
	Public Function GetAtt(ByRef aName,ByRef oNode)
		dim tmpValue
		tmpValue = oNode.getAttribute(aName)
		GetAtt = tmpValue
	End Function
	
	'@GetNodeName(ByRef oNode): 取得节点名称, oNode OBJECT 节点引用
	
	Public Function GetNodeName(ByRef oNode)
		GetNodeName = oNode.nodeName
	End Function
	
	'@Function GetNodeText(ByRef oNode): 取得节点内容, oNode OBJECT 节点引用
	
	Public Function GetNodeText(ByRef oNode)
	    GetNodeText = oNode.childNodes(0).nodeValue
	End Function
	
	'@GetNodeType(ByRef oNode): 取得节点类型, oNode OBJECT 节点引用
	
	Public Function GetNodeType(ByRef oNode)
		GetNodeType = oNode.nodeType
	End Function
	
	'@FindNodes(ByRef sNode): 查找节点名相同的所有节点
	
	Public Function FindNodes(ByRef sNode)
		Dim tmpNodes
		Set tmpNodes = objXml.getElementsByTagName(sNode)
		Set FindNodes = tmpNodes
	End Function
	
	'@FindNode(ByRef sNode): 查找一个相同节点
	
	Public Function FindNode(ByRef sNode)
		Dim TmpNode
		Set TmpNode=objXml.selectSingleNode(sNode)
		Set FindNode = TmpNode
	End Function
	
	'@DelNode(ByRef sNode): 删除一个节点
	
	Public Function DelNode(ByRef sNode)
		Dim TmpNodes,Nodesss
		Set TmpNodes=objXml.selectSingleNode(sNode)
		Set Nodesss=TmpNodes.parentNode
		Nodesss.removeChild(TmpNodes)
	End Function
	
	'@ReplaceNode(ByRef sNode,ByRef sText,ByRef cdBool): 替换一个节点
	
	Public Function ReplaceNode(ByRef sNode,ByRef sText,ByRef cdBool)
		'replaceChild
		Dim TmpNodes,tmpText
		Set TmpNodes=objXml.selectSingleNode(sNode)
		'AddText sText,cdBool,TmpNodes
		If cdBool Then
		   Set tmpText = objXml.createCDataSection(sText)
		Else
		   Set tmpText = objXml.createTextNode(sText)
		End If
		TmpNodes.replaceChild tmpText,TmpNodes.firstChild
	End Function
	
	'创建XML声明
	
	Private Function ProcessingInstruction
		'//--创建XML声明
		Dim objPi
		Set objPi = objXML.createProcessingInstruction("xml", "version="&chr(34)&"1.0"&chr(34)&" encoding="&chr(34)&"ISO-8859-1"&chr(34))
		'//--把xml生命追加到xml文档
		objXML.insertBefore objPi, objXML.childNodes(0)
	End Function

	'@SaveXML(): 保存XML文档
	
	Public Function SaveXML()
		'ProcessingInstruction()
		objXml.save(xmlPath)
	End Function
	
	'@SaveAsXML(ByRef sPath): 另存XML文档
	
	Public Function SaveAsXML(ByRef sPath)
		ProcessingInstruction()
		objXml.save(sPath)
	End Function

	'相关统计
	
	'@Root: 取得根节点
	
	Property Get Root
	    Set Root = xmlDoc
	End Property
	
	'@Length: 取得根节点下子节点数
	
	Property Get Length
	    Length = xmlDoc.childNodes.length
	End Property

	'@TestNode: 相关测试
	
	Property Get TestNode
	    TestNode = xmlDoc.childNodes(0).text
	End Property

End Class
%>