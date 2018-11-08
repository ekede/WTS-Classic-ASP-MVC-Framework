<%
'@title: Class_Crypt_Base64
'@author: ekede.com
'@date: 2017-02-13
'@description: Base64

Class Class_Crypt_Base64
	
	'@Base642Bytes(str): Base642Bytes
	
	Public Function Base642Bytes(str)
        Dim objXML, objXMLNode
        Set objXML = server.CreateObject("msxml2.domdocument")
        Set objXMLNode = objXML.createelement("b64")
			objXMLNode.datatype = "bin.base64"
			objXMLNode.text = str
			Base642Bytes = objXMLNode.nodetypedvalue
        Set objXMLNode = Nothing
        Set objXML = Nothing
	End Function
	
	'@Bytes2Base64(bytes): Bytes2Base64
	
	Public Function Bytes2Base64(bytes)
        Dim objXML, objXMLNode
        Set objXML = server.CreateObject("msxml2.domdocument")
        Set objXMLNode = objXML.createelement("b64")
			objXMLNode.datatype = "bin.base64"
			objXMLNode.nodetypedvalue = bytes
			Bytes2Base64 = objXMLNode.text
        Set objXMLNode = Nothing
        Set objXML = Nothing
	End Function
	
End Class
%>