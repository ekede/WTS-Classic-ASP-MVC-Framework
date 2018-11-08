<%
'@title: Crypt/Hex
'@author: ekede.com
'@date: 2017-02-13
'@description: Hex

Class Class_Crypt_Hex

	'@Hex2Bytes(Str): Hex2Bytes
	
	Function Hex2Bytes(Str)
		Set objXML = Server.CreateObject("Msxml2.DOMDocument")
		Set objXMLNode = objXML.createElement("a")
		objXMLNode.DataType = "bin.hex"
		objXMLNode.Text = Str
		Bytes = objXMLNode.NodeTypedValue
		Set objXML = Nothing
		Set objXMLNode = Nothing
		Hex2Bytes=Bytes
	End Function

	'@Bytes2Hex(Bytes): Bytes2Hex
	
	Function Bytes2Hex(Bytes)
		Set objXML = Server.CreateObject("Msxml2.DOMDocument")
		Set objXMLNode = objXML.createElement("a")
		objXMLNode.DataType = "bin.hex"
		objXMLNode.NodeTypedValue = Bytes
		Outstr = Replace(objXMLNode.Text,Chr(10),"")
		Set objXML = Nothing
		Set objXMLNode = Nothing
		Bytes2Hex = Outstr
	End Function
	
End Class
%>