<%
'@title: Class_Ext_Json
'@author: json.org
'@date: 2009-05-12
'@description: 系统JSON类文件 Version 2.0.2

Class Class_Ext_Json

    Public collection
	
    Public count
	
	'@quotedVars: 是否为变量增加引号
	 
    Public quotedVars 

    Public kind '0 = object, 1 = array

    Private Sub Class_Initialize
        Set collection = Server.CreateObject("Scripting.Dictionary")
        quotedVars = True
        count = 0
    End Sub

    Private Sub Class_Terminate
        Set collection = Nothing
    End Sub

    'counter

    Private Property Get counter
        counter = count
        count = count + 1
    End Property

    '@setKind: 设置对象类型 0 = object, 1 = array

    Public Property Let setKind(ByVal fpKind)
        Select Case LCase(fpKind)
            Case "object"
                kind = 0
            Case "array"
                kind = 1
        End Select
    End Property

    '@Pair: Pair(p)=v

    Public Property Let Pair(p, v)
        If IsNull(p) Then p = counter
        collection(p) = v
    End Property

    Public Property Set Pair(p, v)
        If IsNull(p) Then p = counter
        If TypeName(v) <> "Class_Ext_Json" Then
            Err.Raise &hD, "class: class", "class object: '" & TypeName(v) & "'"
        End If
        Set collection(p) = v
    End Property

    Public Default Property Get Pair(p)
    If IsNull(p) Then p = Count - 1
    If IsObject(collection(p)) Then
        Set Pair = collection(p)
    Else
        Pair = collection(p)
    End If
	End Property
	
	'
	
	Public Sub Clean
		collection.RemoveAll
	End Sub
	
	Public Sub Remove(vProp)
		collection.Remove vProp
	End Sub
	
	' data maluplation
	
	' encoding
	
	Public Function JsEncode(Str)
		Dim i, j, aL1, aL2, c, p
	
		aL1 = Array(&h22, &h5C, &h2F, &h08, &h0C, &h0A, &h0D, &h09)
		aL2 = Array(&h22, &h5C, &h2F, &h62, &h66, &h6E, &h72, &h74)
		For i = 1 To Len(Str)
			p = True
			c = Mid(Str, i, 1)
			For j = 0 To 7
				If c = Chr(aL1(j)) Then
					JsEncode = JsEncode & "\" & Chr(aL2(j))
					p = False
					Exit For
				End If
			Next
	
			If p Then
				Dim a
				a = AscW(c)
				If a > 31 And a < 127 Then
					JsEncode = JsEncode & c
				ElseIf a > -1 Or a < 65535 Then
					JsEncode = JsEncode & "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
				End If
			End If
		Next
	End Function
	
	' converting
	
	Public Function ToJSON(vPair)
		Select Case VarType(vPair)
			Case 1 ' Null
				ToJSON = "null"
			Case 7 ' Date
				' yaz saati problemi var
				' jsValue = "new Date(" & Round((vVal - #01/01/1970 02:00#) * 86400000) & ")"
				ToJSON = """" & CStr(vPair) & """"
			Case 8 ' String
				ToJSON = """" & JsEncode(vPair) & """"
			Case 9 ' Object
				Dim bFI, i
				bFI = True
				If vPair.kind Then ToJSON = ToJSON & "[" Else ToJSON = ToJSON & "{"
				For Each i In vPair.collection
					If bFI Then bFI = False Else ToJSON = ToJSON & ","
	
					If vPair.kind Then
						ToJSON = ToJSON & ToJSON(vPair(i))
					Else
						If quotedVars Then
							ToJSON = ToJSON & """" & i & """:" & ToJSON(vPair(i))
						Else
							ToJSON = ToJSON & i & ":" & ToJSON(vPair(i))
						End If
					End If
				Next
				If vPair.kind Then ToJSON = ToJSON & "]" Else ToJSON = ToJSON & "}"
			Case 11
				If vPair Then ToJSON = "true" Else ToJSON = "false"
			Case 12, 8192, 8204
				Dim sEB
				ToJSON = MultiArray(vPair, 1, "", sEB)
			Case Else
				ToJSON = Replace(vPair, ",", ".")
		End Select
	End Function
	
	Public Function MultiArray(aBD, iBC, sPS, ByRef sPT) ' Array BoDy, Integer BaseCount, String PoSition
		Dim iDU, iDL, i ' Integer DimensionUBound, Integer DimensionLBound
		On Error Resume Next
		iDL = LBound(aBD, iBC)
		iDU = UBound(aBD, iBC)
	
		Dim sPB1, sPB2 ' String PointBuffer1, String PointBuffer2
		If Err = 9 Then
			sPB1 = sPT & sPS
			For i = 1 To Len(sPB1)
				If i <> 1 Then sPB2 = sPB2 & ","
				sPB2 = sPB2 & Mid(sPB1, i, 1)
			Next
			MultiArray = MultiArray & ToJSON(Eval("aBD(" & sPB2 & ")"))
		Else
			sPT = sPT & sPS
			MultiArray = MultiArray & "["
			For i = iDL To iDU
				MultiArray = MultiArray & MultiArray(aBD, iBC + 1, i, sPT)
				If i < iDU Then MultiArray = MultiArray & ","
			Next
			MultiArray = MultiArray & "]"
			sPT = Left(sPT, iBC - 2)
		End If
	End Function
	
	'@ToString: Json String
	
	Public Property Get ToString
		ToString = ToJSON(Me)
	End Property
	
	Public Sub Flush
		If TypeName(Response) <> "Empty" Then
			Response.Write(ToString)
		End If
	End Sub
	
	Public Function Clone
		Set Clone = ColClone(Me)
	End Function
	
	Private Function ColClone(core)
		Dim jsc, i
		Set jsc = New Class_Ext_Json
		jsc.kind = core.kind
		For Each i In core.collection
			If IsObject(core(i)) Then
				Set jsc(i) = ColClone(core(i))
			Else
				jsc(i) = core(i)
			End If
		Next
		Set ColClone = jsc
	End Function

End Class
%>