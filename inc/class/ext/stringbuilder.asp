<%
'@title: Class_Ext_StringBuilder
'@author: Steve McMahon
'@date: 2009-05-12
'@description: vbAccelerator

'vbAccelerator的Steve McMahon给我们提供了一个好用的cStringBuilder类，便于我们实现StringBuilder的功能，
'据作者讲添加10,000次类似于”http://vbaccelerator.com/”这样字符串，标准VB方式需要34秒，而使用Steve McMahon的cStringBuilder类只需要0.35秒。效率和速度还是相当不错的。

Class Class_Ext_StringBuilder

	  'the array of strings to concatenate
	  Private arr
	  
	  '@growth: the rate at which the array grows
	  Private growthRate
	  
	  Public Property Let growth(Value) 
		growthRate = Value
	  End Property
	  
	  'the number of items in the array
	  Private itemCount

	  Private Sub Class_Initialize()
		growthRate = 10
		itemCount = 0
		ReDim arr(growthRate)
	  End Sub
	 
	  '@Append(ByVal strValue): Append a new string to the end of the array. 
	  'If the number of items in the array is larger than the actual capacity of the array, then "grow" the array by ReDimming it.
	  
	  Public Sub Append(ByVal strValue)
		strValue=strValue & "" 'code borrowed from FastString to prevent crash on NULL
		If itemCount > UBound(arr) Then
		  ReDim Preserve arr(UBound(arr) + growthRate)
		End If
		arr(itemCount) = strValue
		itemCount = itemCount + 1
	  End Sub
	  
	  '@ToString(): Concatenate the strings
	  'by simply joining your array of strings and adding no separator between elements.
	  
	  Public Function ToString()
		ToString = Join(arr, "")
	  End Function

End Class
%>