<%
'@title: Class_Ext_Array
'@author: POPASP
'@date: 2017-02-13
'@description: 数组操作类

Class Class_Ext_Array

	Private objectTypeArr_

	Private Sub Class_Initialize
		objectTypeArr_ = array( "Dictionary" , "Recordset" )
	End Sub
	
	'@Push( ByRef arr,ByVal item ): 向数组尾部添加一个元素

	Public Sub Push( ByRef arr,ByVal item )
		dim index
		if not isArray(arr) then
			if isEmpty(arr) then
				arr = array()
			else 
				arr = array(arr)
			end if
		end if
		
		index = ubound(arr)+1
		
		redim preserve arr( index )
		
		if Me.Exists(objectTypeArr_,typename( item )) then
			set arr(index) = item
		else
			arr(index) = item
		end if
	End Sub
	
	'@Unshift( ByRef arr,ByVal item ): 向数组头部添加一个元素

	Public Sub Unshift( ByRef arr,ByVal item )
		dim index,i
		if not isArray(arr) then
			if isEmpty(arr) then
				arr = array()
			else 
				arr = array(arr)
			end if
		end if
		
		index = ubound(arr)+1		
		redim preserve arr( index )
		
		for i = index to 1 step -1
			if isObject( arr(i-1) ) then
				set arr(i) = arr(i-1)
			else
				arr(i) = arr(i-1)
			end if
		next
		
		if Me.Exists(objectTypeArr_,typename( item )) then
			set arr(0) = item
		else
			arr(0) = item
		end if
	End Sub
	
	'@Insert( ByRef arr,ByVal pos,ByVal item ): 向数组某个下标处插入元素，其它元素后移
	'从0处插入元素，相当于使用了Unshift
	'插入位置大于数组长度，则从尾部插入，相当于Push
	'如果插入位置是负数，那么从尾部开始算起，-1为最后一个元素，-2为倒数第2个元素
	'如果负数的绝对值大于数组长度，则相当于使用了Unshift

	Public  Sub Insert( ByRef arr,ByVal pos,ByVal item )
		dim index,i,bound
		if not isArray(arr) then
			if isEmpty(arr) then
				arr = array()
			else 
				arr = array(arr)
			end if
		end if
		
		bound = ubound(arr)	
		
		if pos < 0 Then
			pos = bound + pos + 1
		end if
		
		if pos < 0 Then
			pos = 0		
		End If
		
		if pos = 0 then
			Unshift arr,item
			Exit sub
		end if
		
		index = bound+1
		
		if pos >= index then
			Push arr,item
			Exit Sub
		end if
		
		redim preserve arr( index )
		
		for i = index to pos step -1
			if isObject( arr(i-1) ) then
				set arr(i) = arr(i-1)
			else
				arr(i) = arr(i-1)
			end if
		next
		
		if isObject( item ) then
			set arr(pos) = item
		else
			arr(pos) = item
		end if
	End Sub	
	
	
	'@InsertArr( ByRef arr,ByVal pos,ByRef items ): 向数组某个下标处插入元素，其它元素后移
	'从0处插入元素，相当于使用了Unshift
	'插入位置大于数组长度，则从尾部插入，相当于Merge
	'如果插入位置是负数，那么从尾部开始算起，-1为最后一个元素，-2为倒数第2个元素
	'如果负数的绝对值大于数组长度，还是相当于Merge

	Public Sub InsertArr( ByRef arr,ByVal pos,ByRef items )
		dim index,left,right
		if not isArray(arr) then
			if isEmpty(arr) then
				arr = array()
			else 
				arr = array(arr)
			end if
		end if
		
		if not isArray(items) then
			items = array( items )
		end if
		
		bound = ubound(arr)	
		
		if pos < 0 Then
			pos = bound + pos + 1
		end if
		
		if pos < 0 Then
			pos = 0		
		End If
		
		if pos = 0 then
			arr = Merge( items,arr )
			Exit sub
		end if
		
		index = bound+1
		
		if pos >= index then
			arr = Merge( arr,items )
			Exit Sub
		end if
		
		left = slice( arr,0,pos)
		right = slice(arr,pos,-1)		
		arr = Merge( left, items)
		arr = Merge( arr,right)
	End Sub	
	
	'@Pop( ByRef arr ): 从尾部删除一个元素，并返回该元素
	'如果arr不是数组，返回值为Empty，注意返回的元素可能为对象，如果要判断返回元素的类型，可以使用 POP_MVC.Arr.item ，该属性保存的是最后一次删除的元素

	Public Function Pop( ByRef arr )
		dim index
		if not isArray(arr) then
			item = Empty : exit Function
		end if
		index = ubound(arr)
		if index < 0 then
			item = Empty : exit Function		
		end if
		
		if isObject(arr(index)) then
			set item = arr(index) : set Pop = arr(index)
		else
			item = arr(index) : Pop = arr(index)
		end if
		
		if index = 0 then
			arr = array()
		else 
			redim preserve arr( index - 1 )
		end if
	end Function
	
	'@Shift( ByRef arr ): 从头部删除一个元素，并返回该元素
	'如果arr不是数组，返回值为Empty，注意返回的元素可能为对象，如果要判断返回元素的类型，可以使用 POP_MVC.Arr.item ，该属性保存的是最后一次删除的元素

	Public Function Shift( ByRef arr )
		dim index,i
		if not isArray(arr) then
			item = Empty
			exit Function
		end if
		index = ubound(arr)
		if index < 0 then
			item = Empty : exit Function		
		end if
		
		if isObject(arr(0)) then
			set item = arr(0) : set Shift = arr(0)
		else
			item = arr(0) : Shift = arr(0)
		end if
		
		for i = 1 to index
			if isObject(arr(i)) then
				set arr(i-1) = arr(i)
			else
				arr(i-1) = arr(i)
			end if
		next
		
		if index = 0 then
			arr = array()
		else 
			redim preserve arr( index - 1 )
		end if
	end Function
	
	'@Remove( ByRef arr , pos ): 从数组中按下标位置删除一个元素

	Public Function Remove( ByRef arr , pos )
		dim index,i
		if not isArray(arr) then
			item = Empty
			exit Function
		end if
		
		index = ubound(arr) 
		
		if pos < 0 then pos = index + 1 + pos
		
		if index < 0 OR index < pos OR pos < 0 then
			item = Empty : exit Function
		end if
		
		if isObject(arr(pos)) then
			set item = arr(pos) : set Remove = arr(pos)
		else
			item = arr(pos) : Remove = arr(pos)
		end if
		
		for i = pos+1 to index
			if isObject(arr(i)) then
				set arr(i-1) = arr(i)
			else
				arr(i-1) = arr(i)
			end if
		next
		
		if index <= 0 then
			arr = array()
		else 
			redim preserve arr( index - 1 )
		end if
	end Function
	
	'@Swap( ByRef arr, ByVal i, ByVal j): 交换数组中两个下标的值

	Public Sub Swap( ByRef arr, ByVal i, ByVal j)
		dim temp,bnd
		
		bnd = ubound(arr)
		
		'如果下标非法，直接退出
		if i > bnd OR j > bnd Then
			Exit Sub
		end if
		
		' 下标可以小于0，-1为倒数第一个，依次类推
		if i < 0 then i = bnd + i + 1
		if j < 0 then j = bnd + j + 1
		
		'如果下标非法，直接退出
		if i < 0 OR j < 0 OR i = j Then
			Exit Sub
		end if
		
		if isObject( arr(j) ) then
			set temp = arr(j)
		else
			temp = arr(j)
		end if
		if isObject( arr(i) ) then
			set arr(j) = arr(i)
		else
			arr(j) = arr(i)
		end if
		if isObject( temp ) then
			set arr(i) = temp	
		else
			arr(i) = temp	
		end if		
	End Sub
	
	'iReplace(ByRef arr,ByVal find,ByRef replacement): 在数组中搜索给定的值，如果成功则返回相应的键名，否则返回-1

	Public Function iReplace(ByRef arr,ByVal find,ByRef replacement)
		dim i,cnt
		iReplace = 0
		if not isArray(arr) then exit Function
		cnt = Ubound(arr)
		find = LCase(find)
		for i = 0 to cnt			
			if LCase(arr(i)) = find then
				iReplace = iReplace + 1
				arr(i) = replacement	
			end if
		next
	End Function
	
	'@Exists( ByRef arr,ByRef val ): 判断某个值是否存在于数组中，返回True或者False

	Public Function Exists( ByRef arr,ByRef val )
		Exists = (Search( arr,val ) > -1 )
	End Function
	
	'@iExists( ByRef arr,ByRef val ):判断某个值是否存在于数组中，并且不区分大小写，返回True或者False

	Public Function iExists( ByRef arr,ByRef val )
		dim i,cnt
		iExists = false
		if not isArray(arr) then exit Function
		cnt = Ubound(arr)
		for i = 0 to cnt
			if lcase(arr(i)) = lcase(val) then
				iExists = true
				exit Function			
			end if
		next
	End Function
	
	'@iSearch( ByRef arr,ByRef val ): 在数组中搜索给定的值，且不区分大小写，如果成功则返回相应的键名，否则返回-1

	Public Function iSearch( ByRef arr,ByRef val )
		dim i,cnt
		iSearch = -1
		if not isArray(arr) then exit Function
		cnt = Ubound(arr)
		for i = 0 to cnt
			if lcase(arr(i)) = lcase(val) then
				iSearch = i
				exit Function			
			end if
		next
	End Function
	
	'@Search(ByRef arr,ByRef val): 在数组中搜索给定的值，如果成功则返回相应的键名，否则返回-1

	Public Function Search(ByRef arr,ByRef val)
		dim i,cnt
		Search = -1
		if not isArray(arr) then exit Function
		cnt = Ubound(arr)
		for i = 0 to cnt
			if arr(i) = val then
				Search = i
				exit Function			
			end if
		next
	End Function
	
	'@Slice( ByRef arr, ByVal offset,ByVal length ): 取片段函数
	'从数组中取出一段，offset为偏移值，length为取出长度
	'如果 offset 非负，则序列将从 array 中的此偏移量开始。如果 offset 为负，则序列将从 array 中距离末端这么远的地方开始。 
	'如果给出了 length 并且为正，则序列中将具有这么多的单元。如果给出了 length 并且为负，则序列将终止在距离数组末端这么远的地方。

	Public Function Slice( ByRef arr, ByVal offset,ByVal length )
		dim bound : bound = ubound(arr)
		dim i,f,e,ret
		if offset > bound or length = 0 then
			Slice = array() : exit Function
		end if		
		
		if offset<0 then
			f = bound + 1 + offset
			if f < 0 then f = 0
		else
			f = offset
		end if
		
		if length < 0 then
			e = bound + 1 + length
			if e < 0 or e < f then
				Slice = array() : exit Function
			end if
		else
			e = f + length -1
			if e > bound then e = bound
		end if
		
		for i = f to e step 1
			push ret,arr(i)
		next
		Slice = ret
	end Function
	
	'@Unique( ByRef arr ): 移除数组中重复的值并将剩余的值返回一个数组（原数组不变）

	Public Function Unique( ByRef arr )
		dim ret,item
		ret = array()
		for each item in arr
			if Not Exists(ret,item) then
				Push ret,item
			end if
		next
		Unique = ret
	end Function
	
	'@Merge( ByRef arr1,ByRef arr2 ): 将两个数组合并

	Public Function Merge( ByRef arr1,ByRef arr2 )
		dim arr,i,bound
		if isArray(arr1) then
			bound = ubound(arr1)
			for i = 0 to bound
				Push arr,arr1(i)
			next
		end if
		if isArray(arr2) then
			bound = ubound(arr2)
			for i = 0 to bound
				Push arr,arr2(i)
			next
		end if	
		Merge = arr
	end Function
	
	'@Sorts( ByRef arr,ByVal ord ): 按字符串对比排序 asc/desc
	
	Public Sub Sorts( ByRef arr,ByVal ord )
		dim i,j,cnt,bool
		cnt = ubound(arr)		
		for i = 0 to cnt-1			
			for j = i+1 to cnt
			    If ord = "asc" Then
			        bool = arr(i)>arr(j)
				Else
			        bool = arr(i)<arr(j)
				End If
				if bool then
					call Swap(arr,i,j)
				end if
			next
		next
	End Sub	
	
	'@Reverse( ByRef arr ): 返回一个单元顺序相反的数组

	Public Function Reverse( ByRef arr )
		dim i,bnd,counter,ret
		If Not isArray( arr ) Then
			Reverse = Array()
			Exit Function
		End if
		
		bnd = Ubound(arr)
		
		if bnd < 0 Then	'如果是空数组，返回一个空数组
			Reverse = Array()
			Exit Function		
		end if
		
		if bnd = 0 Then	'如果只有一个元素，原样返回
			Reverse = arr
			Exit Function
		end if
		
		ret = arr
		counter = Int(bnd/2)
		for i = 0 to counter
			Swap ret,i,bnd-i
		next
		Reverse = ret
	End Function
	
	Function Range( ByRef min , ByRef  max)
		dim arr,i
		for i = min to max
			Me.Push arr,i
		next
		Range = arr
	end function

	
	'@Shuffle (ByRef arrInput): 将数组打乱

	Public Sub Shuffle (ByRef arrInput)
		Dim arrIndices, iSize, x
		Dim arrOriginal
		iSize = UBound(arrInput)+1
		arrIndices = RandomNoDuplicates(0, iSize-1, iSize)
		arrOriginal = arrInput
		For x=0 To UBound(arrIndices)
			arrInput(x) = arrOriginal(arrIndices(x))
		Next
	End Sub
	
	'this function will return array with "iElements" elements, each of them is random

	Private Function RandomNoDuplicates ( ByRef iMin,ByRef iMax,ByRef iElements )	
		on error resume next		

		If (iMax-iMin+1)<iElements Then
			Exit Function
		End If
		Dim RndArr, x, curRand,pos,temp
		Dim iCount, arrValues()
		Redim arrValues(iMax-iMin)
		For x=iMin To iMax
			arrValues(x-iMin) = x
		Next
		
		RndArr = array()
		'initialize random numbers generator engine:
		Randomize
		iCount=0
		temp = iMax-iMin + 1

		Do Until iCount>=iElements				
			pos = CLng((Rnd*(temp-1))+1)-1
			curRand = arrValues( pos )
			if not Exists( RndArr,curRand ) Then
				Me.Push RndArr,curRand
				temp = temp - 1
				if temp = 0 then
					Exit Do
				end if
				call Me.remove( arrValues,pos )
				iCount = iCount + 1
			end if
		Loop
		err.clear
		RandomNoDuplicates = RndArr
	End Function
End Class
%>