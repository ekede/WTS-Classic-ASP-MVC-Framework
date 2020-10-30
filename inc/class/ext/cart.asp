<%
'@title: Class_Ext_Cart
'@author: ekede.com
'@date: 2018-02-11
'@description: 购物车类

Class Class_Ext_Cart

    Private cart, cartName

    Private Sub Class_Initialize
    End Sub

    Private Sub Class_Terminate()
    End Sub

    '@cartId: 设置购物车id，便于区别不同站点
	
    Public Property Let cartId(Value)
	    If Value <> "" Then
           cartName = "cart."&Value
		Else
           cartName = "cart"
		End If
		'
        If Not IsObject(Session(cartName)) Then
            Set Session(cartName) = Server.CreateObject("Scripting.Dictionary")
        End If
        Set cart = Session(cartName)
    End Property
	
    '@HasNum(): 购物车产品总数量
	
    Public Function HasNum()
	    Dim n
		n=0
		For Each x in Cart
			n = n + Cart(x)
		Next
		HasNum = n
	end function

    '@Has(): 购物车产品数量

    Public Function Has()
        has = cart.Count
    End Function

    '@GetAll(): 打印购物车信息

    Public Function GetAll()
        Dim Keys, Items, I
        Keys = cart.Keys
        Items = cart.Items
        For I = 0 To cart.Count -1
            response.Write Keys(I)&":"&Items(I)&Chr(10)
        Next
    End Function

    '@GetIds(): 返回购物车ID

    Public Function GetIds()
        On Error Resume Next
        Dim Keys, I, str
        Keys = cart.Keys
        For I = 0 To cart.Count -1
            If I = 0 Then
                str = Keys(I)
            Else
                str = Str&","&Keys(I)
            End If
        Next
        GetIds = str
    End Function

    '@GetById(ByRef productId): 取产品数量

    Public Function GetById(ByRef productId)
        If cart.Exists(product_id) Then
            GetById = cart.Item(productId)
        Else
            GetById = 0
        End If
    End Function

    '@Add(ByRef productId,ByRef productNum): 添加购物车

    Public Function Add(ByRef productId,ByRef productNum)
        If Not cart.Exists(productId) Then
            cart.Add productId, CInt(productNum)
        Else
            edit productId, cart.Item(productId) + CInt(productNum)
        End If
        Set Session(cartName) = cart
    End Function

    '@Edit(ByRef productId, ByRef productNum): 修改购物车

    Public Function Edit(ByRef productId, ByRef productNum)
        If cart.Exists(productId) Then
            cart.Item(productId) = CInt(productNum)
        Else
            Add productId, productNum
        End If
        Set Session(cartName) = cart
    End Function

    '@Remove(ByRef productId): 移除购物车产品

    Public Function Remove(ByRef productId)
        cart.Remove(productId)
        Set Session(cartName) = cart
    End Function

    '@RemoveAll(): 清空购物车

    Public Function RemoveAll()
        cart.RemoveAll()
        Set Session(cartName) = cart
    End Function
	
	'@CurrencyPrice(ByRef prices, ByRef currencys, ByRef decimals): 公式 - 汇率计算
	
	Public Function CurrencyPrice(ByRef prices, ByRef currencys, ByRef decimals)
		If IsNull(prices) Then
			CurrencyPrice = 0
		Else
			CurrencyPrice = Round(prices * currencys, decimals)
		End If
	End Function
	
    '@BuyDiscount(discount, discounts, quatity): 公式 - 折扣表计算 
	'10:9.5,20:9,30:8.5,40:8

    Public Function BuyDiscount(ByRef discount, ByRef discounts, ByRef quatity)
        Dim i, discount_array, unit_array
        BuyDiscount = 10
        If discount>0 And discount<10 Then
            BuyDiscount = discount
        ElseIf discounts&"" <> "" Then
            discount_array = Split(discounts, ",")
            For i = 0 To UBound(discount_array)
                unit_array = Split(discount_array(i), ":")
                If UBound(unit_array) = 1 Then
                    If CDbl(quatity)>= CDbl(unit_array(0)) Then
                        BuyDiscount = CDbl(unit_array(1))
                    End If
                End If
            Next
        End If
    End Function
	
    '@ShipWeight(ByRef country_code, ByRef sum_weight, ByRef table_fee): 公式 - 重量运费表计算 
	'us,gb,es|1:2,2:3,10:100
	
    Public Function ShipWeight(ByRef country_code, ByRef sum_weight, ByRef table_fee)
        Dim i, j, k
        Dim line_array, country_array, fee_array, unit_array
        ShipWeight = -1
        '
        line_array = Split(table_fee, vbCrLf)
        For i = 0 To UBound(line_array)
            '
            country_array = Split(line_array(i), "|")
            If UBound(country_array) = 1 Then
                If InStr(country_array(0), country_code)>0 Then
                    '
                    fee_array = Split(country_array(1), ",")
                    For k = 0 To UBound(fee_array)
                        '
                        unit_array = Split(fee_array(k), ":")
                        If UBound(unit_array) = 1 Then
                            If CDbl(sum_weight)>= CDbl(unit_array(0)) Then
                                ShipWeight = CDbl(unit_array(1))
                            End If
                        End If
                    Next
                End If
            End If
        Next
    End Function

End Class
%>