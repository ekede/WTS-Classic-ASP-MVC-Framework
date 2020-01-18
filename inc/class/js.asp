<script language="javascript" runat="server">
/*
'@title: Js
'@author: ekede.com
'@date: 2020-01-18
'@description: include预加载,不支持loader,Javascript补充Vbscript的不足
*/

Array.prototype.get = function(x)
{ 
	return this[x]; 
}
var js = {
	//'@parseJSON: json字符串转对象
	parseJSON:function(strJSON){ 
	   return eval("(" + strJSON + ")"); 
	},
	//'@decodeUrl: decodeURIComponent
	decodeUrl:function(strUrl){ 
	   return decodeURIComponent(strUrl); 
	}
};
</script>