<%
'@title: Class_Ext_WebStat
'@author: ekede.com
'@date: 2017-02-13
'@description: 网站统计

Class Class_Ext_WebStat

	Private regEx_
	Private fun_
	
    '@fun: fun对象依赖

    Public Property Let fun(Value)
        Set fun_ = Value
    End Property

    Private Sub Class_Initialize()
		Set regEx_ = New RegExp
			regEx_.IgnoreCase = true
			regEx_.Global = True
    End Sub

    Private Sub Class_Terminate()
		Set regEx_ = Nothing
    End Sub

    '@GetIp(): IP

    Public Function GetIp()
        Dim strIP,strIpAddr
		strIP=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If strIP = "" Or InStr(strIP, "unknown") > 0 Then
            strIpAddr = Request.ServerVariables("REMOTE_ADDR")
        ElseIf InStr(strIP, ",") > 0 Then
            strIpAddr = Mid(strIP, 1, InStr(strIP, ",") -1)
        ElseIf InStr(strIP, ";") > 0 Then
            strIpAddr = Mid(strIP, 1, InStr(strIP, ";") -1)
        Else
            strIpAddr = strIP
        End If
        GetIp = Trim(Mid(strIpAddr, 1, 30))
    End Function

    '@GetSys(): 操作系统

    Public Function GetSys()
        Dim v_soft, msm
        '
        v_soft = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
        If v_soft = "" Or Len(v_soft)<1 Then Exit Function
        '
        If InStr(v_soft, "windows nt 5.0") Then
            msm = "Win 2000"
        ElseIf InStr(v_soft, "windows nt 5.1") Then
            msm = "Win XP"
        ElseIf InStr(v_soft, "windows nt 5.2") Then
            msm = "Win 2003"
        ElseIf InStr(v_soft, "4.0") Then
            msm = "Win NT"
        ElseIf InStr(v_soft, "nt") Then
            msm = "Win NT"
        ElseIf InStr(v_soft, "windows ce") Then
            msm = "Windows CE"
        ElseIf InStr(v_soft, "windows 9") Then
            msm = "Win 9x"
        ElseIf InStr(v_soft, "9x") Then
            msm = "Windows ME"
        ElseIf InStr(v_soft, "98") Then
            msm = "Windows 98"
        ElseIf InStr(v_soft, "windows 95") Then
            msm = "Windows 95"
        ElseIf InStr(v_soft, "win32") Then
            msm = "Win32"
        ElseIf InStr(v_soft, "unix") Or InStr(v_soft, "linux") Or InStr(v_soft, "SunOS") Or InStr(v_soft, "BSD") Then
            msm = "As Unix"
        ElseIf InStr(v_soft, "mac") Then
            msm = "Mac"
        ElseIf InStr(v_soft, "ipod")>0 Or InStr(v_soft, "iphone")>0 Or InStr(v_soft, "ipad")>0 Then
            msm = "IOS"
        ElseIf InStr(v_soft, "android")>0 Then
            msm = "Android"
        ElseIf InStr(v_soft, "series60")>0 Or InStr(v_soft, "series 60")>0 Then
            msm = "Symbian"
        Else
            msm = "Other"
        End If
        GetSys = msm
    End Function

    '@GetBrowser(): 浏览器

    Public Function GetBrowser()
        Dim v_soft, msm
        '
        v_soft = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
        If v_soft = "" Or Len(v_soft)<1 Then Exit Function
        '
        If InStr(v_soft, "msie") Then
            msm = "IE"
        ElseIf InStr(v_soft, "firefox") Then
            msm = "Firefox"
        ElseIf InStr(v_soft, "chrome") Then
            msm = "Chrome"
        ElseIf InStr(v_soft, "safari") Then
            msm = "Safari"
        ElseIf InStr(v_soft, "camino") Then
            msm = "Camino"
        Else
            msm = "Other"
        End If
        GetBrowser = msm

    End Function

    '@GetLanguage(): 语言

    Public Function GetLanguage()
        Dim v_soft, strLanguage
        '
        v_soft = LCase(request.servervariables("HTTP_ACCEPT_LANGUAGE"))
        If v_soft = "" Or Len(v_soft)<1 Then Exit Function
        '
        strLanguage = "af|sq|eu|bg|be|ca|zh-cn|zh-tw|zh-hk|zh|zh-sg|hr|cs|da|nl|nl-be|en-gb|en-us|en-au|en-ca|en-nz|en-ie|en-za|en-jm|en-bz|en-tt|en|et|fo|fa|fi|fr-be|fr-fr|fr-ch|fr-ca|fr-lu|fr|gd|gl|de-at|de-de|de-ch|de-lu|de-li|de|el|hi|hu|is|idorin|ga|it|it-ch|ja|ko|lv|lt|mk|ms|mt|no|pl|pt-br|pt|rm|ro-mo|ro|ru-mo|ru|gd|sr|sk|sl|sb|esores-do|es-ar|es-co|es-mx|es-es|es-gt|es-cr|es-pa|es-ve|es-pe|es-ec|es-cl|es-uy|es-py|es-bo|es-sv|es-hn|es-ni|es-pr|sx|sv|sv-fi|ts|tn|tr|uk|ur|vi|xh|ji|zu"
        strLanguage = Split(strLanguage, "|")
        For i = 0 To UBound(strLanguage)
            If InStr(v_soft, LCase(strLanguage(i)))>0 Then
                GetLanguage = strLanguage(i)
                Exit Function
            End If
        Next
    End Function

    '@GetSearchKeyword(ByRef str): 关键词_

    Public Function GetSearchKeyword(ByRef str)
        Dim v_soft, msm
        '
        v_soft = str
        If v_soft = "" Or Len(v_soft)<1 Then Exit Function
        '
        pattern = "(yandex.*text=([^&]*)|ask.*q=([^&]*)|goo.*MT=([^&]*)|excite.*q=([^&]*)|lycos.*q=([^&]*)|parseek.*w=([^&]*)|t-online.*q=([^&]*)|google.*q=([^&]*)|bing.*q=([^&]*)|yahoo.*p=([^&]*)|aol.*&q=([^&]*)|baidu.*?wd=([^&]*)|baidu.*?word=([^&]*)|sogou.*query=([^&]*)|naver.*query=([^&]*)|yodao.*q=([^&]*))"
        msm = getExp(v_soft,pattern)
        If msm = "" Then
            GetSearchKeyword = msm
        Else
            GetSearchKeyword = fun_.UrlDecodes(msm)
        End If
    End Function

    '@GetPage(ByRef str): 排名_

    Public Function GetPage(ByRef str)
        Dim v_soft
        '
        v_soft = LCase(str)
        If v_soft = "" Or Len(v_soft)<1 Then Exit Function
        '
        Pattern = "(yandex.*p=([^&]*)|ask.*page=([^&]*)|goo.*fr=([^&]*)|excite.*&p=([^&]*)|parseek.*page=([^&]*)|t-online.*start=([^&]*)|google.*start=([^&]*)|bing.*first=([^&]*)|yahoo.*b=([^&]*)|aol.*&page=([^&]*)|baidu.*?pn=([^&]*)|sogou.*page=([^&]*)|naver.*start=([^&]*)|yodao.*page([^&]*))"
        msm = getExp(v_soft,Pattern)
        If msm = "" Then
            getpage = 1
        Else
            getpage = msm
        End If
    End Function

    '@GetEngine(ByRef str): 搜索引擎_

    Public Function GetEngine(ByRef str)
        Dim v_soft, strEngine
        '
        v_soft = LCase(str)
        If v_soft = "" Or Len(v_soft)<1 Then Exit Function
        '
        strEngine = "yandex|ask|goo|excite|parseek|t-online|google|bing|yahoo|aol|baidu|sogou|naver|yodao|lycos"
        strEngine = Split(strEngine, "|")
        For i = 0 To UBound(strEngine)
            If InStr(v_soft, LCase(strEngine(i)&"."))>0 Then
                GetEngine = strEngine(i)
                Exit Function
            End If
        Next
    End Function

    '@GetSpiderBot(): 蜘蛛

    Public Function GetSpiderBot()
        Dim agent, strBot
        '
        agent = Replace(LCase(request.servervariables("http_user_agent")), "+", " ") '蜘蛛名称全部转成小写字母，+加号替换为空格
        If agent = "" Or Len(agent)<1 Then Exit Function

        '百度
        strBot = strBot&"Baiduspider|BaiduCustomer|Baidu-Thumbnail|Baiduspider-Mobile-Gate|Baidu-Transcoder/1.0.6.0|"
        '谷歌google
        strBot = strBot&"Googlebot/2.1|Googlebot-Image/1.0|Feedfetcher-Google|Google Adsense|Google AdWords|Googlebot-Mobile/2.1|GoogleFriendConnect/1.0|"
        '雅虎yahoo
        strBot = strBot&"Yahoo! Slurp|Yahoo! Slurp/3.0|Yahoo! Slurp China|YahooFeedSeeker/2.0|Yahoo Blogs|Yahoo Image|Yahoo AD|"
        '微软bing
        strBot = strBot&"msnbot/1.1|msnbot/2.0b|msrabot/2.0/1.0|msnbot-media/1.0|MSNBot-Products|MSNBot-Academic|MSNBot-NewsBlogs|"
        '腾讯搜搜soso
        strBot = strBot&"Sosospider|Sosoblogspider|Sosoimagespider|"
        '网易有道
        strBot = strBot&"YoudaoBot/1.0|YodaoBot Image/1.0|YodaoBot-Reader/1.0|"
        '搜狐搜狗
        strBot = strBot&"Sogou web robot|Sogou web spider/3.0|Sogou web spider/4.0|Sogou head spider/3.0|Sogou-Test-Spider/4.0|Sogou Orion spider/4.0|"
        'Alexa
        strBot = strBot&"Ia_archiver|Iaarchiver|"
        'Cuil
        strBot = strBot&"Twiceler-0.9|"
        '奇虎
        strBot = strBot&"Qihoo|"
        'ASK.com
        strBot = strBot&"Ask Jeeves/Teoma|50.nu/0.01|ASPSeek|betaBot|BlogPulseLive|BlogPulse (ISSpider-3.0)|BlogVibeBot-v1.1|BlogSearch/2|BuiltWith/0.3|BuzzBot/1.0|Daumoa/2.0|DomainTools|DotBot/1.1|eApolloBot|"
        strBot = strBot&"Exabot|Alltheweb|FeedBurner/1.0|FollowSite Bot|Gaisbot/3.0|Gigabot|GingerCrawler/1.0|hitcrawler_0.1|iaskspider/1.0|iaskspider/2.0|iearthworm|Jyxobot/1|Larbin|lanshanbot|Lycos|MarkMonitor Robots|"
        strBot = strBot&"MJ12bot/v1.2.1|MJ12bot/v1.2.2|MJ12bot/v1.2.3|MJ12bot/v1.2.4|MJ12bot/v1.2.5|NaverBot/1.0|NetcraftSurveyAgent/1.0|Netcraft Web Server Survey|Page2RSS/0.5|PKU Student Spider|psbot/0.1|Altavista|Servage Robot|"
        strBot = strBot&"Snapbot|Spinn3r|Stealer|Tagoobot/3.0|Twingly Recon|urlfan-bot/1.0|WebAlta|Yandex/1.01.001|Yeti/1.0|"
        strBot = strBot&"sqworm|" 'AOL
        strBot = strBot&"panscient.com|" '恶意爬虫
        '
        strBot = Split(strBot, "|")
        For i = 0 To UBound(strBot)
            If InStr(str, LCase(strBot(i)))>0 Then
                GetSpiderBot = strBot(i)
                Exit Function
            End If
        Next
    End Function

    '@GetCountry(ByRef str): 标准国家_

    Public Function GetCountry(ByRef str)
        Dim strProvince, strCountry, i
        If str = "" Then Exit Function
        '根据省份判断中国
        strProvince = "北京市|上海市|天津市|重庆市|香港|澳门|广东省|河北省|山西省|内蒙古|辽宁省|吉林省|黑龙江省|江苏省|浙江省|安徽省|福建省|江西省|山东省|河南省|湖北省|湖南省|广西|海南省|四川省|贵州省|云南省|西藏|陕西省|甘肃省|青海省|宁夏|新疆|台湾省"
        strProvince = Split(strProvince, "|")
        For i = 0 To UBound(strProvince)
            If InStr(str, strProvince(i))>0 Then
                GetCountry = "中国"
                Exit Function
            End If
        Next
        '标准国家名称
        strCountry = "阿富汗|奥兰群岛|阿尔巴尼亚|阿尔及利亚|美属萨摩亚|安道尔|安哥拉|安圭拉|安提瓜和巴布达|阿根廷|亚美尼亚|阿鲁巴|澳大利亚|奥地利|阿塞拜疆|孟加拉|巴林|巴哈马|巴巴多斯|白俄罗斯|比利时|伯利兹|贝宁|百慕大|不丹|玻利维亚|波斯尼亚和黑塞哥维那|博茨瓦纳|布维岛|巴西|文莱|保加利亚|布基纳法索|布隆迪|柬埔寨|喀麦隆|加拿大|佛得角|中非|乍得|智利|圣诞岛|科科斯（基林）群岛|哥伦比亚|科摩罗|刚果（金）|刚果|库克群岛|哥斯达黎加|科特迪瓦|中国|克罗地亚|古巴|捷克|塞浦路斯|丹麦|吉布提|多米尼加|东帝汶|厄瓜多尔|埃及|赤道几内亚|厄立特里亚|爱沙尼亚|埃塞俄比亚|法罗群岛|斐济|芬兰|法国|法国大都会|法属圭亚那|法属波利尼西亚|加蓬|冈比亚|格鲁吉亚|德国|加纳|直布罗陀|希腊|格林纳达|瓜德罗普岛|关岛|危地马拉|根西岛|几内亚比绍|几内亚|圭亚那|香港 （中国）|海地|洪都拉斯|匈牙利|冰岛|印度|印度尼西亚|伊朗|伊拉克|爱尔兰|马恩岛|以色列|意大利|牙买加|日本|泽西岛|约旦|哈萨克斯坦|肯尼亚|基里巴斯|韩国|朝鲜|科威特|吉尔吉斯斯坦|老挝|拉脱维亚|黎巴嫩|莱索托|利比里亚|利比亚|列支敦士登|立陶宛|卢森堡|澳门（中国）|马其顿|马拉维|马来西亚|马达加斯加|马尔代夫|马里|马耳他|马绍尔群岛|马提尼克岛|毛里塔尼亚|毛里求斯|马约特|墨西哥|密克罗尼西亚|摩尔多瓦|摩纳哥|蒙古|黑山|蒙特塞拉特|摩洛哥|莫桑比克|缅甸|纳米比亚|瑙鲁|尼泊尔|荷兰|新喀里多尼亚|新西兰|尼加拉瓜|尼日尔|尼日利亚|纽埃|诺福克岛|挪威|阿曼|巴基斯坦|帕劳|巴勒斯坦|巴拿马|巴布亚新几内亚|巴拉圭|秘鲁|菲律宾|皮特凯恩群岛|波兰|葡萄牙|波多黎各|卡塔尔|留尼汪岛|罗马尼亚|卢旺达|俄罗斯联邦|圣赫勒拿|圣基茨和尼维斯|圣卢西亚|圣文森特和格林纳丁斯|萨尔瓦多|萨摩亚|圣马力诺|圣多美和普林西比|沙特阿拉伯|塞内加尔|塞舌尔|塞拉利昂|新加坡|塞尔维亚|斯洛伐克|斯洛文尼亚|所罗门群岛|索马里|南非|西班牙|斯里兰卡|苏丹|苏里南|斯威士兰|瑞典|瑞士|叙利亚|塔吉克斯坦|坦桑尼亚|台湾 （中国）|泰国|特立尼达和多巴哥|东帝汶|多哥|托克劳|汤加|突尼斯|土耳其|土库曼斯坦|图瓦卢|乌干达|乌克兰|阿拉伯联合酋长国|英国|美国|乌拉圭|乌兹别克斯坦|瓦努阿图|梵蒂冈|委内瑞拉|越南|瓦利斯群岛和富图纳群岛|西撒哈拉|也门|南斯拉夫|赞比亚|津巴布韦"
        strCountry = Split(strCountry, "|")
        For i = 0 To UBound(strCountry)
            If InStr(str, strCountry(i))>0 Then
                GetCountry = strCountry(i)
                Exit Function
            End If
        Next
    End Function
	
	'@GetMobile(): 判断手机浏览器 0-PC, 1-Smart Phone, 2-touch Phone  3-ipad
	
	Public Function GetMobile()
		Dim mobile_browser_type,user_agent,accept
		
		mobile_browser_type = 0
		user_agent=Lcase(Request.ServerVariables("HTTP_USER_AGENT")) 
		accept = Lcase(Request.ServerVariables("HTTP_ACCEPT"))
		
		Select case true
		case InStr(user_agent,"ipad")>0
			mobile_browser_type = 3
		case InStr(user_agent,"ipod")>0 
			mobile_browser_type = 2
		case InStr(user_agent,"iphone")>0 
			mobile_browser_type = 2
		case InStr(user_agent,"android")>0
			mobile_browser_type = 2
		case InStr(user_agent,"opera mini")>0
			mobile_browser_type = 1
		case InStr(user_agent,"blackberry")>0 
			mobile_browser_type = 1
		case InStr(user_agent,"series60")>0 or InStr(user_agent,"series 60")>0  'Symbian OS 
			mobile_browser_type = 1
		case CheckExp(user_agent,"(pre\/|palm os|palm|hiptop|avantgo|plucker|xiino|blazer|elaine)")'Palm OS 
			mobile_browser_type = 1
		case CheckExp(user_agent,"(iris|3g_t|windows ce|opera mobi|iemobile)")'Windows OS 
			mobile_browser_type = 1
		case CheckExp(user_agent,"(maemo|tablet|qt embedded|com2)")'Nokia Tablet 
			mobile_browser_type = 1
		case CheckExp(user_agent,"(mini 9.5|vx1000|lge |m800|e860|u940|ux840|compal|wireless|mobi|ahong|lg380|lgku|lgu900|lg210|lg47|lg920|lg840|lg370|sam-r|mg50|s55|g83|t66|vx400|mk99|d615|d763|el370|sl900|mp500|samu3|samu4|vx10|xda_|samu5|samu6|samu7|samu9|a615|b832|m881|s920|n210|s700|c-810|_h797|mob-x|sk16d|848b|mowser|s580|r800|471x|v120|rim8|c500foma:|160x|x160|480x|x640|t503|w839|i250|sprint|w398samr810|m5252|c7100|mt126|x225|s5330|s820|htil-g1|fly v71|s302|-x113|novarra|k610i|-three|8325rc|8352rc|sanyo|vx54|c888|nx250|n120|mtk |c5588|s710|t880|c5005|i;458x|p404i|s210|c5100|teleca|s940|c500|s590|foma|samsu|vx8|vx9|a1000|_mms|myx|a700|gu1100|bc831|e300|ems100|me701|me702m-three|sd588|s800|8325rc|ac831|mw200|brew |d88|htc\/|htc_touch|355x|m50|km100|d736|p-9521|telco|sl74|ktouch|m4u\/|me702|8325rc|kddi|phone|lg |sonyericsson|samsung|240x|x320|vx10|nokia|sony cmd|motorola|up.browser|up.link|mmp|symbian|smartphone|midp|wap|vodafone|o2|pocket|kindle|mobile|psp|treo|vnd.rim|wml|nitro|nintendo|wii|xbox|archos|openweb|mini|docomo)")
			mobile_browser_type = 1
		case InStr(accept,"text/vnd.wap.wml")>0 or InStr(accept,"application/vnd.wap.xhtml+xml")>0
			mobile_browser_type = 1
		case Request.ServerVariables("HTTP_X_WAP_PROFILE")<>"" or Request.ServerVariables("HTTP_PROFILE")<>""
			mobile_browser_type = 1
		case ubound(filter(array("1207","3gso","4thp","501i","502i","503i","504i","505i","506i","6310","6590","770","802s","a wa","acer","acs","airn","alav","asus","attw","au-m","aur ","aus ","abac","acoo","aiko","alco","alca","amoi","anex","anny","anyw","aptu","arch","argo","bell","bird","bw-n","bw-u","beck","benq","bilb","blac","c55/","cdm-","chtm","capi","cond","craw","dall","dbte","dc-s","dica","ds-d","ds12","dait","devi","dmob","doco","dopo","el49","erk0","esl8","ez40","ez60","ez70","ezos","ezze","elai","emul","eric","ezwa","fake","fly-","fly_","g-mo","g1 u","g560","gf-5","grun","gene","go.w","good","grad","hcit","hd-m","hd-p","hd-t","hei-","hp i","hpip","hs-c","htc ","htc-","htca","htcg","htcp","htcs","htct","htc_","haie","hita","huaw","hutc","i-20","i-go","i-ma","i230","iac","iac-","iac/","ig01","im1k","inno","iris","jata","java","kddi","kgt","kgt/","kpt ","kwc-","klon","lexi","lg g","lg-a","lg-b","lg-c","lg-d","lg-f","lg-g","lg-k","lg-l","lg-m","lg-o","lg-p","lg-s","lg-t","lg-u","lg-w","lg/k","lg/l","lg/u","lg50","lg54","lge-","lge/","lynx","leno","m1-w","m3ga","m50","maui","mc01","mc21","mcca","medi","meri","mio8","mioa","mo01","mo02","mode","modo","mot ","mot-","mt50","mtp1","mtv ","mate","maxo","merc","mits","mobi","motv","mozz","n100","n101","n102","n202","n203","n300","n302","n500","n502","n505","n700","n701","n710","nec-","nem-","newg","neon","netf","noki","nzph","o2 x","o2-x","opwv","owg1","opti","oran","p800","pand","pg-1","pg-2","pg-3","pg-6","pg-8","pg-c","pg13","phil","pn-2","pt-g","palm","pana","pire","pock","pose","psio","qa-a","qc-2","qc-3","qc-5","qc-7","qc07","qc12","qc21","qc32","qc60","qci-","qwap","qtek","r380","r600","raks","rim9","rove","s55/","sage","sams","sc01","sch-","scp-","sdk/","se47","sec-","sec0","sec1","semc","sgh-","shar","sie-","sk-0","sl45","slid","smb3","smt5","sp01","sph-","spv ","spv-","sy01","samm","sany","sava","scoo","send","siem","smar","smit","soft","sony","t-mo","t218","t250","t600","t610","t618","tcl-","tdg-","telm","tim-","ts70","tsm-","tsm3","tsm5","tx-9","tagt","talk","teli","topl","hiba","up.b","upg1","utst","v400","v750","veri","vk-v","vk40","vk50","vk52","vk53","vm40","vx98","virg","vite","voda","vulc","w3c ","w3c-","wapj","wapp","wapu","wapm","wig ","wapi","wapr","wapv","wapy","wapa","waps","wapt","winc","winw","wonu","x700","xda2","xdag","yas-","your","zte-","zeto","acs-","alav","alca","amoi","aste","audi","avan","benq","bird","blac","blaz","brew","brvw","bumb","ccwa","cell","cldc","cmd-","dang","doco","eml2","eric","fetc","hipt","http","ibro","idea","ikom","inno","ipaq","jbro","jemu","java","jigs","kddi","keji","kyoc","kyok","leno","lg-c","lg-d","lg-g","lge-","libw","m-cr","maui","maxo","midp","mits","mmef","mobi","mot-","moto","mwbp","mywa","nec-","newt","nok6","noki","o2im","opwv","palm","pana","pant","pdxg","phil","play","pluc","port","prox","qtek","qwap","rozo","sage","sama","sams","sany","sch-","sec-","send","seri","sgh-","shar","sie-","siem","smal","smar","sony","sph-","symb","t-mo","teli","tim-","tosh","treo","tsm-","upg1","upsi","vk-v","voda","vx52","vx53","vx60","vx61","vx70","vx80","vx81","vx83","vx85","wap-","wapa","wapi","wapp","wapr","webc","whit","winw","wmlb","xda-"),Mid(user_agent,1,4),true))>-1 'Catch all 
			mobile_browser_type = 1
		End Select
		GetMobile = mobile_browser_type		
	End Function


    '@GetLangChs(ByRef strLang): 标准语言

    Private Function GetLangChs(ByRef strLang)
        Select Case strLang
            Case "af"
                outStr = "南非荷兰语"
            Case "sq"
                outStr = "阿尔巴尼亚语"
            Case "ar-ae"
                outStr = "阿拉伯语 - 阿拉伯联合酋长国"
            Case "ar-bh"
                outStr = "阿拉伯语 - 巴林"
            Case "ar-dz"
                outStr = "阿拉伯语 - 阿尔及利亚"
            Case "ar-eg"
                outStr = "阿拉伯语 - 埃及"
            Case "ar-iq"
                outStr = "阿拉伯语 - 伊拉克"
            Case "ar-jo"
                outStr = "阿拉伯语 - 约旦"
            Case "ar-kw"
                outStr = "阿拉伯语 - 科威特"
            Case "ar-lb"
                outStr = "阿拉伯语 - 黎巴嫩"
            Case "ar-ly"
                outStr = "阿拉伯语 - 利比亚"
            Case "ar-ma"
                outStr = "阿拉伯语 - 摩洛哥"
            Case "ar-om"
                outStr = "阿拉伯语 - 阿曼"
            Case "ar-qa"
                outStr = "阿拉伯语 - 卡塔尔"
            Case "ar-sa"
                outStr = "阿拉伯语 - 沙特阿拉伯"
            Case "ar-sy"
                outStr = "阿拉伯语 - 叙利亚"
            Case "ar-tn"
                outStr = "阿拉伯语 - 突尼斯"
            Case "ar-ye"
                outStr = "阿拉伯语 - 也门"
            Case "hy"
                outStr = "亚美尼亚语"
            Case "az-az"
                outStr = "阿泽里语 - 拉丁"
            Case "az-az"
                outStr = "阿泽里语 - 西里尔语"
            Case "eu"
                outStr = "巴斯克语"
            Case "be"
                outStr = "白俄罗斯语"
            Case "bg"
                outStr = "保加利亚语"
            Case "ca"
                outStr = "加泰罗尼亚语"
            Case "zh"
                outStr = "中文"
            Case "zh-cn"
                outStr = "中文 - 中华人民共和国"
            Case "zh-hk"
                outStr = "中文 - 中华人民共和国香港特别行政区"
            Case "zh-mo"
                outStr = "中文 - 中华人民共和国澳门特别行政区"
            Case "zh-sg"
                outStr = "中文 - 新加坡"
            Case "zh-tw"
                outStr = "中文 - 台湾地区"
            Case "hr"
                outStr = "克罗地亚语"
            Case "cs"
                outStr = "捷克语"
            Case "da"
                outStr = "丹麦语"
            Case "nl"
                outStr = "荷兰语"
            Case "nl-nl"
                outStr = "荷兰语"
            Case "nl-be"
                outStr = "荷兰语 - 比利时"
            Case "en"
                outStr = "英语"
            Case "en-au"
                outStr = "英语 - 澳大利亚"
            Case "en-bz"
                outStr = "英语 - 伯利兹"
            Case "en-ca"
                outStr = "英语 - 加拿大"
            Case "en-cb"
                outStr = "英语 - 加勒比地区"
            Case "en-ie"
                outStr = "英语 - 爱尔兰"
            Case "en-jm"
                outStr = "英语 - 牙买加"
            Case "en-nz"
                outStr = "英语 - 新西兰"
            Case "en-ph"
                outStr = "英语 - 菲律宾"
            Case "en-za"
                outStr = "英语 - 南非"
            Case "en-tt"
                outStr = "英语 - 特立尼达岛"
            Case "en-gb"
                outStr = "英语 - 英国"
            Case "en-us"
                outStr = "英语 - 美国"
            Case "et"
                outStr = "爱沙尼亚语"
            Case "fa"
                outStr = "波斯语"
            Case "fi"
                outStr = "芬兰语"
            Case "fo"
                outStr = "法罗语"
            Case "fr"
                outStr = "法语"
            Case "fr-fr"
                outStr = "法语 - 法国"
            Case "fr-be"
                outStr = "法语 - 比利时"
            Case "fr-ca"
                outStr = "法语 - 加拿大"
            Case "fr-lu"
                outStr = "法语 - 卢森堡"
            Case "fr-ch"
                outStr = "法语 - 瑞士"
            Case "gd-ie"
                outStr = "盖尔语 - 爱尔兰"
            Case "gd"
                outStr = "盖尔语 - 苏格兰"
            Case "de"
                outStr = "德语"
            Case "de-de"
                outStr = "德语 - 德国"
            Case "de-at"
                outStr = "德语 - 奥地利"
            Case "de-li"
                outStr = "德语 - 列支敦士登"
            Case "de-lu"
                outStr = "德语 - 卢森堡"
            Case "de-ch"
                outStr = "德语 - 瑞士"
            Case "el"
                outStr = "希腊语"
            Case "he"
                outStr = "希伯来语"
            Case "hi"
                outStr = "印地语"
            Case "hu"
                outStr = "匈牙利语"
            Case "is"
                outStr = "冰岛语"
            Case "id"
                outStr = "印度尼西亚语"
            Case "it"
                outStr = "意大利语"
            Case "it-it"
                outStr = "意大利语 - 意大利"
            Case "it-ch"
                outStr = "意大利语 - 瑞士"
            Case "ja"
                outStr = "日语"
            Case "ko"
                outStr = "朝鲜语"
            Case "lv"
                outStr = "拉脱维亚语"
            Case "lt"
                outStr = "立陶宛语"
            Case "mk"
                outStr = "FYRO 马其顿语"
            Case "ms-my"
                outStr = "马来语 - 马来西亚"
            Case "ms-bn"
                outStr = "马来语 - 文莱"
            Case "mt"
                outStr = "马耳他语"
            Case "mr"
                outStr = "马拉地语"
            Case "no"
                outStr = "挪威语"
            Case "no-no"
                outStr = "挪威语 - 博克马尔"
            Case "no-no"
                outStr = "挪威语 - 尼诺斯克"
            Case "pl"
                outStr = "波兰语"
            Case "pt"
                outStr = "葡萄牙语"
            Case "pt-pt"
                outStr = "葡萄牙语 - 葡萄牙"
            Case "pt-br"
                outStr = "葡萄牙语 - 巴西"
            Case "rm"
                outStr = "拉托-罗马语"
            Case "ro"
                outStr = "罗马尼亚语"
            Case "ro-mo"
                outStr = "罗马尼亚语 - 摩尔多瓦"
            Case "ru"
                outStr = "俄语"
            Case "ru-mo"
                outStr = "俄语 - 摩尔多瓦"
            Case "sa"
                outStr = "梵语"
            Case "sr"
                outStr = "塞尔维亚语"
            Case "sr-sp"
                outStr = "塞尔维亚语 - 西里尔语"
            Case "sr-sp"
                outStr = "塞尔维亚语 - 拉丁"
            Case "tn"
                outStr = "茨瓦纳语"
            Case "sl"
                outStr = "斯洛文尼亚语"
            Case "sk"
                outStr = "斯洛伐克语"
            Case "sb"
                outStr = "索布语"
            Case "es"
                outStr = "西班牙语"
            Case "es-es"
                outStr = "西班牙语 - 西班牙"
            Case "es-ar"
                outStr = "西班牙语 - 阿根廷"
            Case "es-bo"
                outStr = "西班牙语 - 玻利维亚"
            Case "es-cl"
                outStr = "西班牙语 - 智利"
            Case "es-co"
                outStr = "西班牙语 - 哥伦比亚"
            Case "es-cr"
                outStr = "西班牙语 - 哥斯达黎加"
            Case "es-do"
                outStr = "西班牙语 - 多米尼加共和国"
            Case "es-ec"
                outStr = "西班牙语 - 厄瓜多尔"
            Case "es-gt"
                outStr = "西班牙语 - 危地马拉"
            Case "es-hn"
                outStr = "西班牙语 - 洪都拉斯"
            Case "es-mx"
                outStr = "西班牙语 - 墨西哥"
            Case "es-ni"
                outStr = "西班牙语 - 尼加拉瓜"
            Case "es-pa"
                outStr = "西班牙语 - 巴拿马"
            Case "es-pe"
                outStr = "西班牙语 - 秘鲁"
            Case "es-pr"
                outStr = "西班牙语 - 波多黎各"
            Case "es-py"
                outStr = "西班牙语 - 巴拉圭"
            Case "es-sv"
                outStr = "西班牙语 - 萨尔瓦多"
            Case "es-uy"
                outStr = "西班牙语 - 乌拉圭"
            Case "es-ve"
                outStr = "西班牙语 - 委内瑞拉"
            Case "sx"
                outStr = "苏图语"
            Case "sw"
                outStr = "斯瓦希里语"
            Case "sv"
                outStr = "瑞典语"
            Case "sv-se"
                outStr = "瑞典语"
            Case "sv-fi"
                outStr = "瑞典语 - 芬兰"
            Case "ta"
                outStr = "泰米尔语"
            Case "tt"
                outStr = "鞑靼语"
            Case "th"
                outStr = "泰语"
            Case "tr"
                outStr = "土耳其语"
            Case "ts"
                outStr = "汤加语"
            Case "uk"
                outStr = "乌克兰语"
            Case "ur"
                outStr = "乌尔都语 - 巴基斯坦"
            Case "uz-uz"
                outStr = "乌兹别克语 - 西里尔"
            Case "uz-uz"
                outStr = "乌兹别克语 - 拉丁"
            Case "vi"
                outStr = "越南语"
            Case "xh"
                outStr = "科萨语"
            Case "yi"
                outStr = "意第绪语"
            Case "zu"
                outStr = "祖鲁语"
            Case Else
                outStr = ""
        End Select
        GetLangChs = outStr
    End Function
	
	'查找字符串是否存在
	
	Private Function CheckExp(ByRef strng,ByRef patrn)
			regEx_.Pattern = patrn
			CheckExp = regEx_.Test(strng)
	End Function

    '正则返回第一个匹配结果的第一个值

    Private Function GetExp(ByRef strng,ByRef pattern)
        On Error Resume Next
        Dim a, b, j
        regEx_.Pattern = pattern
		Set a = regEx_.Execute(strng)
			If  a.Count>0 Then
				Set b = a(a.Count -1).SubMatches
					For j = 1 To b.Count
						If Len(b(j))>0 Then
							GetExp = b(j)
							Exit Function
						End If
					Next
				Set b = Nothing
			End If
		Set a = Nothing
    End Function

End Class
%>