<%
'@title: config
'@author: ekede.com
'@date: 2018-06-17
'@description: 基础,常量,配置

'#配置文件:
'@DEBUGS: 开启调试,会自动关闭loader文件缓存,关闭调试,可提升运行效率
Const DEBUGS = TRUE

'@PATH_INC: 核心程序文件路径,不建议修改
Const PATH_INC = "inc/"
      PATH_CLASS = PATH_INC&"class/"
      PATH_MODULE = PATH_INC&"module/"
Const PATH_MODEL = "model/"
Const PATH_CONTROL = "control/"
Const PATH_VIEW = "view/"
Const PATH_LANGUAGE = "language/"

'@PATH_APP: 定制程序文件路径,文件路径与PATH_INC对应覆盖
'Const PATH_APP = "app/"

'@PATH_DATA: 数据文件路径
Const PATH_DATA = "data/"
      PATH_LOG = PATH_DATA&"log/"
      PATH_PIC = PATH_DATA&"pic/"
      PATH_STATIC = PATH_DATA&"static/"
Const PATH_PIC_THUMBS = "thumb/"
Const PATH_PIC_IMAGES = "image/"

'@DB_TYPE: 数据库配置
Const DB_TYPE = 1 '1:access ; 3:sqlserver ; 5:dsn
      DB_USER = "sa" 'sqlserver
      DB_PASS = "111" 'sqlserver
Const DB_NAME = "hello.mdb" 'access:caca.asp ; sqlserver:caca ; dsn:caca.dsn ;
Const DB_PRE = "wts_"
      DB_PATH = PATH_DATA&"db/" 'access:data/ ; sqlserver:. ; dsn:"/_dsn"
	  
'@MODULES: 模块配置 默认为default
Const MODULES = "help,default"

'##
%>