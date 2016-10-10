<%
'path
DEFAULT_CFG_FILE = "./WEB-INF/trophy-cfg.xml"
DEFAULT_TROPHY_PATH = "./"

'config object
Public Const CONFIG_OBJ_KEY = "_trophy_cfg_obj"
Public Const CONFIG_DATE_KEY = "_trophy_cfg_time"
Public Const CONFIG_SESSION_VALIDATOR_KEY = "_trophy_page_validatorstack"
Public Const CONFIG_SESSION_REQUEST_KEY = "_trophy_page_requeststack"
applicationConfigKey = "_trophy_config_source"
applicationTimeKey = "_trophy_config_time"

'generic
Public Const FILE_SEPARATOR = "\"
Public Const TROPHY_VERSION = "0.4.1"
Public Const CONFIG_FILE_NAME = "trophy-cfg.xml"

'action handling
Public Const DEFAULT_ACTION_PARAMETER = "action"
Public Const DEFAULT_CONTROLLER_SCRIPT = "controller.asp"

'templates
Public Const TROPHY_TEMPLATE_KEYWORDS = "global,request,cookie,session,servervars,loop,include"
Public Const TROPHY_TEMPLATE_BLOCK_PREFIX = "BLOCK_POINTER_"
Public Const VAR_PATTERN = "\[\![^]]*\]"
'VAR_PATTERN = "\[\![a-zA-Z0-9= :'.(){}<>*/\\$&+\-*#,@`" & """" & "_çþýÝÇÞöÖüÜ\n]*\]"

INCLUDE_PATTERN = "\[\!include[ ]*file\=" & """" & "[a-zA-Z_.:0-9;-]*" & """" & "\]"
FUNCTION_PATTERN = "\[\![a-zA-Z0-9_]*\(.*\)\]"
IF_HEADER_PATTERN = "\[\!if([^]]*)\]"
IF_FOOTER_PATTERN = "\[\!\/if\]"
IF_ATTR_PATTERN = "[a-zA-Z0-9_.-]*\=" & """" & ".*" & """"
ARG_PATTERN = "[^(]*\,[^)]*"

ARG_SEPARATORS = array("'","""","`")
%>