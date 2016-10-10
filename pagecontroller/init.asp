<%
Dim ntime
ntime = timer
%>
<!--#include file="includes.asp"-->
<%
Dim req,objTrophy,objConf,objTemplate,objError,objFs
Set objFs = newFileSystem()
Set objTrophy = new Trophy
objTrophy.loadInstance
objTrophy.createTemplateInstance
Call loadVariableFiles(objTemplate,objTrophy.config.filestoload)
Call objTrophy.findAction
Call objTrophy.page.runStartup()
Call objTrophy.selectedAction.runValidator()
Call objTrophy.selectedAction.runStartup()
actionFound = true

if Not objTrophy.isController then
	Call req.loadValidatorFields()
end if
%>
