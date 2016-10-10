<!--#include file="pagecontroller/init.asp"-->
<%
ntime = timer
Call objTemplate.displayFile("t:templates\form01.tpl")
echo "<p>" & timer-ntime & "</p>"
%>