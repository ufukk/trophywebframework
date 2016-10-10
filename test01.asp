<!--#include file="pagecontroller/init.asp"-->
<!--#include file="lib/cls/tools/clsTrophy_TextFormat.asp"-->
<%
Set txtformat = new Trophy_TextFormat_SpecialCode
txt = "akak skd a [B]aka sda[/B] asd a dad[P] adwda daw d[/P] [LINK href=""http://www.ufux.com""]UFUX[/LINK] wa dew ad" & vbNewLine & " asd a dawd awd" & vbNewLine & " as da d "
txt = txtFormat.parse(txt)
%>
<head></head>
<body>
<pre>
<%=txt %>
</pre>
</body>
<!--#include file="pagecontroller/terminate.asp"-->
