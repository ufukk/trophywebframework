<!--#include file="pagecontroller/init.asp"-->
<!--#include file="lib/cls/tools/clsTrophy_Image.asp"-->
<%
ntime = timer

Set objImage = new Trophy_Image

objImage.path = "nico3.jpg"
echo objImage.html(array("border"),array("1"))

%>