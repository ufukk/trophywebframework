<!--#include file="lib/cls/core/trophy.const.asp"-->
<!--#include file="lib/cls/core/clsTrophy_Configuration.asp"-->
<!--#include file="lib/cls/core/clsTrophy_Page.asp"-->
<!--#include file="lib/cls/core/clsTrophy_Error.asp"-->
<!--#include file="lib/cls/core/clsTrophy_Action.asp"-->
<!--#include file="lib/cls/core/clsTrophy_RequestHandler.asp"-->
<!--#include file="lib/cls/core/clsTrophy.asp"-->
<!--#include file="lib/cls/db/clsTrophy_Datasource.asp"-->
<!--#include file="lib/cls/tools/clsTrophy_Logger.asp"-->
<!--#include file="lib/commons/commons-trophy.asp"-->
<!--#include file="lib/commons/commons-misc.asp"-->
<!--#include file="pagecontroller/filters.asp"-->
<!--#include file="pagecontroller/shutdown.asp"-->
<%
Dim req
Set objFs = newFileSystem()
Set objTrophy = new Trophy
objTrophy.isController = true
objTrophy.loadInstance
Call objTrophy.findAction()
%>
<!--#include file="pagecontroller/actionrunner.asp"-->
<!--#include file="pagecontroller/terminate.asp"-->