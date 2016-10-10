<%
Dim a,cmd
shutdownFunctions = objTrophy.Config.shutdown_Functions 
For a=0 to UBound(shutdownFunctions)
	cmd = "Call " & shutdownFunctions(a) & "()"
	Execute(cmd)
Next

if Not objTrophy.isController then
	etime = timer-ntime
	echo "<br>TIMER: " & etime & "<br>"
end if
%>