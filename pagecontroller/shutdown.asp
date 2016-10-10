<%

Public Sub pageShutdown_clearSession()
	Call req.clearSessionData()
End Sub

Public Sub pageShutdown_freeResources()
	Call objTrophy.terminate()
End Sub
%>