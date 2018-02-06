Option Explicit

Dim strDriveLetter

' Dieses Script entfernt unter Windows das Netz-Laufwerk wieder
' Hier alle Angaben machen
strDriveLetter = "P:" ' Zum Beispiel P wie Protonet, inklusive Doppelpunkt

' Ab hier bitte nichts mehr ändern
' ================================

RemoveDrive strDriveLetter

Function RemoveDrive(strDriveLetter)
	Dim objNetwork, objDrives, i
	Set objNetwork = WScript.CreateObject("WScript.Network")
	Set objDrives = objNetwork.EnumNetworkDrives
	For i = 0 to objDrives.Count - 1 Step 2
		if objDrives.Item(i) = strDriveLetter then
			objNetwork.RemoveNetworkDrive strDriveLetter, TRUE, TRUE
		end if
	Next
End Function
