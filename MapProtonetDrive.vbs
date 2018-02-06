Option Explicit

Dim strDriveLetter, strIP, strURL, strUsr, strPas

' Dieses Script bindet einen Protonet-Server unter Windows als Netz-Laufwerk ein.
' Hier alle Angaben machen
strDriveLetter = "P:" ' Zum Beispiel P wie Protonet, inklusive Doppelpunkt
strIP = "192.168.xxx.xxx" ' IP-Adresse der Protonet-Box im LAN
strURL = "box.host.tld" ' Web-Adresse der Protonet-Box (ohne Protokoll https://)
strUsr = "benutzer.name" ' Benutzername (nicht die E-Mail-Adresse!) wie in Einstellungen / Mein Profil angegeben
strPas = "passwort" ' Das Passwort

' Ab hier bitte nichts mehr ändern
' ================================

' Wenn die Protonet-Box über LAN erreichbar ist, dann über LAN verbinden. Andernfalls über WebDav.
if Ping(strIP) = True then
	MapDrive strDriveLetter, "\\" + strIP + "\" + strUsr + "\Groups", FALSE, strUsr, strPas
	 
Else
	if Ping(strURL) = True then
		MapDrive strDriveLetter, "\\" + strURL + "@SSL\DavWWWRoot\dav\Protonet\Groups", FALSE, strUsr, strPas
	end if
end if

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

Function MapDrive(strDriveLetter, strRemoteShare, strPer, strUsr, strPas)
	RemoveDrive strDriveLetter
	Dim objNetwork
	Set objNetwork = WScript.CreateObject("WScript.Network")
	if strUsr = FALSE then
		objNetwork.MapNetworkDrive strDriveLetter, strRemoteShare, strPer
	else
		objNetwork.MapNetworkDrive strDriveLetter, strRemoteShare, strPer, strUsr, strPas
	end if
End Function

Function Ping(strHost)
	dim objPing, objRetStatus

	set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
		("select * from Win32_PingStatus where address = '" & strHost & "'")

	for each objRetStatus in objPing
		if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode<>0 then
			Ping = False
			' Die folgende Zeile auskommentieren, um zu debuggen
			' WScript.Echo "Status code is " & objRetStatus.StatusCode
		else
			Ping = True
			' Die folgenden Zeilen auskommentieren, um zu debuggen
			' Wscript.Echo "Bytes = " & vbTab & objRetStatus.BufferSize
			' Wscript.Echo "Time (ms) = " & vbTab & objRetStatus.ResponseTime
			' Wscript.Echo "TTL (s) = " & vbTab & objRetStatus.ResponseTimeToLive
		end if
	next
End Function 
