' ================================================================
'  SF2 Live Jam Server - Windows Einrichtung
'  Doppelklick auf diese Datei zum Starten
'  Getestet auf Windows 10 / 11
' ================================================================

Option Explicit

Dim oShell, oFSO, oExec
Dim scriptDir, serverFile, nodeExe
Dim nodeVersion, npmVersion
Dim answer

Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

' Ordner dieser .vbs-Datei
scriptDir  = oFSO.GetParentFolderName(WScript.ScriptFullName)
serverFile = scriptDir & "\jam_server.js"

' ================================================================
'  HELPER: Befehl ausfuehren und Ausgabe zurueckgeben
' ================================================================
Function RunCmd(cmd)
    Dim oExec, output
    On Error Resume Next
    Set oExec = oShell.Exec("cmd.exe /c " & cmd & " 2>&1")
    output = ""
    Do While Not oExec.StdOut.AtEndOfStream
        output = output & oExec.StdOut.ReadLine() & vbCrLf
    Loop
    RunCmd = Trim(output)
    On Error GoTo 0
End Function

Function RunCmdSilent(cmd)
    On Error Resume Next
    oShell.Run "cmd.exe /c " & cmd, 0, True
    On Error GoTo 0
End Function

' ================================================================
'  HELPER: Einfache Box mit Ja/Nein
' ================================================================
Function JaNein(msg)
    Dim r
    r = MsgBox(msg, vbYesNo + vbQuestion, "SF2 Jam Server Setup")
    JaNein = (r = vbYes)
End Function

Sub Info(msg)
    MsgBox msg, vbInformation, "SF2 Jam Server Setup"
End Sub

Sub Fehler(msg)
    MsgBox msg, vbCritical, "SF2 Jam Server Setup"
End Sub

' ================================================================
'  1. PRUEFE OB jam_server.js VORHANDEN
' ================================================================
If Not oFSO.FileExists(serverFile) Then
    Fehler "jam_server.js wurde nicht gefunden!" & vbCrLf & vbCrLf & _
           "Bitte diese Datei in denselben Ordner legen wie:" & vbCrLf & _
           serverFile
    WScript.Quit 1
End If

' ================================================================
'  2. PRUEFE NODE.JS
' ================================================================
nodeVersion = RunCmd("node --version")

If Left(nodeVersion, 1) <> "v" Then
    ' Node.js nicht gefunden
    answer = JaNein("Node.js ist nicht installiert." & vbCrLf & vbCrLf & _
                    "Jetzt automatisch herunterladen und installieren?" & vbCrLf & _
                    "(Oeffnet nodejs.org im Browser)")
    If answer Then
        oShell.Run "https://nodejs.org/en/download/"
        Info "Nach der Node.js-Installation bitte diese Datei erneut ausfuehren." & vbCrLf & vbCrLf & _
             "Empfehlung: LTS-Version (Long Term Support) waehlen."
    End If
    WScript.Quit 1
End If

' ================================================================
'  3. PRUEFE package.json UND INSTALLIERE ABHAENGIGKEITEN
' ================================================================
Dim packageFile, wsFile
packageFile = scriptDir & "\package.json"
wsFile      = scriptDir & "\node_modules\ws\package.json"

If Not oFSO.FileExists(packageFile) Then
    ' package.json fehlt - erstelle eine minimale
    Dim oFile
    Set oFile = oFSO.CreateTextFile(packageFile, True)
    oFile.Write "{" & vbCrLf & _
                "  ""name"": ""sf2-jam-server""," & vbCrLf & _
                "  ""version"": ""1.0.0""," & vbCrLf & _
                "  ""dependencies"": {" & vbCrLf & _
                "    ""ws"": ""^8.16.0""" & vbCrLf & _
                "  }" & vbCrLf & _
                "}"
    oFile.Close
End If

If Not oFSO.FileExists(wsFile) Then
    MsgBox "Abhaengigkeiten werden installiert (npm install)..." & vbCrLf & _
           "Dies kann 1-2 Minuten dauern. Bitte warten.", _
           vbInformation, "SF2 Jam Server Setup"

    ' npm install im Hintergrund
    oShell.Run "cmd.exe /k ""cd /d " & Chr(34) & scriptDir & Chr(34) & _
               " && npm install && echo. && echo Installation abgeschlossen! && pause""", _
               1, True

    ' Nochmal pruefen
    If Not oFSO.FileExists(wsFile) Then
        Fehler "npm install ist fehlgeschlagen!" & vbCrLf & vbCrLf & _
               "Bitte manuell ausfuehren:" & vbCrLf & _
               "1. CMD oeffnen" & vbCrLf & _
               "2. cd " & scriptDir & vbCrLf & _
               "3. npm install"
        WScript.Quit 1
    End If
End If

' ================================================================
'  4. ERMITTLE LOKALE IP-ADRESSE
' ================================================================
Dim localIP
localIP = ""
Dim ipOutput, ipLines, ipLine
ipOutput = RunCmd("for /f ""tokens=2 delims=:"" %a in ('ipconfig ^| findstr /i ""IPv4""') do @echo %a")

' Erste nicht-leere Zeile nehmen
Dim ipArr
ipArr = Split(ipOutput, vbCrLf)
Dim i
For i = 0 To UBound(ipArr)
    Dim ipCandidate
    ipCandidate = Trim(ipArr(i))
    If Len(ipCandidate) > 6 And Left(ipCandidate, 3) <> "169" Then
        localIP = ipCandidate
        Exit For
    End If
Next

If localIP = "" Then localIP = "unbekannt (ipconfig pruefen)"

' ================================================================
'  5. FRAGE OB TUNNEL GESTARTET WERDEN SOLL
' ================================================================
Dim startTunnel
startTunnel = JaNein("Soll auch ein Internet-Tunnel (localhost.run) gestartet werden?" & vbCrLf & vbCrLf & _
                     "Damit koennen Mitspieler von UEBERALL verbinden." & vbCrLf & _
                     "(Kein Account noetig - benoetigt SSH)" & vbCrLf & vbCrLf & _
                     "Nein = nur lokales WLAN (schneller)")

' ================================================================
'  6. ZEIGE ZUSAMMENFASSUNG VOR DEM START
' ================================================================
Dim summary
summary = "Alles bereit! Server wird jetzt gestartet." & vbCrLf & vbCrLf & _
          "Node.js:  " & nodeVersion & vbCrLf & _
          "Ordner:   " & scriptDir & vbCrLf & vbCrLf & _
          "============================================" & vbCrLf & _
          "LOKALES WLAN:" & vbCrLf & _
          "  ws://" & localIP & ":8765" & vbCrLf

If startTunnel Then
    summary = summary & vbCrLf & _
              "INTERNET-TUNNEL:" & vbCrLf & _
              "  URL wird im Tunnel-Fenster angezeigt" & vbCrLf & _
              "  (wss://XXXXX.lhr.life)" & vbCrLf
End If

summary = summary & vbCrLf & _
          "============================================" & vbCrLf & _
          "Zum Beenden: Fenster schliessen"

MsgBox summary, vbInformation, "SF2 Jam Server Setup"

' ================================================================
'  7. SERVER STARTEN
' ================================================================
oShell.Run "cmd.exe /k ""title SF2 Jam Server && cd /d " & Chr(34) & scriptDir & Chr(34) & _
           " && echo. && echo  ============================= && " & _
           "echo   SF2 Live Jam Server && " & _
           "echo  ============================= && " & _
           "echo   Lokale IP:  " & localIP & ":8765 && " & _
           "echo. && " & _
           "node jam_server.js""", 1, False

' ================================================================
'  8. TUNNEL STARTEN (falls gewuenscht)
' ================================================================
If startTunnel Then
    ' Kurz warten bis Server oben ist
    WScript.Sleep 2000

    ' SSH verfuegbar?
    Dim sshCheck
    sshCheck = RunCmd("ssh -V")

    If InStr(sshCheck, "OpenSSH") > 0 Or InStr(sshCheck, "SSH") > 0 Then
        oShell.Run "cmd.exe /k ""title SF2 Jam Tunnel (Internet) && " & _
                   "echo. && " & _
                   "echo  Tunnel wird gestartet... && " & _
                   "echo  Warte auf URL... && echo. && " & _
                   "ssh -R 80:localhost:8765 nokey@localhost.run""", 1, False
    Else
        Info "SSH wurde nicht gefunden auf diesem System." & vbCrLf & vbCrLf & _
             "Tunnel manuell starten:" & vbCrLf & _
             "1. OpenSSH installieren (Windows Einstellungen > Apps > Optionale Features)" & vbCrLf & _
             "2. PowerShell oeffnen" & vbCrLf & _
             "3. Eingeben: ssh -R 80:localhost:8765 nokey@localhost.run" & vbCrLf & vbCrLf & _
             "Server laeuft aber bereits lokal auf:" & vbCrLf & _
             "ws://" & localIP & ":8765"
    End If
End If

' ================================================================
'  FERTIG
' ================================================================
' Script beenden - Server laeuft in eigenem Fenster weiter
WScript.Quit 0
