' Jednorazove vytvori zastupce Kontrola Clearance s ikonou KC (ve slozce aplikace a na plose).
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
CreateObject("WScript.Shell").Run "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & dir & "\klirenc.ps1"" -VytvoritZastupce", 0, False
