Set fso=CreateObject("Scripting.FileSystemObject")
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")
Dim objFILE
Dim mover
strHomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
strMove = strHomeFolder & "\AppData\Roaming\Mozilla\Firefox\Profiles\"
set mover = objFSO.GetFolder(strMove)
for each subdir in mover.subfolders
  'normalmente se usa mais de uma string para inserir certificados de segurança no Firefox
  str1=" "
  str2=" "	
  str3=" "
  retorno = subdir 
 outFile = subdir & "\cert_override.txt"	 
  Set objFile = objFSO.CreateTextFile(outFile,True)
  objFile.Write str1 & vbCrLf
  objFile.Write str2 &  vbCrLf
 objFile.Write  str3 & vbCrLf
Next