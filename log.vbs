strComputer = “.”
Set objWMIService = GetObject(“winmgmts:\\” & strComputer & “\root\cimv2”)
Set objFSO = CreateObject(“Scripting.FileSystemObject”)
‘Coloque aqui a pasta que será verificada
strPasta = “F:\BKPs Atuais\”
‘Coloque aqui a pasta para a qual os arquivos serão copiados (Se quiser copiar, se nao quiser apague essa linha)
strDest = “F:\BKPS antigos\\BAK-LOG-OLD\”
‘Coloque aqui os tipos de arquivos que serão copiados ou deletados, separados por “;”
arrTipos = “log;bak”
‘ NOME DO ARQUIVO DE LOG
strLogFile = “F:\LOGs\logMover.txt”
‘quantidade de dias
strData = 7
arrTipos = split(arrTipos,”;”)
Set objLogFile = objFSO.OpenTextFile(strLogFile, 8, True, 0)
objLogFile.WriteLine  VBCRLF
objLogFile.WriteLine “===========================================”
objLogFile.WriteLine “ARQUIVOS MOVIDOS EM: ” & now
objLogFile.WriteLine “===========================================”
If (objFSO.FolderExists(strPasta) = True) Then
Set Folder = ObjFSO.GetFolder(strPasta)
Set MyFiles = Folder.files
For Each tipo in arrTipos
For Each MyFiles in Folder.Files
If Right(myfiles.name,3) = tipo And DateDiff(“d”,myfiles.DateLastModified,now) > strData Then
objFSO.Copyfile strPasta & myfiles.name,strDest,True
objLogFile.WriteLine “Arquivo : ”  & myfiles.name &  ” copiado em : ” & Now
objFSO.Deletefile strPasta & myfiles.name
End If
Next
Next
End if
wscript.quit
