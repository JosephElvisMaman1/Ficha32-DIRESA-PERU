' ============================================================
'  Ficha N.° 32 - Telemedicina (MINSA / DITEL)
'  Lanzador de escritorio silencioso (sin ventana de consola)
'  Doble clic para iniciar la aplicación gráfica al instante.
' ============================================================
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetParentFolderName(WScript.ScriptFullName)

' Invocar pythonw para inicio limpio en segundo plano sin ventana negra
strCmd = "cmd /c cd /d """ & strPath & """ && (where pythonw >nul 2>nul && start pythonw app_ficha32.py || python app_ficha32.py)"
WshShell.Run strCmd, 0, False

Set WshShell = Nothing
Set fso = Nothing
