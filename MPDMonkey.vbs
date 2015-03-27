Option Explicit     ' report undefined variables, ...

' SDB variable is connected to MediaMonkey application object
Dim pythonpath
pythonpath = "C:\Python34\python.exe"
Dim scriptpath
scriptpath = "C:\Users\Dan\AppData\Roaming\MediaMonkey\Scripts\"

Sub MPDMonkeyRunCommand(command)  	
	Dim wsh, cmd
	Set wsh= CreateObject("WScript.Shell")
	cmd = pythonpath + " " + scriptpath + "MPDMonkey.py" + " " + "-" + command
	wsh.run cmd
End Sub

Sub MPDMonkeySyncPlaylists()  	
	MPDMonkeyRunCommand("syncplaylists")
End Sub

Sub MPDMonkeySyncNowPlaying()  	
	MPDMonkeyRunCommand("syncnowplaying")
End Sub

Sub MPDMonkeyStartMonitor()  	
	MPDMonkeyRunCommand("startmonitor")
End Sub

Sub MPDMonkeyPlay()  	
	MPDMonkeyRunCommand("play")
End Sub

Sub MPDMonkeyStop()  	
	MPDMonkeyRunCommand("stop")
End Sub

Sub MPDMonkeyPause()  	
	MPDMonkeyRunCommand("pause")
End Sub

Sub MPDMonkeyNext()  	
	MPDMonkeyRunCommand("next")
End Sub

Sub MPDMonkeyPrevious()  	
	MPDMonkeyRunCommand("previous")
End Sub
