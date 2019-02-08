<!--#include file="layouts/header.asp"-->

  <h1>Read File</h1>

  <%
    On Error Resume Next
    Dim fso, file, fileSpec, fileName
    Dim countrySplit

    fileName = "/files/newfile.txt"
    fileSpec = Server.MapPath(fileName)
    
    ' OpenTextFile has several mode 
    ' 1 for reading file
    ' 2 for writing file
    ' 8 for append file content
    ' see : https://msdn.microsoft.com/en-us/library/314cz14s.aspx
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filespec) Then

      Set file = fso.OpenTextFile(filespec,1) 

      Response.write "Reading file " & fileSpec & "<br/><br/>"
      Do While Not file.AtEndOfStream
        line = file.ReadLine()
        Response.write line &"<br/>"
        Response.Flush()
      Loop

        file.Close()
        file=Nothing

    Else
      Response.Write "File doesn't exist at : " & filespec

    End If
  %>
<!--#include file="layouts/footer.asp"-->

