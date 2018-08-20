      </div>
    </div>

    <% 
    Dim q
    q = Request.QueryString("showcode")

    If q="true" Then
    %>

    Path: <%=Response.Write(Server.MapPath("/"))%><br>
    Request: <%=Response.Write(Request.ServerVariables("APPL_PHYSICAL_PATH"))%><br>
    Script: <%=Response.Write(Request.ServerVariables("SCRIPT_NAME"))%><br>

    <%
      Dim Path, p2
      p2 = Right(Request.ServerVariables("SCRIPT_NAME"), Len(Request.ServerVariables("SCRIPT_NAME"))-1)
      path = Request.ServerVariables("APPL_PHYSICAL_PATH") & p2
    %>

    Local Path: <%=path%><br>
    
    <%
    Dim fs, f, contents
    Set fs=Server.CreateObject("Scripting.FileSystemObject")
      Set f = fs.OpenTextFile(path)
        contents = Server.HtmlEncode(f.ReadAll())
        contents = Replace(contents, vbCRLF, "<br>")
      f.Close()
      Set f= Nothing
    set fs=Nothing
    %>

    Text: <br><code><%=contents%></code><br>

    <% End If %>
  
    

    
    <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/9.12.0/highlight.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>

    <!-- UIkit JS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/uikit/3.0.0-beta.25/js/uikit.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/uikit/3.0.0-beta.25/js/uikit-icons.min.js"></script>
    <script type="text/javascript">
      hljs.initHighlightingOnLoad();
    </script>
  </body>
</html>
