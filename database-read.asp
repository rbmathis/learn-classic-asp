<!--#include file="./models/product.asp" -->
<%
  ' --------------------------------------------------------
  '                     VBSCRIPT PART 
  ' --------------------------------------------------------
  ' declare the variables
  Dim connection, recordset, sql, connectionString
  Dim countries

  ' to setup connectionString follow this tutorial https://stackoverflow.com/a/5678835/1843755
  connectionString = Application("connectionString")

  ' create an instance of ADO connection and recordset objects
  Set connection = Server.CreateObject("ADODB.Connection")
  Set products = Server.CreateObject("Scripting.Dictionary")
  
  ' open connection in the database
  connection.ConnectionString = connectionString
  connection.Open()
  
  Dim myProduct, seq
  Set recordset = connection.Execute("select * from Products")
  seq = 0
  Do While Not recordset.EOF
    seq = seq+1
    set myProduct = New Product
    myProduct.SKU = recordset.Fields("SKUNumber")
    myProduct.Title = recordset.Fields("Title")
    myProduct.Description = recordset.Fields("Description")
    myProduct.Price = recordset.Fields("Price")
    products.add seq, myProduct
    recordset.MoveNext
  Loop 
  connection.Close()

%>

<%
  ' --------------------------------------------------------
  '                     HTML PART 
  ' --------------------------------------------------------
%>
<!--#include file="layouts/header.asp"-->
  <h1 class="uk-title">Database Read </h1>
  <h3 class="uk-title">List of Products </h3>
  <table class="uk-table uk-table-divider">
    <thead>
      <tr>
        <th>SKU</th>
        <th>Title</th>
        <th>Description</th>
        <th>Price</th>
      </tr>
    </thead>
    <tbody>
      <% For Each item in products %> 
      <tr>
        <td><%= products(item).SKU %></td>
        <td><%= products(item).Title %></td>
        <td><%= products(item).Description %></td>
        <td><%= products(item).Price %></td>
      </tr>
      <% Next %>
    </tbody>
  </table>
<!--#include file="layouts/footer.asp"-->
