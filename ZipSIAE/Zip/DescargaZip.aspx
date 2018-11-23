<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="DescargaZip.aspx.cs" Inherits="Zip.DescargaZip" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <link href="css/Zip.css" rel="stylesheet" />
    <link href="Content/bootstrap.min.css" rel="stylesheet" media="screen"/> 
     <link href="Content/bootstrap.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
  <asp:ScriptManager ID="smManager" runat="server" AsyncPostBackTimeout="28800"></asp:ScriptManager><%--AsyncPostBackTimeout en segundos--%>
       <asp:UpdatePanel ID="upaDescarga" runat="server">
           <ContentTemplate>
         <div class="input-group">
            <h1>Descarga de Zip</h1>
            <br />
            <asp:TextBox ID="txtNodo" runat="server" CssClass="form-control" Width="250px" placeholder="Ingresar Nodo"></asp:TextBox>     
            <asp:Button id="btnDescargar" runat="server" Text="Descargar" CssClass="btn btn-info" OnClick="btnDescargar_Click"/>
        </div>
               </ContentTemplate>
           </asp:UpdatePanel>
        <asp:UpdateProgress ID="prgLoadingStatus" runat="server" DynamicLayout="true" AssociatedUpdatePanelID="upaDescarga">
        <ProgressTemplate>
            <div id="overlay">
              <%--  <div id="modalprogress">
                    <div id="theprogress">--%>
                        <asp:Image id="imgWaitIcon" runat="server" ImageAlign="Middle" CssClass="imagen" src="img/load.gif"/>
                      <div class="textodiv" >Descargando Zip...</div> 
                 <%--   </div>
                </div>--%>
            </div>
        </ProgressTemplate>
    </asp:UpdateProgress>
    </form>
    <script src="Scripts/jquery-3.0.0.js"></script>
    <script src="Scripts/bootstrap.js"></script>
    <script src="Scripts/js/Zip.js"></script>
</body>
</html>
