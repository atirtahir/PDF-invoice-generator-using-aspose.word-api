<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="AsposeWordApplication._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">




    <h2 style="text-align: center">Invoice Details</h2>
    <p style="text-align: center">
        Input the following fields.
    </p>

   

    <hr />

    <div class="clientsInfo" style="border-style:groove; text-align:center; padding:8px">
        <label>Client's Name</label><br />
        <input type="text" name="cname" runat="server" id="cname" /><br />
        <label>Product Name</label><br />
        <textarea name="pinfo" runat="server" id="pinfo" cols="50" rows="5"></textarea><br />
        <label>Dscription</label><br />
        <textarea name="msg" runat="server" id="msg" cols="50" rows="5"></textarea><br />
        <label>Price</label><br />
        <input type="text" name="price" id="price" runat="server" /><br />
        <br />
        <label>Attach Logo</label><br />
        <input type="file" runat="server" id="file" name="file" style="margin-left:500px" /><br />
        <br />
        <asp:Button ID="Button1" runat="server" Text="Generate Invoice" OnClick="Button1_Click" />
    </div>



</asp:Content>
