<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="AsposeWordApplication._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">




    <h2 style="text-align: center">Invoice Details</h2>
    <p style="text-align: center">
        Input the following fields.
    </p>



    <hr />


    <table class="table table-bordered">
        <tr>
            <td>
                <div class="clientsInfo">
                    <h2>Client's Information</h2>
                    <label>Name</label><br />
                    <input type="text" name="cname" runat="server" id="cname"  style="padding:5px" /><br />
                    <br />
                    <label>Email</label><br />
                    <input type="text" name="mail" runat="server" id="mail"  style="padding:5px" /><br />
                    <br />
                    <p class="lead">Item Details</p>
                    <table class="table table-bordered table-striped" id="dataTable">
                       
                        <tr> 
                            <td><input type="text" runat="server" id="item1" placeholder="Item 1" name="item1" style="padding:5px" /></td>
                            <td><input type="text" runat="server" placeholder="Price" id="price1" name="price1"  style="padding:5px" /></td>
                            <td><input type="text" runat="server" placeholder="Quantity" id="qty1" name="qty1"  style="padding:5px" /></td>
                        </tr>
                        <tr> 
                            <td><input type="text" runat="server" id="item2" placeholder="Item 2" name="item2" style="padding:5px" /></td>
                            <td><input type="text" runat="server" placeholder="Price" id="price2" name="price2"  style="padding:5px" /></td>
                            <td><input type="text" runat="server" placeholder="Quantity" id="qty2" name="qty2"  style="padding:5px" /></td>
                        </tr>
                        <tr> 
                            <td><input type="text" runat="server" id="item3" placeholder="Item 3" name="item3" style="padding:5px" /></td>
                            <td><input type="text" runat="server" placeholder="Price" id="price3" name="price3"  style="padding:5px" /></td>
                            <td><input type="text" runat="server" placeholder="Quantity" id="qty3" name="qty3"  style="padding:5px" /></td>
                        </tr>
                    </table>
                   <%-- <label>Product Name</label><br />
                    <input type="text" name="pinfo" runat="server" id="pinfo" /><br />
                    <br />
                    <label>Dscription</label><br />
                    <textarea name="msg" runat="server" id="msg" cols="50" rows="5"></textarea><br />
                    <br />
                    <label>Price</label><br />
                    <input type="text" name="price" id="price" runat="server" /><br /> --%>
                </div>
            </td>
            <td>
                <div ">
                    <h2>Company's Information</h2>
                    <label>Company's Name</label><br />
                    <input type="text" name="comname" runat="server" id="comname"  style="padding:5px" /><br />
                    <br />
                    <label>Email</label><br />
                    <input type="text" name="compmail" runat="server" id="compmail"  style="padding:5px" /><br />
                    <br />
                    <label>Address</label><br />
                    <textarea name="msg" runat="server" id="add" cols="50" rows="5"></textarea><br />
                    <br />
                    <label>Attach Logo</label><br />
                    <input type="file" runat="server" id="file1" name="file" style="margin-left: 5px" /><br />
                    <br />
                    <asp:Button ID="Button2" runat="server" Text="Generate Invoice" OnClick="Button1_Click" />
                </div>
            </td>
        </tr>
    </table>






</asp:Content>
