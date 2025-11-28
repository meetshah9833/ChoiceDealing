<%@ Page Title="Moti Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ChoiceDealing._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="row">
        <div class="col-sm-12 col-md-12 col-lg-12 col-xs-12">
            <div class="panel panel-default"> 
                <div class="panel-heading">
                    <h4 class="panel-title col-lg-offset-5 col-md-offset-5 col-sm-offset-5 col-xs-offset-5">
                        <span style="margin-left: 10px"></span><span class="HeadingName">Moti OMS Basket File</span>
                    </h4>
                    <ul class="list-inline panel-actions">
                        <li><a href="#" id="panel-fullscreen" role="button" title="Toggle fullscreen"><i class="glyphicon glyphicon-resize-full" aria-hidden="true"></i></a></li>
                    </ul>
                </div>
                <div class="panel-body" style="padding: 5px; overflow-x: auto">
                    <div class="container-fluid">
                        <table class="col-lg-offset-2 col-md-offset-2 col-sm-offset-2 col-lg-12 col-md-12 col-sm-12 table-bordered"
                            style="max-width: 60%; background-color: aliceblue;">
                            <tr>
                                <td class="col-lg-12 col-md-12 col-sm-12 text-center" colspan="4">
                                    <h4>Upload OMS File</h4>
                                </td>
                                <td class="col-lg-12 col-md-12 col-sm-12 text-center" colspan="4">
                                     <asp:FileUpload ID="fileUpload" runat="server" CssClass="form-control" 
                                     Style="display: inline-block; width: auto; margin-left: 10px;" />
                                </td>
                            </tr>
                            <tr>
                                <td class="col-lg-12 col-md-12 col-sm-12 text-center" colspan="4">
                                    <h4>Upload Moti EQ File</h4>
                                </td>
                                <td class="col-lg-12 col-md-12 col-sm-12 text-center" colspan="4">
                                     <asp:FileUpload ID="lblMotiEQFile" runat="server" CssClass="form-control" 
                                     Style="display: inline-block; width: auto; margin-left: 10px;" />
                                </td>
                            </tr>
                            <tr>
                                <td class="col-lg-12 col-md-12 col-sm-12 text-center" colspan="4">
                                    <asp:Button ID="btnUpload" runat="server" Text="Upload" CssClass="btn btn-primary" 
                                 OnClick="btnUpload_Click" />
                                    <asp:Label ID="lblMessage" runat="server" CssClass="text-info"></asp:Label>
                                </td>
                                <%-- <td class="col-lg-12 col-md-12 col-sm-12 text-center" colspan="4">
                                    <asp:Button ID="btnUpload2" runat="server" Text="Upload" CssClass="btn btn-primary" 
                                 OnClick="btnUploading_Click" />
                                    <asp:Label ID="Label1" runat="server" CssClass="text-info"></asp:Label>
                                </td>--%>
                            
                                <td class="col-lg-12 col-md-12 col-sm-12 text-center" colspan="4">
                                    <asp:Button ID="btnView" runat="server" Text="View" CssClass="btn btn-primary"
                                        OnClick="btnView_Click" ValidationGroup="Date" />
                                    <asp:Button ID="btnDownload" runat="server" Text="Download" CssClass="btn btn-primary"
                                        OnClick="btnDownload_Click" ValidationGroup="Date" />
<%--                                    <asp:ImageButton ID="btnExcel" runat="server" ImageUrl="~/Images/excel.jpg" OnClick="btnView_Click" />--%>
                                </td>
                            </tr>
                        </table>
                    </div>
                </div>
                <div class="adjusttop container-fluid BgGridCss" style="width: 100%; height: 312px; overflow-y: auto">
                    <asp:GridView ID="Moti_ORMSReport" runat="server" AutoGenerateColumns="false" 
                        CssClass="LabelTextCSS" Width="100%" Height="100%">
                        <HeaderStyle CssClass="HeaderCSS" />
                        <Columns>
                            <asp:TemplateField HeaderText="OMSID" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblOMSID" runat="server" ToolTip="OMSID" Text='<% #Bind("OMSID") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Exchange" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblExchange" runat="server" ToolTip="Exchange" Text='<% #Bind("Exchange") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="OrderQuantity" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblOrderQuantity" runat="server" ToolTip="OrderQuantity" Text='<% #Bind("OrderQuantity") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="OrderPrice" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblOrderPrice" runat="server" ToolTip="OrderPrice" Text='<% #Bind("OrderPrice") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="ISINCode" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblISINCode" runat="server" ToolTip="ISINCode" Text='<% #Bind("ISINCode") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                           <asp:TemplateField HeaderText="Sub_Ac_Code_Client_Id" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblSub_Ac_Code_Client_Id" runat="server" ToolTip="Sub_Ac_Code_Client_Id" Text='<% #Bind("Sub_Ac_Code_Client_Id") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                           <%-- <asp:ButtonField HeaderText="View." CommandName="View" ButtonType="Link" Text="View"
                                ItemStyle-ForeColor="blue" ItemStyle-Font-Underline="true" ItemStyle-Font-Bold="true"
                                ItemStyle-Font-Size="11px" ItemStyle-Font-Names="Verdana" />--%>

                           <%-- <asp:ButtonField HeaderText="Download." CommandName="Download" ButtonType="Link" Text="Download"
                                ItemStyle-ForeColor="blue" ItemStyle-Font-Underline="true" ItemStyle-Font-Bold="true"
                                ItemStyle-Font-Size="11px" ItemStyle-Font-Names="Verdana" />--%>
                        </Columns>
                        <AlternatingRowStyle CssClass="AlternativeRowCss" />
                    </asp:GridView>
                </div>
            </div>
        </div>
    </div>

</asp:Content>
