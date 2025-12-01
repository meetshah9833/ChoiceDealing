<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Site.Master" CodeBehind="DSP.aspx.cs" Inherits="ChoiceDealing.DSP" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <div class="row">
        <div class="col-sm-12 col-md-12 col-lg-12 col-xs-12">
            <div class="panel panel-default">
                <div class="panel-heading">
                    <h4 class="panel-title col-lg-offset-5 col-md-offset-5 col-sm-offset-5 col-xs-offset-5">
                        <span style="margin-left: 10px"></span><span class="HeadingName">DSP Basket File</span>
                    </h4>
                    <ul class="list-inline panel-actions">
                        <li><a href="#" id="panel-fullscreen" role="button" title="Toggle fullscreen"><i class="glyphicon glyphicon-resize-full" aria-hidden="true"></i></a></li>
                    </ul>
                </div>
                <div class="panel-body" style="padding: 5px; overflow-x: auto">
                    <div class="container-fluid">
                        <table class="table-bordered"
                            style="max-width: 100%; background-color: aliceblue;">
                            <tr>
                                <td class="text-left" colspan="4">
                                    <label>DSP File Upload:-</label>
                                    <asp:FileUpload ID="fileUpload" runat="server" CssClass="form-control"
                                        Style="display: inline-block; width: auto; margin-left: 10px;" />
                                </td>
                                <%--<td class="text-centre" colspan="4">
                                    <label>Bhav File:-</label>
                                     <asp:FileUpload ID="BhavfileUpload" runat="server" CssClass="form-control" 
                                     Style="display: inline-block; width: auto; margin-left: 10px;" />
                                </td>--%>
                            </tr>
                            <tr>
                                <td class="text-left" colspan="4">
                                    <asp:Button ID="btnUpload" runat="server" Text="Upload" CssClass="btn btn-primary"
                                        OnClick="btnUpload_Click" />
                                    <asp:Label ID="lblMessage" runat="server" CssClass="text-info"></asp:Label>
                                </td>
                                <%--<td class="text-center" colspan="4">
                                    <asp:Button ID="btnBhavUpload" runat="server" Text="Bhav Upload" CssClass="btn btn-primary" 
                                 OnClick="btnBhavUpload_Click" />
                                    <asp:Label ID="lblBhavMessage" runat="server" CssClass="text-info"></asp:Label>
                                </td>--%>
                                <td class="text-right" colspan="4">
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
                <div class="adjusttop container-fluid BgGridCss" style="width: 100%">
                    <asp:GridView ID="DSPReport" runat="server" AutoGenerateColumns="false"
                        OnRowDataBound="DSPReport_RowDataBound"
                        OnRowCommand="DSPReport_RowCommand" CssClass="LabelTextCSS" Width="100%" Height="100%" Font-Size="15px">
                        <HeaderStyle CssClass="HeaderCSS" Font-Size="15px" />
                        <Columns>
                            <asp:TemplateField HeaderText="SCHEMECODE" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblSCHEMECODE" runat="server" ToolTip="SCHEMECODE" Text='<% #Bind("SCHEMECODE") %>'></asp:Label>
                                </ItemTemplate>

                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Buy/Sell" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:DropDownList ID="ddlBuySell" runat="server" AutoPostBack="true" OnSelectedIndexChanged="RowValueChanged">
                                        <asp:ListItem Text="Buy" Value="B"></asp:ListItem>
                                        <asp:ListItem Text="Sell" Value="S"></asp:ListItem>
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="CLIENTCODE" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtClientcode" runat="server" CssClass="noofBaskets"
                                        Text='<%# Eval("CLIENTCODE") %>' AutoPostBack="true" OnTextChanged="RowValueChanged"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="NoofBaskets" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtNoofBaskets" runat="server" CssClass="noofBaskets"
                                        Text='<%# Eval("NoofBaskets") %>' AutoPostBack="true" OnTextChanged="RowValueChanged"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="TradeInstruction" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtTradeInstruction" runat="server" CssClass="noofBaskets"
                                        Text='<%# Eval("TradeInstructions") %>' AutoPostBack="true" OnTextChanged="RowValueChanged"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Initial QUANTITY" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblInitialQUANTITY" runat="server" ToolTip="InitialQUANTITY" Text='<% #Bind("Transactions") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="QUANTITY" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblQUANTITY" runat="server" ToolTip="QUANTITY" Text='<% #Bind("QUANTITY") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="LINECOUNT" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblLINECOUNT" runat="server" ToolTip="LINECOUNT" Text='<% #Bind("LINECOUNT") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>

                            <asp:ButtonField HeaderText="View." CommandName="View" ButtonType="Link" Text="View"
                                ItemStyle-ForeColor="blue" ItemStyle-Font-Underline="true" ItemStyle-Font-Bold="true"
                                ItemStyle-Font-Size="11px" ItemStyle-Font-Names="Verdana" />

                            <asp:ButtonField HeaderText="Download." CommandName="Download" ButtonType="Link" Text="Download"
                                ItemStyle-ForeColor="blue" ItemStyle-Font-Underline="true" ItemStyle-Font-Bold="true"
                                ItemStyle-Font-Size="11px" ItemStyle-Font-Names="Verdana" />
                        </Columns>
                        <AlternatingRowStyle CssClass="AlternativeRowCss" />
                    </asp:GridView>
                </div>
            </div>
        </div>
    </div>
</asp:Content>
