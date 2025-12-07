<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Site.Master" CodeBehind="DSPTabName.aspx.cs" Inherits="ChoiceDealing.DSPTabName" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <div class="adjusttop container-fluid BgGridCss" style="width: 100%; height: 100vh; overflow-y: auto">
                    <asp:GridView ID="DSPTabNameReport" runat="server" AutoGenerateColumns="false" CssClass="LabelTextCSS" Width="100%" Height="100%">
                        <HeaderStyle CssClass="HeaderCSS" />
                        <Columns>
                            <asp:TemplateField HeaderText="TABNAME" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblTABNAME" runat="server" ToolTip="TABNAME" Text='<% #Bind("BASKET_ID") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="ISIN" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblISIN" runat="server" ToolTip="ISIN" Text='<% #Bind("ISIN") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="SecurityName" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblSecurityName" runat="server" ToolTip="SecurityName" Text='<% #Bind("SECURITY") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pricedate" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblPricedate" runat="server" ToolTip="Pricedate" Text='<% #Bind("BASKET_DATE") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="ClosingMarketPriceNSE" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblClosingMarketPriceNSE" runat="server" ToolTip="ClosingMarketPriceNSE" Text='<% #Bind("PRICE") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="AdjustedClosingMarketPriceNSE" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblAdjustedClosingMarketPriceNSE" runat="server" ToolTip="AdjustedClosingMarketPriceNSE" Text='<% #Bind("AdjustedClosingMarketPriceNSE") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="PurchaseableUnits" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblPurchaseableUnits" runat="server" ToolTip="PurchaseableUnits" Text='<% #Bind("PurchaseableUnits") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Adjustedvalue" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblAdjustedvalue" runat="server" ToolTip="Adjustedvalue" Text='<% #Bind("Adjustedvalue") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                           <%-- <asp:TemplateField HeaderText="PercentageinCreationUnit" HeaderStyle-CssClass="GridTextCenter" ItemStyle-CssClass="GridTextCenter">
                                <ItemTemplate>
                                    <asp:Label ID="lblPercentageinCreationUnit" runat="server" ToolTip="PercentageinCreationUnit" Text='<% #Bind("PercentageinCreationUnit") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>--%>
                        </Columns>
                        <AlternatingRowStyle CssClass="AlternativeRowCss" />
                    </asp:GridView>
                </div>
    </asp:Content>