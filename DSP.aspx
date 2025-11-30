<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Site.Master" CodeBehind="DSP.aspx.cs" Inherits="ChoiceDealing.DSP" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
        <div class="row">
        <div class="col-sm-12 col-md-12 col-lg-12 col-xs-12">
            <div class="panel panel-default"> 
                <div class="panel-heading">
                    <h4 class="panel-title col-lg-offset-5 col-md-offset-5 col-sm-offset-5 col-xs-offset-5">
                        <span style="margin-left: 10px"></span><span class="HeadingName">Moti AMC Basket File</span>
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
    </asp:Content>