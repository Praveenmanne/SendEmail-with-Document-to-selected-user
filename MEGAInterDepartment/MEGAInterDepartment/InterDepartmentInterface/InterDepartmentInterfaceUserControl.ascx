<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls"
    Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages"
    Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="InterDepartmentInterfaceUserControl.ascx.cs"
    Inherits="MEGAInterDepartment.InterDepartmentInterface.InterDepartmentInterfaceUserControl" %>
<style type="text/css">
    .LeftLabelTextFormat
    {
        border-top: 1px solid #D8D8D8;
        color: #525252;
        font-family: verdana;
        font-weight: normal;
        padding-bottom: 6px;
        padding-right: 8px;
        padding-top: 3px;
        text-align: left;
    }
    .ms-formbody
    {
        background: none repeat scroll 0 0 #F6F6F6;
        border-top: 1px solid #D8D8D8;
        padding: 3px 6px 4px;
        vertical-align: top;
    }
</style>
<asp:UpdatePanel ID="upMegaInterdepartment" runat="server">
    <ContentTemplate>
        <table width="100%" cellpadding="2" cellspacing="0" runat="server">
            <tr>
                <td align="left" colspan="2">
                    <span style="color: Red;">* Indicates mandatory field</span>
                </td>
            </tr>
            <tr>
                <td height="10" colspan="2">
                </td>
            </tr>
            <tr>
                <td width="20%" class="LeftLabelTextFormat" valign="top">
                    <span style="color: Red;">*</span>Select Document
                </td>
                <td width="80%" class="ms-formbody">
                    <asp:DropDownList runat="server" ID="ddlDocuments" Width="318px" AutoPostBack="false"
                        Height="21px">
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="rfvddlDocuments" runat="server" ControlToValidate="ddlDocuments"
                        Display="Dynamic" ErrorMessage="Document is required" Font-Bold="True" ForeColor="Red"
                        InitialValue="Select" ValidationGroup="AddEmployee"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td width="20%" class="LeftLabelTextFormat" valign="top">
                    <span style="color: Red;">*</span>Select Department
                </td>
                <td width="80%" class="ms-formbody">
                    <asp:DropDownList runat="server" ID="ddlDepartment" Width="318px" AutoPostBack="true"
                        Height="21px" OnSelectedIndexChanged="ddlDepartment_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="rfvddlDepartment" runat="server" ControlToValidate="ddlDepartment"
                        Display="Dynamic" ErrorMessage="Department is required" Font-Bold="True" ForeColor="Red"
                        InitialValue="Select" ValidationGroup="AddEmployee"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td width="20%" class="LeftLabelTextFormat" valign="top">
                    <span style="color: Red;">*</span>Select Employee
                </td>
                <td width="80%" class="ms-formbody">
                    <asp:ListBox ID="lbEmployees" runat="server" Height="110px" Width="318px" SelectionMode="Multiple">
                    </asp:ListBox>
                    &nbsp;&nbsp;
                    <asp:Button ID="btnAdd" runat="server" Text="Add" OnClick="btnAdd_Click" ValidationGroup="AddEmployee" />
                    &nbsp;&nbsp;
                    <asp:Label ID="lblmessage" runat="server" Text="" Font-Bold="True" ForeColor="Red"></asp:Label>
                </td>
            </tr>
            <tr>
                <td width="20%" class="LeftLabelTextFormat" valign="top">
                </td>
                <td width="80%" class="ms-formbody">
                    <asp:Label ID="lblAddEmployee" runat="server" Text="" Font-Bold="True" ForeColor="Red"></asp:Label>
                </td>
            </tr>
            <tr>
                <td width="20%" class="LeftLabelTextFormat" valign="top">
                    &nbsp;
                </td>
                <td style="width: 80%">
                    <asp:GridView ID="gvEmployees" runat="server" AllowPaging="True" AllowSorting="True"
                        AutoGenerateColumns="False" CellPadding="4" EnableModelValidation="True" Width="600px"
                        OnRowDeleting="gvEmployees_RowDeleting" PageSize="10" ShowFooter="True" ForeColor="#333333"
                        GridLines="None" OnRowDataBound="gvEmployees_RowDataBound">
                        <AlternatingRowStyle BackColor="White" />
                        <Columns>
                            <asp:TemplateField HeaderText="Sr.No">
                                <ItemTemplate>
                                    <%# Container.DataItemIndex + 1 %>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Document Name" SortExpression="DocumentName">
                                <ItemTemplate>
                                    <asp:Label ID="lblDocumentName" runat="server" Text='<%# Bind("DocumentName") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Department" SortExpression="Department">
                                <ItemTemplate>
                                    <asp:Label ID="lblDepartment" runat="server" Text='<%# Bind("Department") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Employee Name" SortExpression="EmployeeName">
                                <ItemTemplate>
                                    <asp:Label ID="lblEmployeeName" runat="server" Text='<%# Bind("EmployeeName") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:CommandField ShowDeleteButton="true" />
                            <%--  <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Button ID="btnDelete" runat="server" Text="Delete" OnClientClick="javascript:return ConfirmDelete();" />
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>--%>
                        </Columns>
                        <EditRowStyle BackColor="#2461BF" />
                        <EmptyDataTemplate>
                            <span class="ms-vb">There is no pending item to show </span>
                        </EmptyDataTemplate>
                        <FooterStyle BackColor="#507CD1" ForeColor="White" Font-Bold="True" />
                        <HeaderStyle BackColor="#2589b4" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                        <RowStyle BackColor="#EFF3FB" />
                        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="2" height="10px">
                </td>
            </tr>
            <tr>
                <td colspan="2" height="10px" align="right">
                    <asp:Button runat="server" ID="btnSave" Text="Submit" Width="100px" OnClick="btnSave_Click"
                        ValidationGroup="AddEmployee" />&nbsp;&nbsp;
                    <asp:Button runat="server" ID="btnCancel" Text="Cancel" Width="100px"  
                        CausesValidation="false" onclick="btnCancel_Click" />
                </td>
            </tr>
        </table>
    </ContentTemplate>
</asp:UpdatePanel>
