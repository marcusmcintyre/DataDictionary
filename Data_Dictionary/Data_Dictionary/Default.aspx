<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="Data_Dictionary._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" runat="server">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style.css" />
    <script type="text/javascript" src="home.js"></script>
    <title>Data Dictionary</title>
</head>

<body>
    <form id="form1" runat="server">

        <asp:SqlDataSource ID="SQLServer" runat="server" ConnectionString="<%$ ConnectionStrings:SQLConnection %>"></asp:SqlDataSource>
        <asp:AccessDataSource ID="AccessDataSource1" runat="server" DataFile="~/Database1.accdb" SelectCommand="SELECT * FROM [DATA_DICTIONARY] WHERE [ID] = 0"></asp:AccessDataSource>

        <header>
            <asp:Panel ID="pnlSearch" runat="server" DefaultButton="btnSearch">
                <table>
                    <tr>
                        <td>
                            <asp:Label ID="Label16" runat="server" Text="Choose Table: "></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="Label1" runat="server" Text="Choose Column: "></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:DropDownList ID="ddlTable" runat="server" AutoPostBack="True"></asp:DropDownList>
                        </td>
                        <td>
                            <asp:TextBox ID="tbSearch" runat="server"></asp:TextBox>
                            <asp:Label ID="lblExactMatch" runat="server" Text="Exact Match: "></asp:Label>
                            <asp:CheckBox ID="chkExactMatch" runat="server" />
                            <asp:Button ID="btnClear" runat="server" Text="Clear" CausesValidation="false" />
                            <asp:Button ID="btnSearch" runat="server" Text="Search" CausesValidation="false" />
                            <asp:Label ID="lblSQL" runat="server" Text="SQL Server"></asp:Label><asp:RadioButton ID="rbSQL" runat="server" AutoPostBack="true" GroupName="radio"/>
                            <asp:Label ID="lblAccess" runat="server" Text="Access"></asp:Label><asp:RadioButton ID="rbAccess" runat="server" AutoPostBack="true" GroupName="radio" Checked="true"/>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
        </header>

        <div class="wrapper" runat="server">
            <article>
                <asp:GridView ID="gvDictionary" runat="server" AutoGenerateColumns="False" DataKeyNames="ID">
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText="    " ShowSelectButton="True" />
                        <asp:BoundField DataField="ID" HeaderText="ID" />
                        <asp:BoundField DataField="TABLE_NAME" HeaderText="Table Name" />
                        <asp:BoundField DataField="COLUMN_NAME" HeaderText="Column Name" />
                        <asp:BoundField DataField="COLUMN_TYPE" HeaderText="Column Type" />
                        <asp:BoundField DataField="COLUMN_SIZE" HeaderText="Column Size">
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="PRECISION" HeaderText="Precision">
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SCALE" HeaderText="Scale">
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="NULLABILITY" HeaderText="Nullability">
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="KEY" HeaderText="Key">
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DESCRIPTION" HeaderText="Description" />
                    </Columns>

                    <RowStyle BackColor="White" />
                    <AlternatingRowStyle BackColor="#F5F5F5" />
                    <SelectedRowStyle BackColor="LightCyan" />
                </asp:GridView>
            </article>
        </div>

        <div class="footer">
            <table class="Details">
                <tr>
                    <asp:Label ID="lblStatus" runat="server" Visible="false" />
                    <br />
                    <asp:Label ID="lblCurrID" runat="server" Visible="false" />
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblTableName" runat="server" Text="Table Name: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbTableName" runat="server" MaxLength="255"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvTableName" CssClass="validator" runat="server" ControlToValidate="tbTableName" ErrorMessage="You must enter a Table Name.">*</asp:RequiredFieldValidator>
                    </td>
                    <td>
                        <asp:Label ID="lblColumnType" runat="server" Text="Column Type: "></asp:Label>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlColumnType" runat="server" AutoPostBack="True"></asp:DropDownList>
                        <asp:RequiredFieldValidator ID="rfvColumnType" CssClass="validator" runat="server" ControlToValidate="ddlColumnType"></asp:RequiredFieldValidator>
                    </td>
                    <td>
                        <asp:Label ID="lblDescription" runat="server" Text="Description: "></asp:Label>
                    </td>
                    <td><asp:Button ID="btnCancel" runat="server" Text="Cancel" CausesValidation="false" /></td>
                    <td><asp:Button ID="btnDelete" runat="server" OnClientClick="return confirmation();" Text="Delete" CausesValidation="false" /></td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblColumnName" runat="server" Text="Column Name: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbColumnName" runat="server" MaxLength="255"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvColumnName" CssClass="validator" runat="server" ControlToValidate="tbColumnName" ErrorMessage="You must enter a Column Name.">*</asp:RequiredFieldValidator>
                    </td>
                    <td>
                        <asp:Label ID="lblNullable" runat="server" Text="Nullable: "></asp:Label>
                    </td>
                    <td>
                        <asp:CheckBox ID="chkNullable" runat="server"></asp:CheckBox>
                    </td>
                    <td rowspan="3">
                        <asp:TextBox ID="tbDescription" runat="server" Rows="4" TextMode="MultiLine" MaxLength="64000"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvDescription" CssClass="validator" runat="server" ControlToValidate="tbDescription" ErrorMessage="You must enter a Description.">*</asp:RequiredFieldValidator>
                    </td>
                    <td><asp:Button ID="btnAdd" runat="server" Text="Add New Entry" CausesValidation="false" /></td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblColumnSize" runat="server" Text="Column Size: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbColumnSize" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvColumnSize" CssClass="validator" runat="server" ControlToValidate="tbColumnSize" ErrorMessage="You must enter a Column Size.">*</asp:RequiredFieldValidator>
                    </td>
                    <td>
                        <asp:Label ID="lblPrecision" runat="server" Text="Precision: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbPrecision" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvPrecision" CssClass="validator" runat="server" ControlToValidate="tbPrecision" ErrorMessage="You must enter a Precision value.">*</asp:RequiredFieldValidator>
                        <asp:RangeValidator ID="RangeValidator2" CssClass="validator" runat="server" ControlToValidate="tbPrecision" MinimumValue="0" MaximumValue="65535" ErrorMessage="Precision must be within this range (0 - 65535)">*</asp:RangeValidator>
                    </td>
                    <td>
                        <asp:Button ID="btnApplyDescription" runat="server" Text="Apply to All Results"  CausesValidation="false"/>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblKey" runat="server" Text="Key: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbKey" runat="server" TextMode="Number" MaxLength="1"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvKey" CssClass="validator" runat="server" ControlToValidate="tbKey" ErrorMessage="You must enter a Key.">*</asp:RequiredFieldValidator>
                    </td>
                    <td>
                        <asp:Label ID="lblScale" runat="server" Text="Scale: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbScale" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvScale" CssClass="validator" runat="server" ControlToValidate="tbScale" ErrorMessage="You must enter a Scale.">*</asp:RequiredFieldValidator>
                        <asp:RangeValidator ID="RangeValidator1" CssClass="validator" runat="server" ControlToValidate="tbScale" MinimumValue="0" MaximumValue="65535" ErrorMessage="Scale must be within this range (0 - 65535)">*</asp:RangeValidator>
                    </td>
                    <td>
                        <asp:Button ID="btnSave" runat="server" Text="Save" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
