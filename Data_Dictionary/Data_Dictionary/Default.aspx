<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="Data_Dictionary._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" runat="server">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style.css" />
    <script type="text/javascript" src="home.js"></script>
    <title></title>
</head>
<body id="body" runat="server">
    <header>
        <a href="Default.aspx">
            <h1>Data Dictionary</h1>
        </a>
    </header>
    <form id="form1" runat="server">
        <div id="contain" runat="server">
                        <asp:Label ID="Label16" runat="server" Text="Choose Table: "></asp:Label>
            <asp:DropDownList ID="ddlTable" runat="server" AutoPostBack="True"></asp:DropDownList>
            <asp:TextBox ID="tbSearch" runat="server"></asp:TextBox>
            <asp:Button ID="btnSearch" runat="server" Text="Search" CausesValidation="false" />
            <asp:Label ID="lblExactMatch" runat="server" Text="Exact Match: "></asp:Label>
            <asp:CheckBox ID="chkExactMatch" runat="server" /><br />
            <br />
            <asp:GridView ID="gvDictionary" runat="server" AutoGenerateColumns="False" DataKeyNames="ID" DataSourceID="SqlDataSource1">
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

            <table class="Details">
                <tr>
                    <td>
                        <asp:Label ID="lblTableName" runat="server" Text="Table Name: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbTableName" runat="server" MaxLength="255"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvTableName" runat="server" ControlToValidate="tbTableName" ErrorMessage="You must enter a Table Name.">*</asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblColumnName" runat="server" Text="Column Name: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbColumnName" runat="server" MaxLength="255"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvColumnName" runat="server" ControlToValidate="tbColumnName" ErrorMessage="You must enter a Column Name.">*</asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblColumnType" runat="server" Text="Column Type: "></asp:Label>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlColumnType" runat="server" AutoPostBack="True"></asp:DropDownList>
                        <asp:RequiredFieldValidator ID="rfvColumnType" runat="server" ControlToValidate="ddlColumnType"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblColumnSize" runat="server" Text="Column Size: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbColumnSize" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvColumnSize" runat="server" ControlToValidate="tbColumnSize" ErrorMessage="You must enter a Column Size.">*</asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr id="rowPrecision" runat="server">
                    <td>
                        <asp:Label ID="lblPrecision" runat="server" Text="Precision: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbPrecision" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvPrecision" runat="server" ControlToValidate="tbPrecision" ErrorMessage="You must enter a Precision value.">*</asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr id="rowScale" runat="server">
                    <td>
                        <asp:Label ID="lblScale" runat="server" Text="Scale: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbScale" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvScale" runat="server" ControlToValidate="tbScale" ErrorMessage="You must enter a Scale.">*</asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblNullable" runat="server" Text="Nullable: "></asp:Label>
                    </td>
                    <td>
                        <asp:CheckBox ID="chkNullable" runat="server"></asp:CheckBox>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblKey" runat="server" Text="Key: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbKey" runat="server" TextMode="Number" MaxLength="1"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvKey" runat="server" ControlToValidate="tbKey" ErrorMessage="You must enter a Key.">*</asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblDescription" runat="server" Text="Description: "></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbDescription" runat="server" Rows="4" TextMode="MultiLine" MaxLength="64000"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvDescription" runat="server" ControlToValidate="tbDescription" ErrorMessage="You must enter a Description.">*</asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td>
                        <asp:Button ID="btnSave" runat="server" Text="Save" />
                    </td>
                </tr>
            </table>

            <br />
            <asp:Button ID="btnCancel" runat="server" Text="Cancel" CausesValidation="false" />
            <asp:Button ID="btnAdd" runat="server" Text="Add New Entry" CausesValidation="false" />
            <asp:Button ID="btnDelete" runat="server" OnClientClick="return confirmation();" Text="Delete" CausesValidation="false" />
            <br />
            <asp:Label ID="lblStatus" runat="server" Visible="false" />
            <br />
            <asp:Label ID="lblCurrID" runat="server" Visible="false" />
            <br />
            <br />
            <asp:SqlDataSource ID="SqlDataSource1" runat="server"
                ConnectionString="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB\Database3.accdb"
                ProviderName="<%$ ConnectionStrings:AccessConnection.ProviderName %>"
                SelectCommand="SELECT * FROM [DATA_DICTIONARY] WHERE TABLE_NAME = @Table ORDER BY [COLUMN_NAME]">
                <SelectParameters>
                    <asp:ControlParameter Name="Table"
                        ControlID="ddlTable"
                        PropertyName="SelectedValue" />
                </SelectParameters>
            </asp:SqlDataSource>      
        </div>
    </form>
</body>
</html>
