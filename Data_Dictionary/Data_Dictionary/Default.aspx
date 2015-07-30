<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="Data_Dictionary._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" type="text/css" href="Style.css" />
    <script type="text/javascript" src="home.js"></script>
    <title></title>
</head>
<body>
    <header>
        <a href="Default.aspx">
            <h1>Data Dictionary</h1>
        </a>
    </header>
    <form id="form1" runat="server">
        <div>
            <asp:Label ID="Label16" runat="server" Text="Choose Table: "></asp:Label>
            <asp:DropDownList ID="ddlTable" runat="server" AutoPostBack="True"></asp:DropDownList><br />
            <br />
            <asp:ListView ID="lvDictionary" runat="server" DataSourceID="SqlDataSource1" DataKeyNames="ID">
                <LayoutTemplate>
                    <table runat="server" id="tblDictionary" width="600px">
                        <tr>
                            <th></th>
                            <th>ID</th>
                            <th>Table Name</th>
                            <th>Column Name</th>
                            <th>Column Type</th>
                            <th>Column Size</th>
                            <th>Precision</th>
                            <th>Scale</th>
                            <th>Nullability</th>
                            <th>Key</th>
                            <th>Description</th>
                        </tr>
                        <tr runat="server" id="itemPlaceholder" />
                    </table>
                </LayoutTemplate>
                <ItemTemplate>
                    <tr id="Tr1" runat="server" class="Even">
                        <td>
                            <asp:Button ID="btnSelect" Text="Select" CausesValidation="false" CommandArgument='<%#DataBinder.Eval(Container.DataItem, "ID")%>' CommandName="Select" runat="server"></asp:Button>
                        </td>
                        <td>
                            <asp:Label ID="lblID" runat="server" Text='<%# Bind("ID") %>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblTN" runat="server" Text='<%# Bind("TABLE_NAME")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCN" runat="server" Text='<%# Bind("COLUMN_NAME")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCT" runat="server" Text='<%# Bind("COLUMN_TYPE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCS" runat="server" Text='<%# Bind("COLUMN_SIZE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblP" runat="server" Text='<%# Bind("PRECISION")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblS" runat="server" Text='<%# Bind("SCALE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:CheckBox ID="chkN" runat="server" Enabled="false" Checked='<%# Bind("NULLABILITY")%>'></asp:CheckBox>
                        </td>
                        <td>
                            <asp:Label ID="lblK" runat="server" Text='<%# Bind("KEY")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblD" runat="server" Text='<%# Bind("DESCRIPTION") %>'></asp:Label>
                        </td>
                    </tr>
                </ItemTemplate>
                <SelectedItemTemplate>
                    <tr id="Tr2" runat="server" class="Selected">
                        <td>
                            <asp:Button ID="btnSelect" Text="Select" CausesValidation="false" CommandArgument='<%#DataBinder.Eval(Container.DataItem, "ID")%>' CommandName="Select" runat="server"></asp:Button>
                        </td>
                        <td>
                            <asp:Label ID="lblID" runat="server" Text='<%# Bind("ID") %>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblTN" runat="server" Text='<%# Bind("TABLE_NAME")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCN" runat="server" Text='<%# Bind("COLUMN_NAME")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCT" runat="server" Text='<%# Bind("COLUMN_TYPE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCS" runat="server" Text='<%# Bind("COLUMN_SIZE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblP" runat="server" Text='<%# Bind("PRECISION")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblS" runat="server" Text='<%# Bind("SCALE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:CheckBox ID="chkN" runat="server" Enabled="false" Checked='<%# Bind("NULLABILITY")%>'></asp:CheckBox>
                        </td>
                        <td>
                            <asp:Label ID="lblK" runat="server" Text='<%# Bind("KEY")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblD" runat="server" Text='<%# Bind("DESCRIPTION") %>'></asp:Label>
                        </td>
                    </tr>
                </SelectedItemTemplate>
                <AlternatingItemTemplate>
                    <tr id="Tr2" runat="server" class="Odd">
                        <td>
                            <asp:Button ID="btnSelect" Text="Select" CausesValidation="false" CommandArgument='<%#DataBinder.Eval(Container.DataItem, "ID")%>' CommandName="Select" runat="server"></asp:Button>
                        </td>
                        <td>
                            <asp:Label ID="lblID" runat="server" Text='<%# Bind("ID") %>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblTN" runat="server" Text='<%# Bind("TABLE_NAME")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCN" runat="server" Text='<%# Bind("COLUMN_NAME")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCT" runat="server" Text='<%# Bind("COLUMN_TYPE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblCS" runat="server" Text='<%# Bind("COLUMN_SIZE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblP" runat="server" Text='<%# Bind("PRECISION")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblS" runat="server" Text='<%# Bind("SCALE")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:CheckBox ID="chkN" runat="server" Enabled="false" Checked='<%# Bind("NULLABILITY")%>'></asp:CheckBox>
                        </td>
                        <td>
                            <asp:Label ID="lblK" runat="server" Text='<%# Bind("KEY")%>'></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="lblD" runat="server" Text='<%# Bind("DESCRIPTION") %>'></asp:Label>
                        </td>
                    </tr>
                </AlternatingItemTemplate>
            </asp:ListView>

            <asp:SqlDataSource ID="SqlDataSource1" runat="server"
                ConnectionString="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Marcus\Documents\Database3.accdb"
                ProviderName="<%$ ConnectionStrings:AccessConnection.ProviderName %>"
                SelectCommand="SELECT * FROM [DATA_DICTIONARY] WHERE TABLE_NAME = @Table ORDER BY [COLUMN_NAME]">
                <SelectParameters>
                    <asp:ControlParameter Name="Table"
                        ControlID="ddlTable"
                        PropertyName="SelectedValue" />
                </SelectParameters>
            </asp:SqlDataSource>
            <br />
            <asp:Button ID="btnCancel" runat="server" Text="Cancel" CausesValidation="false" />
            <asp:Button ID="btnAdd" runat="server" Text="Add New Entry" CausesValidation="false" />
            <asp:Button ID="btnDelete" runat="server" OnClientClick="return confirmation();" Text="Delete" CausesValidation="false" />
            <br />
            <asp:Label ID="lblStatus" runat="server" Text="Adding" Visible="false" />
            <br />
            <br />

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
        </div>
    </form>
</body>
</html>
