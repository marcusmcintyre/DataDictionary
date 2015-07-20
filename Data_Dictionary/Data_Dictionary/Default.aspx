<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="Data_Dictionary._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Label ID="Label0" runat="server" Text="Data Dictionary" Font-Size="X-Large"></asp:Label><br />
        <asp:Label ID="Label16" runat="server" Text="Choose Table: "></asp:Label>
        <asp:DropDownList ID="ddlTable" runat="server" AutoPostBack="True"></asp:DropDownList><br /><br />
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False" DataKeyNames="ID" DataSourceID="SqlDataSource1" PageSize="5">
            <Columns>
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:Button ID="Button1" runat="server" CausesValidation="True" CommandName="Update" Text="Update" />
                        &nbsp;<asp:Button ID="Button2" runat="server" CausesValidation="False" CommandName="Cancel" Text="Cancel" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Button ID="Button1" runat="server" CausesValidation="False" CommandName="Edit" Text="Edit" />
                    </ItemTemplate>
                    <ControlStyle Width="75px" />
                    <ItemStyle HorizontalAlign="Center" Width="75px" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="TABLE_NAME" SortExpression="TABLE_NAME">
                    <EditItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("TABLE_NAME") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("TABLE_NAME") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="COLUMN_NAME" SortExpression="COLUMN_NAME">
                    <EditItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("COLUMN_NAME") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("COLUMN_NAME") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="COLUMN_TYPE" SortExpression="COLUMN_TYPE">
                    <EditItemTemplate>
                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("COLUMN_TYPE") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label6" runat="server" Text='<%# Bind("COLUMN_TYPE") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="COLUMN_SIZE" SortExpression="COLUMN_SIZE" ItemStyle-HorizontalAlign="Center">
                    <EditItemTemplate>
                        <asp:Label ID="Label7" runat="server" Text='<%# Bind("COLUMN_SIZE") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label8" runat="server" Text='<%# Bind("COLUMN_SIZE") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="PRECISION" SortExpression="PRECISION" ItemStyle-HorizontalAlign="Center">
                    <EditItemTemplate>
                        <asp:Label ID="Label9" runat="server" Text='<%# Bind("PRECISION") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label10" runat="server" Text='<%# Bind("PRECISION") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle HorizontalAlign="Center" />
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="SCALE" SortExpression="SCALE" ItemStyle-HorizontalAlign="Center">
                    <EditItemTemplate>
                        <asp:Label ID="Label11" runat="server" Text='<%# Bind("SCALE") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label12" runat="server" Text='<%# Bind("SCALE") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="NULLABILITY" SortExpression="NULLABILITY" ItemStyle-HorizontalAlign="Center">
                    <EditItemTemplate>
                        <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# Bind("NULLABILITY") %>' Enabled="false" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("NULLABILITY") %>' Enabled="false" />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="KEY" SortExpression="KEY" ItemStyle-HorizontalAlign="Center">
                    <EditItemTemplate>
                        <asp:Label ID="Label13" runat="server" Text='<%# Bind("KEY") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label14" runat="server" Text='<%# Bind("KEY") %>'></asp:Label>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="DESCRIPTION" SortExpression="DESCRIPTION">
                    <EditItemTemplate>
                        <asp:TextBox ID="tbEditDescription" runat="server" Text='<%# Bind("DESCRIPTION") %>'></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label15" runat="server" Text='<%# Bind("DESCRIPTION") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\Documents\Database3.accdb" ProviderName="<%$ ConnectionStrings:AccessConnection.ProviderName %>" SelectCommand="SELECT * FROM [DATA_DICTIONARY] WHERE TABLE_NAME=' '"></asp:SqlDataSource><br />
        <asp:Button ID="btnAdd" runat="server" Text="Add New Entry" /><br /><br />

        <table>
            <tr>
                <td>
                    <asp:Label ID="lblTableName" runat="server" Text="Table Name: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbTableName" runat="server" MaxLength="255"></asp:TextBox>
                    <asp:RequiredFieldValidator id="rfvTableName" runat="server" ControlToValidate="tbTableName" ErrorMessage="You must enter a Table Name.">*</asp:RequiredFieldValidator>
                </td>
            </tr>     
            <tr>
                <td>
                    <asp:Label ID="lblColumnName" runat="server" Text="Column Name: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbColumnName" runat="server" MaxLength="255"></asp:TextBox>
                    <asp:RequiredFieldValidator id="rfvColumnName" runat="server" ControlToValidate="tbColumnName" ErrorMessage="You must enter a Column Name.">*</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblColumnType" runat="server" Text="Column Type: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbColumnType" runat="server" MaxLength="255"></asp:TextBox>
                    <asp:RequiredFieldValidator id="rfvColumnType" runat="server" ControlToValidate="tbColumnType" ErrorMessage="You must enter a Column Type.">*</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblColumnSize" runat="server" Text="Column Size: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbColumnSize" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                    <asp:RequiredFieldValidator id="rfvColumnSize" runat="server" ControlToValidate="tbColumnSize" ErrorMessage="You must enter a Column Size.">*</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblPrecision" runat="server" Text="Precision: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbPrecision" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                    <asp:RequiredFieldValidator id="rfvPrecision" runat="server" ControlToValidate="tbPrecision" ErrorMessage="You must enter a Precision value.">*</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblScale" runat="server" Text="Scale: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbScale" runat="server" TextMode="Number" MaxLength="9"></asp:TextBox>
                    <asp:RequiredFieldValidator id="rfvScale" runat="server" ControlToValidate="tbScale" ErrorMessage="You must enter a Scale.">*</asp:RequiredFieldValidator>
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
                    <asp:RequiredFieldValidator id="rfvKey" runat="server" ControlToValidate="tbKey" ErrorMessage="You must enter a Key.">*</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblDescription" runat="server" Text="Description: "></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbDescription" runat="server" Rows="4" TextMode="MultiLine" MaxLength="64000"></asp:TextBox>
                    <asp:RequiredFieldValidator id="rfvDescription" runat="server" ControlToValidate="tbDescription" ErrorMessage="You must enter a Description.">*</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                    
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
