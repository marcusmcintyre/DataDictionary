Imports System.Data.OleDb
Imports System.Data.Sql
Imports System
Imports System.Drawing.Color

Public Class _Default
    Inherits System.Web.UI.Page

    Const connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\Marcus\Documents\Database3.accdb;"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not (IsPostBack()) Then
            EnableDisableForm("False")
            populateTableDropDown()
            filterDataDictionaryByTable()
            populateDataDropDown("")
        End If

    End Sub

    Private Sub ddlTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTable.SelectedIndexChanged
        If (GridView1.EditIndex <> -1) Then
            GridView1.EditIndex = -1
        End If
        clearForm()
        filterDataDictionaryByTable()
    End Sub

#Region "GridView Events"

    Private Sub GridView1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        filterDataDictionaryByTable()
    End Sub
    Private Sub GridView1_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles GridView1.RowCommand

    End Sub

    Private Sub GridView1_SelectedIndexChanging(sender As Object, e As GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridView1.SelectedIndexChanged
        If (GridView1.SelectedIndex <> -1) Then
            EnableDisableForm("True")
            fillSelectedDetails()
        End If
    End Sub


#End Region

#Region "Button Events"
    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        updateEntry()
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        EnableDisableForm("True")
        populateDataDropDown("")
        btnAdd.Visible = False
        lblStatus.Text = "Adding New Entry"
        lblStatus.Visible = True
        tbTableName.Text = ddlTable.SelectedValue
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        btnAdd.Visible = True
        lblStatus.Visible = False
        clearForm()
        EnableDisableForm("False")
    End Sub
#End Region

#Region "Helper Subs/Functions"
    Private Sub updateEntry()
        Dim commandText = "UPDATE [DATA_DICTIONARY] SET  COLUMN_NAME = ? WHERE ID = " & GridView1.SelectedRow.Cells(1).Text
        Dim mySqlConn As New OleDbConnection(ConfigurationManager.ConnectionStrings("AccessConnection").ConnectionString)
        Dim sqlComm As New OleDbCommand()

        mySqlConn.Open()
        sqlComm.Parameters.AddWithValue("COLUMN_NAME", tbColumnName.Text)
        sqlComm.Connection = mySqlConn
        sqlComm.CommandText = commandText
        'SqlDataSource1.UpdateCommand = sqlComm.CommandText
        'SqlDataSource1.DataBind()
        sqlComm.ExecuteNonQuery()

        mySqlConn.Close()
    End Sub

    Private Sub fillSelectedDetails()
        clearForm()

        tbTableName.Text = GridView1.SelectedRow.Cells(2).Text
        tbColumnName.Text = GridView1.SelectedRow.Cells(3).Text
        ddlColumnType.SelectedIndex = populateDataDropDown(GridView1.SelectedRow.Cells(4).Text)
        tbColumnSize.Text = GridView1.SelectedRow.Cells(5).Text
        tbPrecision.Text = GridView1.SelectedRow.Cells(6).Text
        tbScale.Text = GridView1.SelectedRow.Cells(7).Text
        chkNullable.Checked = GridView1.SelectedRow.Cells(8).Text
        tbKey.Text = GridView1.SelectedRow.Cells(9).Text

        If Not (GridView1.SelectedRow.Cells(10).Text.Equals("&nbsp;")) Then
            tbDescription.Text = GridView1.SelectedRow.Cells(10).Text
        End If
    End Sub


    Private Function populateDataDropDown(ByVal str As String) As Integer
        Dim types As String() = ConfigurationManager.AppSettings("DataTypes").Split(" ")
        Dim data As String
        Dim index As Integer = 0

        For Each data In types
            ddlColumnType.Items.Add(data)
            If (data.Equals(str)) Then
                Return index
            End If
            index += 1
        Next

        Return 0
    End Function

    Private Sub populateTableDropDown()
        Dim mySqlConn As New OleDbConnection(ConfigurationManager.ConnectionStrings("AccessConnection").ConnectionString)
        Dim sqlComm As New OleDbCommand()
        Dim arrTable As New ArrayList()

        mySqlConn.Open()
        sqlComm.CommandText = "SELECT DISTINCT TABLE_NAME FROM [DATA_DICTIONARY] ORDER BY TABLE_NAME"
        sqlComm.Connection = mySqlConn
        arrTable.Add("")

        Dim reader As OleDbDataReader = sqlComm.ExecuteReader()
        While reader.Read()
            arrTable.Add(reader(0).ToString)
        End While

        ddlTable.DataSource = arrTable
        ddlTable.DataBind()
        mySqlConn.Close()
    End Sub

    Private Sub filterDataDictionaryByTable()
        If (ddlTable.SelectedIndex <> 0) Then
            GridView1.Columns(1).Visible = True
            GridView1.SelectedIndex = -1
            SqlDataSource1.SelectCommand = "SELECT * FROM [DATA_DICTIONARY] WHERE TABLE_NAME='" & ddlTable.SelectedValue & "' ORDER BY COLUMN_NAME"
            GridView1.DataBind()
            GridView1.Columns(1).Visible = False
            tbTableName.Text = ddlTable.SelectedValue
        End If
    End Sub

    Private Sub clearForm()
        tbTableName.Text = ""
        tbColumnName.Text = ""
        ddlColumnType.SelectedIndex = -1
        tbColumnSize.Text = ""
        tbPrecision.Text = ""
        tbScale.Text = ""
        chkNullable.Checked = False
        tbKey.Text = ""
        tbDescription.Text = ""
    End Sub

    Private Sub EnableDisableForm(ByVal value)
        'Input Controls
        tbTableName.Enabled = value
        tbColumnName.Enabled = value
        ddlColumnType.Enabled = value
        tbColumnSize.Enabled = value
        tbPrecision.Enabled = value
        tbScale.Enabled = value
        chkNullable.Enabled = value
        tbKey.Enabled = value
        tbDescription.Enabled = value

        'Validators
        rfvTableName.Enabled = value
        rfvColumnName.Enabled = value
        rfvColumnType.Enabled = value
        rfvColumnSize.Enabled = value
        rfvPrecision.Enabled = value
        rfvScale.Enabled = value
        rfvKey.Enabled = value
        rfvDescription.Enabled = value

        'Buttons
        btnSave.Enabled = value
        btnCancel.Enabled = value
    End Sub
#End Region
End Class