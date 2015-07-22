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
            populateDropDown()
            filterDataDictionaryByTable()
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
        DataBind()
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        EnableDisableForm("True")
        tbTableName.Text = ddlTable.SelectedValue
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        EnableDisableForm("True")
        clearForm()
    End Sub
#End Region

#Region "Helper Subs/Functions"
    Private Sub updateDescription(des As String)
        Dim commandText = "UPDATE " & ddlTable.SelectedValue & " SET " & "='@val' WHERE ID = '@row'"
        Dim mySqlConn As New OleDbConnection(connStr)
        Dim sqlComm As New OleDbCommand()

        mySqlConn.Open()
        'sqlComm.Parameters.Add("@field", OleDbType.)
        'sqlComm.CommandText = commandText
        'sqlComm.Parameters.Add(
        'sqlComm.Connection = mySqlConn
        SqlDataSource1.UpdateCommand = sqlComm.CommandText
        sqlComm.ExecuteNonQuery()

        mySqlConn.Close()
    End Sub

    Private Sub fillSelectedDetails()
        clearForm()

        tbTableName.Text = GridView1.SelectedRow.Cells(1).Text
        tbColumnName.Text = GridView1.SelectedRow.Cells(2).Text
        ddlColumnType.SelectedValue = GridView1.SelectedRow.Cells(3).Text
        tbColumnSize.Text = GridView1.SelectedRow.Cells(4).Text
        tbPrecision.Text = GridView1.SelectedRow.Cells(5).Text
        tbScale.Text = GridView1.SelectedRow.Cells(6).Text
        chkNullable.Checked = GridView1.SelectedRow.Cells(7).Text
        tbKey.Text = GridView1.SelectedRow.Cells(8).Text
        If Not (GridView1.SelectedRow.Cells(9).Text.Equals("&nbsp;")) Then
            tbDescription.Text = GridView1.SelectedRow.Cells(9).Text
        End If

    End Sub

    Private Sub populateDropDown()
        Dim mySqlConn As New OleDbConnection(connStr)
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
            GridView1.SelectedIndex = -1
            SqlDataSource1.SelectCommand = "SELECT * FROM [DATA_DICTIONARY] WHERE TABLE_NAME='" & ddlTable.SelectedValue & "' ORDER BY COLUMN_NAME"
            GridView1.DataBind()
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
        btnClear.Enabled = value
    End Sub
#End Region
End Class