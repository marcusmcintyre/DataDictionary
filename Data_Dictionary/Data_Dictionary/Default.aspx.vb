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
            populateTableDropDown("")
            filterDataDictionaryByTable()
            populateDataDropDown("")
        End If

    End Sub

    Private Sub ddlTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTable.SelectedIndexChanged
        If (GridView1.SelectedIndex <> -1) Then
            GridView1.SelectedIndex = -1
        End If
        clearForm()
        filterDataDictionaryByTable()
    End Sub

#Region "GridView Events"
    Private Sub GridView1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.SelectedIndex = -1
        clearForm()
        EnableDisableForm("False")
        filterDataDictionaryByTable()
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridView1.SelectedIndexChanged
        If (GridView1.SelectedIndex <> -1) Then
            lblStatus.Text = "Updating Entry: " & GridView1.Rows(GridView1.SelectedIndex).Cells(2).Text & ", " & GridView1.Rows(GridView1.SelectedIndex).Cells(3).Text
            EnableDisableForm("True")
            fillSelectedDetails()
        End If
    End Sub
#End Region

#Region "Button Events"
    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If GridView1.SelectedIndex <> -1 Then
            modifyEntry("updateCommand")
        Else
            modifyEntry("insertCommand")
        End If
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        lblStatus.Text = "Adding New Entry"
        tbTableName.Text = ddlTable.SelectedValue
        EnableDisableForm("True")
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        btnAdd.Visible = True
        btnCancel.Visible = False
        lblStatus.Visible = False
        GridView1.SelectedIndex = -1
        clearForm()
        EnableDisableForm("False")
    End Sub
#End Region

#Region "Helper Subs/Functions"
    Private Sub modifyEntry(ByVal comm As String)
        Dim mySqlConn As New OleDbConnection(ConfigurationManager.ConnectionStrings("AccessConnection").ConnectionString)
        Dim sqlComm As New OleDbCommand

        sqlComm.Parameters.AddWithValue("TABLE_NAME", tbTableName.Text)
        sqlComm.Parameters.AddWithValue("COLUMN_NAME", tbColumnName.Text)
        sqlComm.Parameters.AddWithValue("COLUMN_TYPE", ddlColumnType.SelectedValue)
        sqlComm.Parameters.AddWithValue("COLUMN_SIZE", tbColumnSize.Text)
        sqlComm.Parameters.AddWithValue("PRECISION", tbPrecision.Text)
        sqlComm.Parameters.AddWithValue("SCALE", tbScale.Text)
        sqlComm.Parameters.AddWithValue("NULLABILITY", chkNullable.Checked)
        sqlComm.Parameters.AddWithValue("KEY", tbKey.Text)
        sqlComm.Parameters.AddWithValue("DESCRIPTION", tbDescription.Text)
        If GridView1.SelectedIndex <> -1 Then
            Dim id As Label = GridView1.SelectedRow.FindControl("lblID")
            sqlComm.Parameters.AddWithValue("ID", id.Text)
        End If

        mySqlConn.Open()
        sqlComm.Connection = mySqlConn
        sqlComm.CommandText = ConfigurationManager.AppSettings(comm)
        sqlComm.ExecuteNonQuery()
        mySqlConn.Close()

        ddlTable.SelectedIndex = populateTableDropDown(tbTableName.Text)
        filterDataDictionaryByTable()
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

    Private Function populateTableDropDown(ByVal str As String) As Integer
        Dim mySqlConn As New OleDbConnection(ConfigurationManager.ConnectionStrings("AccessConnection").ConnectionString)
        Dim sqlComm As New OleDbCommand
        Dim arrTable As New ArrayList
        Dim index As Integer = 1

        mySqlConn.Open()
        sqlComm.CommandText = "SELECT DISTINCT [TABLE_NAME] FROM [DATA_DICTIONARY] ORDER BY [TABLE_NAME]"
        sqlComm.Connection = mySqlConn
        arrTable.Add("")

        Dim reader As OleDbDataReader = sqlComm.ExecuteReader()
        While reader.Read()
            arrTable.Add(reader(0).ToString)
            If Not (reader(0).ToString.Equals(str)) Then
                index += 1
            End If
        End While

        ddlTable.DataSource = arrTable
        ddlTable.DataBind()
        mySqlConn.Close()

        Return index
    End Function

    Private Sub filterDataDictionaryByTable()
        If (ddlTable.SelectedIndex <> 0) Then
            SqlDataSource1.SelectCommand = "SELECT * FROM [DATA_DICTIONARY] WHERE [TABLE_NAME]='" & ddlTable.SelectedValue & "' ORDER BY [COLUMN_NAME]"
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
        btnCancel.Visible = value
        btnAdd.Visible = Not CType(value, Boolean)

        lblStatus.Visible = value
    End Sub
#End Region
End Class