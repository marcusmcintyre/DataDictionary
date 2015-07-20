Imports System.Data.OleDb
Imports System.Data.Sql
Imports System

Public Class _Default
    Inherits System.Web.UI.Page

    Const connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\user\Documents\Database3.accdb;"
    Const strTable = "test"

    Dim arrTables As ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not (IsPostBack()) Then
            EnableDisableAddForm("False")
            populateDropDown()
            filterDataDictionaryByTable()
        End If

    End Sub

    Private Sub EnableDisableAddForm(ByVal value)
        'Input Controls
        tbTableName.Enabled = value
        tbColumnName.Enabled = value
        tbColumnType.Enabled = value
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
            SqlDataSource1.SelectCommand = "SELECT * FROM [DATA_DICTIONARY] WHERE TABLE_NAME='" & ddlTable.SelectedValue & "' ORDER BY COLUMN_NAME"
            GridView1.DataBind()
        End If
    End Sub

    Private Sub updateDescription(des As String)
        Dim mySqlConn As New OleDbConnection(connStr)
        Dim sqlComm As New OleDbCommand()

        mySqlConn.Open()
        sqlComm.Parameters.Add("", OleDbType.VarChar)
        sqlComm.CommandText = "UPDATE [DATA_DICTIONARY] SET DESCIPTION='" & des & "' WHERE TABLE_NAME = '" & ddlTable.SelectedValue & "';"
        sqlComm.Connection = mySqlConn
        SqlDataSource1.UpdateCommand = sqlComm.CommandText
        sqlComm.ExecuteNonQuery()

        mySqlConn.Close()
    End Sub

    Private Sub ddlTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTable.SelectedIndexChanged
        If (GridView1.EditIndex <> -1) Then
            GridView1.EditIndex = -1
        End If
        filterDataDictionaryByTable()
    End Sub

    Private Sub GridView1_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit
        filterDataDictionaryByTable()
    End Sub

    Private Sub GridView1_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles GridView1.RowEditing
        filterDataDictionaryByTable()
    End Sub

    Private Sub GridView1_RowUpdating(sender As Object, e As GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim mySqlConn As New OleDbConnection(connStr)
        Dim sqlComm As New OleDbCommand()

        mySqlConn.Open()

        For Each entry As DictionaryEntry In e.NewValues
            'TODO

            'Parse the entries to get their Data Type and add them to the parameters

            If Not entry.Value.GetType() = GetType(Boolean) Then

            End If
            Select Case entry.Value.GetType()
                Case GetType(Boolean)
                    If entry.Value = True Then
                        sqlComm.Parameters.Add("True", OleDbType.Boolean)
                    Else
                        sqlComm.Parameters.Add("False", OleDbType.Boolean)
                    End If

                Case GetType(Long)
                    sqlComm.Parameters.Add(entry.Value, OleDbType.Numeric)

                Case GetType(String)
                    sqlComm.Parameters.Add(entry.Value.ToString, OleDbType.VarChar)

                Case Else
            End Select
        Next
        DataBind()
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridView1.SelectedIndexChanged
        GridView1.EditIndex = GridView1.SelectedIndex
    End Sub

    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        

        DataBind()
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        EnableDisableAddForm("True")
    End Sub
End Class