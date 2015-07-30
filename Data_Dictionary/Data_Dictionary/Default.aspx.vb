Imports System.Data.OleDb
Imports System.Data.Sql
Imports System
Imports System.Drawing.Color

Public Class _Default
    Inherits System.Web.UI.Page

    Const connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\Marcus\Documents\Database3.accdb;"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not (IsPostBack()) Then
            ViewState("ps") = 0
            EnableDisableForm("False")
            populateTableDropDown()
            filterDataDictionaryByTable()
            populateDataDropDown("")
        End If

    End Sub

    Private Sub ddlTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTable.SelectedIndexChanged
        If (lvDictionary.SelectedIndex <> -1) Then
            lvDictionary.SelectedIndex = -1
        End If
        clearForm()
        filterDataDictionaryByTable()
    End Sub


    Private Sub ddlColumnType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlColumnType.SelectedIndexChanged
        checkDataDropDown()
    End Sub

    Private Function checkDataDropDown() As Boolean
        If (ddlColumnType.SelectedIndex = populateDataDropDown("numeric(p,s)") _
            Or ddlColumnType.SelectedIndex = populateDataDropDown("decimal(p,s)")) Then
            If (ViewState("ps") = 0) Then
                tbPrecision.Text = ""
                tbScale.Text = ""
                rfvPrecision.Enabled = True
                rfvScale.Enabled = True
                tbPrecision.Enabled = True
                tbScale.Enabled = True
                ViewState("ps") = 1
            End If
            Return True
        Else
            tbPrecision.Text = ""
            tbScale.Text = ""
            rfvPrecision.Enabled = False
            rfvScale.Enabled = False
            tbPrecision.Enabled = False
            tbScale.Enabled = False
            ViewState("ps") = 0
            Return False
        End If
        Return False
    End Function

#Region "GridView Events"

    Protected Sub lvDictionary_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvDictionary.SelectedIndexChanged
        If (lvDictionary.SelectedIndex <> -1) Then
            Dim id As Label = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblID")
            Dim table As Label = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblTN")
            Dim column As Label = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblCN")
            lblStatus.Text = "Updating Entry (" & id.Text & "): " & table.Text & ", " & column.Text
            EnableDisableForm("True")
            fillSelectedDetails()
        End If
    End Sub
#End Region

#Region "Button Events"
    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim comm As String = ""
        If lvDictionary.SelectedIndex <> -1 Then
            comm = "updateCommand"
        Else
            comm = "insertCommand"
        End If
        modifyEntry(comm)
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        lblStatus.Text = "Adding New Entry"
        clearForm()
        tbTableName.Text = ddlTable.SelectedValue
        ddlColumnType.SelectedIndex = -1
        lvDictionary.SelectedIndex = -1
        EnableDisableForm("True")
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If (lvDictionary.SelectedIndex = -1) Then
            EnableDisableForm("False")
        End If
        clearForm()
        filterDataDictionaryByTable()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        modifyEntry("deleteCommand")
    End Sub
#End Region

#Region "Helper Subs/Functions"
    Private Sub modifyEntry(ByVal comm As String)
        Dim mySqlConn As New OleDbConnection(ConfigurationManager.ConnectionStrings("AccessConnection").ConnectionString)
        Dim sqlComm As New OleDbCommand

        If (comm.Equals("deleteCommand")) Then
            Dim id As Label = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblID")
            sqlComm.Parameters.AddWithValue("ID", id.Text)
        Else

            sqlComm.Parameters.AddWithValue("TABLE_NAME", tbTableName.Text)
            sqlComm.Parameters.AddWithValue("COLUMN_NAME", tbColumnName.Text)
            sqlComm.Parameters.AddWithValue("COLUMN_TYPE", ddlColumnType.SelectedValue)
            sqlComm.Parameters.AddWithValue("COLUMN_SIZE", tbColumnSize.Text)
            If (checkDataDropDown()) Then
                sqlComm.Parameters.AddWithValue("PRECISION", tbPrecision.Text)
                sqlComm.Parameters.AddWithValue("SCALE", tbScale.Text)
            Else
                sqlComm.Parameters.AddWithValue("PRECISION", 0)
                sqlComm.Parameters.AddWithValue("SCALE", 0)
            End If
            sqlComm.Parameters.AddWithValue("NULLABILITY", chkNullable.Checked)
            sqlComm.Parameters.AddWithValue("KEY", tbKey.Text)
            sqlComm.Parameters.AddWithValue("DESCRIPTION", tbDescription.Text)
            If lvDictionary.SelectedIndex <> -1 Then
                Dim id As Label = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblID")
                sqlComm.Parameters.AddWithValue("ID", id.Text)
            End If
        End If

        mySqlConn.Open()
        sqlComm.Connection = mySqlConn
        sqlComm.CommandText = ConfigurationManager.AppSettings(comm)
        sqlComm.ExecuteNonQuery()
        mySqlConn.Close()

        populateTableDropDown()
        ddlTable.SelectedValue = tbTableName.Text
        filterDataDictionaryByTable()
        selectlv(lvDictionary, tbColumnName.Text)
    End Sub

    Private Sub selectlv(ByRef lv As ListView, ByVal str As String)
        Dim x, i As Integer
        Dim lblCol As Label
        Dim intRowFound As Integer

        If lv.Items.Count > 0 Then
            'EACH ROW
            For x = 0 To lv.Items.Count - 1
                lblCol = lv.Items(x).FindControl("lblCN")
                'FOUND IT
                If lblCol.Text.Equals(str) Then
                    intRowFound = x
                End If
            Next
        End If

        lv.SelectedIndex = intRowFound
        lv.DataBind()

    End Sub

    Private Sub fillSelectedDetails()
        clearForm()
        Dim lbl As Label
        Dim chk As CheckBox

        lbl = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblTN")
        tbTableName.Text = lbl.Text

        lbl = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblCN")
        tbColumnName.Text = lbl.Text

        lbl = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblCT")
        ddlColumnType.SelectedIndex = populateDataDropDown(lbl.Text)

        lbl = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblCS")
        tbColumnSize.Text = lbl.Text

        lbl = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblP")
        tbPrecision.Text = lbl.Text

        lbl = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblS")
        tbScale.Text = lbl.Text

        chk = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("chkN")
        chkNullable.Checked = chk.Checked

        lbl = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblK")
        tbKey.Text = lbl.Text

        lbl = lvDictionary.Items(lvDictionary.SelectedIndex).FindControl("lblD")
        If Not (lbl.Text.Equals("&nbsp;")) Then
            tbDescription.Text = lbl.Text
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
        Dim sqlComm As New OleDbCommand
        Dim arrTable As New ArrayList

        mySqlConn.Open()
        sqlComm.CommandText = "SELECT DISTINCT [TABLE_NAME] FROM [DATA_DICTIONARY] ORDER BY [TABLE_NAME]"
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
            lvDictionary.DataBind()
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
        chkNullable.Enabled = value
        tbKey.Enabled = value
        tbDescription.Enabled = value

        If (value) Then
            checkDataDropDown()
        Else
            tbPrecision.Enabled = False
            tbScale.Enabled = False
        End If


        'Validators
        rfvTableName.Enabled = value
        rfvColumnName.Enabled = value
        rfvColumnType.Enabled = value
        rfvColumnSize.Enabled = value
        rfvKey.Enabled = value
        rfvDescription.Enabled = value

        'Buttons
        btnSave.Enabled = value
        btnCancel.Enabled = value
        If (lvDictionary.SelectedIndex <> -1) Then
            btnDelete.Enabled = True
        Else
            btnDelete.Enabled = False
        End If

        lblStatus.Visible = value
    End Sub
#End Region
End Class