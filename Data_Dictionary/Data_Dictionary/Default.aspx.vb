Imports System.Data.OleDb
Imports System
Imports System.Drawing.Color

Public Class _Default
    Inherits System.Web.UI.Page


#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not (IsPostBack()) Then
            setDataSource()
            Session("lastIndex") = gvDictionary.SelectedIndex
            ViewState("ps") = 0
            EnableDisableForm("False")
            populateTableDropDown("")
            populateDataDropDown("")
        Else
            setDataSource()
        End If
    End Sub
#End Region
#Region "DropDown Events"
    Private Sub ddlTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTable.SelectedIndexChanged
        If (gvDictionary.SelectedIndex <> -1) Then
            gvDictionary.SelectedIndex = -1
        End If
        clearForm()
        updateStatus(1)
    End Sub


    Private Sub ddlColumnType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlColumnType.SelectedIndexChanged
        checkDataDropDown()
    End Sub
#End Region
#Region "GridView Events"
    Protected Sub lvDictionary_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gvDictionary.SelectedIndexChanged
        If (gvDictionary.SelectedIndex <> -1) Then
            updateStatus(0)
            Session("lastIndex") = gvDictionary.SelectedIndex
            Dim ct As String = gvDictionary.SelectedRow.Cells(4).Text
            ddlColumnType.SelectedIndex = populateDataDropDown(ct)
            EnableDisableForm("True")
            fillSelectedDetails()
        End If
    End Sub
#End Region
#Region "Button Events"
    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim comm As String = ""
        If gvDictionary.SelectedIndex <> -1 Then
            comm = "updateCommand"
        Else
            comm = "insertCommand"
        End If
        modifyEntry(comm)
        updateStatus(0)
        executeSearch()
        Page.ClientScript.RegisterStartupScript(GetType(Page), "script", "display('You successfully saved this entry.');", True)
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        updateStatus(1)
        clearForm()
        tbTableName.Text = ddlTable.SelectedValue
        ddlColumnType.SelectedIndex = -1
        Session("lastIndex") = gvDictionary.SelectedIndex
        gvDictionary.SelectedIndex = -1
        EnableDisableForm("True")
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        gvDictionary.SelectedIndex = Session("lastIndex")
        executeSearch()
        fillSelectedDetails()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        modifyEntry("deleteCommand")
        gvDictionary.SelectedIndex = -1
        EnableDisableForm("False")
        clearForm()
        executeSearch()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ddlTable.SelectedIndex = -1
        tbSearch.Text = ""
        clearForm()
    End Sub

    Private Sub btnApplyDescription_Click(sender As Object, e As EventArgs) Handles btnApplyDescription.Click
        If (tbDescription.Text.Length > 0) Then
            modifyEntry("applyCommand")
            updateStatus(0)
            Page.ClientScript.RegisterStartupScript(GetType(Page), "script", "display('You successfully saved this entry.');", True)
        Else
            Page.ClientScript.RegisterStartupScript(GetType(Page), "script", "display('You need to fill out a Description.');", True)
        End If
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        executeSearch()
    End Sub

    Private Sub rbSQL_CheckedChanged(sender As Object, e As EventArgs) Handles rbSQL.CheckedChanged
        setDataSource()
        Page.ClientScript.RegisterStartupScript(GetType(Page), "script", "warning();", True)
    End Sub

    Private Sub rbAccess_CheckedChanged(sender As Object, e As EventArgs) Handles rbAccess.CheckedChanged
        setDataSource()
        Page.ClientScript.RegisterStartupScript(GetType(Page), "script", "warning();", True)
    End Sub
#End Region
#Region "Helper Subs/Functions"
    Private Sub setDataSource()
        If (rbSQL.Checked) Then
            Session("gvDictionary.DataSource") = SQLServer
        ElseIf (rbAccess.Checked) Then
            Session("gvDictionary.DataSource") = AccessDataSource1
        End If
        gvDictionary.DataSource = Session("gvDictionary.DataSource")
    End Sub

    Private Sub executeSearch()
        Dim sel As String = "SELECT * FROM [DATA_DICTIONARY]"
        Dim clause As String = ""
        Dim whereClause As String = " WHERE "
        Dim orderBy As String = "  ORDER BY [KEY] DESC, [COLUMN_NAME] ASC, [COLUMN_TYPE]"
        Dim col As String = ""
        Dim tbl As String = ""
        Dim parm As String = ""

        For i = 0 To (gvDictionary.DataSource.SelectParameters.Count - 1)
            Dim oldParm = gvDictionary.DataSource.SelectParameters.Item(0)
            gvDictionary.DataSource.SelectParameters.Remove(oldParm)
        Next

        If chkExactMatch.Checked Then
            parm = "@search"
        Else
            parm = "'%' + @search + '%'"
        End If

        If ddlTable.SelectedIndex <> 0 Then
            Dim table = New Parameter
            table.Name = "table"
            table.DefaultValue = ddlTable.SelectedValue
            gvDictionary.DataSource.SelectParameters.Add(table)
            tbl = "[TABLE_NAME] LIKE @table"
            clause += tbl
        End If

        If Not tbSearch.Text.Equals("") Then
            Dim column = New Parameter
            column.Name = "search"
            column.DefaultValue = tbSearch.Text
            gvDictionary.DataSource.SelectParameters.Add(column)
            col = "[COLUMN_NAME] LIKE " + parm
            If (clause.Length > 0) Then
                clause += " AND "
            End If
            clause += col
        End If

        If (clause.Length > 0) Then
            whereClause += clause
        Else
            whereClause += "ID = -1"
        End If

        gvDictionary.DataSource.SelectCommand = sel + whereClause + orderBy + ";"
        gvDictionary.DataBind()
    End Sub

    Private Sub filterDataDictionaryByTable()
        Dim parm = New ControlParameter
        parm.Name = "Table"
        parm.ControlID = "ddlTable"
        parm.PropertyName = "SelectedValue"
        parm.DefaultValue = ddlTable.SelectedValue

        If (gvDictionary.DataSource.SelectParameters.Count > 0) Then
            Dim oldParm = gvDictionary.DataSource.SelectParameters.Item(0)
            gvDictionary.DataSource.SelectParameters.Remove(oldParm)
        End If

        gvDictionary.DataSource.SelectParameters.Add(parm)

        If (ddlTable.SelectedIndex <> 0) Then
            gvDictionary.DataSource.SelectCommand = ConfigurationManager.AppSettings("selectCommand")
            gvDictionary.DataBind()
            tbTableName.Text = ddlTable.SelectedValue
        End If
    End Sub

    Private Sub modifyEntry(ByVal comm As String)
        Dim params As New ParameterCollection
        Dim id As String = ""

        If (gvDictionary.SelectedIndex <> -1) Then
            id = gvDictionary.SelectedRow.Cells(1).Text
        End If

        ' DELETE COMMAND
        If (comm.Equals("deleteCommand")) Then
            gvDictionary.DataSource.DeleteParameters.Add("ID", id)
            gvDictionary.DataSource.DeleteCommand = ConfigurationManager.AppSettings(comm)
            gvDictionary.DataSource.Delete()
        ElseIf (comm.Equals("applyCommand")) Then
            Dim description As New Parameter("DESCRIPTION", System.Data.DbType.String, tbDescription.Text)
            gvDictionary.DataSource.UpdateParameters.Add(description)
            For Each row As GridViewRow In gvDictionary.Rows
                Dim rowID As New Parameter("ID", System.Data.DbType.String, row.Cells(1).Text)
                gvDictionary.DataSource.UpdateParameters.Add(rowID)
                gvDictionary.DataSource.UpdateCommand = ConfigurationManager.AppSettings(comm)
                gvDictionary.DataSource.Update()
                gvDictionary.DataSource.UpdateParameters.Remove(rowID)
            Next
        Else
            params.Add("TABLE_NAME", tbTableName.Text)
            params.Add("COLUMN_NAME", tbColumnName.Text)
            params.Add("COLUMN_TYPE", ddlColumnType.SelectedValue)
            params.Add("COLUMN_SIZE", tbColumnSize.Text)
            If (checkDataDropDown()) Then
                params.Add("PRECISION", tbPrecision.Text)
                params.Add("SCALE", tbScale.Text)
            Else
                params.Add("PRECISION", 0)
                params.Add("SCALE", 0)
            End If
            If (chkNullable.Checked) Then
                params.Add("NULLABILITY", 1)
            Else
                params.Add("NULLABILITY", 0)
            End If
            params.Add("KEY", tbKey.Text)
            params.Add("DESCRIPTION", tbDescription.Text)

            ' UPDATE COMMAND
            If gvDictionary.SelectedIndex <> -1 Then
                For i = 0 To params.Count - 1
                    gvDictionary.DataSource.UpdateParameters.Add(params.Item(i))
                Next
                gvDictionary.DataSource.UpdateParameters.Add("ID", id)
                gvDictionary.DataSource.UpdateCommand = ConfigurationManager.AppSettings(comm)
                gvDictionary.DataSource.Update()
            Else
                For i = 0 To params.Count - 1
                    gvDictionary.DataSource.InsertParameters.Add(params.Item(i))
                Next
                gvDictionary.DataSource.InsertCommand = ConfigurationManager.AppSettings(comm)
                gvDictionary.DataSource.Insert()
            End If
        End If

        executeSearch()

        If (gvDictionary.SelectedIndex = -1) Then
            Dim oldCmd As String = gvDictionary.DataSource.SelectCommand
            gvDictionary.DataSource.SelectCommand = "SELECT [ID] FROM [DATA_DICTIONARY]  ORDER BY [KEY] DESC, [COLUMN_NAME] ASC, [COLUMN_TYPE]"
            Dim reader As DataView = gvDictionary.DataSource.Select(DataSourceSelectArguments.Empty)
            If reader.Count > 0 Then
                id = reader(0).ToString()
            End If
            gvDictionary.DataSource.SelectCommand = oldCmd
        End If

        selectlv(gvDictionary, id)
    End Sub

    Private Sub selectlv(ByRef gv As GridView, ByVal str As String)
        Dim x As Integer
        Dim lblID As String
        Dim intRowFound As Integer = -1
        If gv.Rows.Count > 0 Then
            'EACH ROW
            For x = 0 To gv.Rows.Count - 1
                lblID = gv.Rows(x).Cells(1).Text
                'FOUND IT
                If lblID.Equals(str) Then
                    intRowFound = x
                End If
            Next
        End If

        gv.SelectedIndex = intRowFound
        gv.DataBind()
    End Sub

    Private Sub updateStatus(ByVal mode As Integer)
        If (mode = 1) Then
            lblStatus.Text = "Adding New Entry"
            lblCurrID.Text = ""
        Else
            If (gvDictionary.SelectedIndex <> -1) Then
                Dim id As String = gvDictionary.SelectedRow.Cells(1).Text
                Dim table As String = gvDictionary.SelectedRow.Cells(2).Text
                Dim column As String = gvDictionary.SelectedRow.Cells(3).Text
                lblStatus.Text = "Working on: " & table & ", " & column
                lblCurrID.Text = "ID: " & id
            End If
        End If
    End Sub

    Private Sub fillSelectedDetails()
        clearForm()
        If (gvDictionary.SelectedIndex <> -1) Then
            tbTableName.Text = HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(2).Text)
            tbColumnName.Text = HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(3).Text)
            ddlColumnType.SelectedIndex = populateDataDropDown(HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(4).Text))
            tbColumnSize.Text = HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(5).Text)
            tbPrecision.Text = HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(6).Text)
            tbScale.Text = HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(7).Text)
            chkNullable.Checked = HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(8).Text)
            tbKey.Text = HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(9).Text)
            tbDescription.Text = HttpUtility.HtmlDecode(gvDictionary.SelectedRow.Cells(10).Text)
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

    Private Sub populateTableDropDown(ByVal str As String)
        Dim arrTable As New ArrayList
        Dim sel As String = ""
        Dim oldCmd As String = gvDictionary.DataSource.SelectCommand

        arrTable.Add("")
        gvDictionary.DataSource.SelectCommand = "SELECT DISTINCT [TABLE_NAME] FROM [DATA_DICTIONARY] ORDER BY [TABLE_NAME]"

        Dim read As DataView = gvDictionary.DataSource.Select(DataSourceSelectArguments.Empty)
        For i = 0 To read.Count - 1
            If (read(i).Item(0).ToString.Equals(str)) Then
                sel = str
            End If
            arrTable.Add(read(i).Item(0).ToString)
        Next

        ddlTable.DataSource = arrTable
        ddlTable.DataBind()
        If sel.Equals("") Then
            ddlTable.SelectedIndex = 0
        Else
            ddlTable.SelectedValue = sel
        End If
        gvDictionary.DataSource.SelectCommand = oldCmd
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
        If (gvDictionary.SelectedIndex <> -1) Then
            btnDelete.Enabled = True
        Else
            btnDelete.Enabled = False
        End If

        lblStatus.Visible = value
        lblCurrID.Visible = value
    End Sub
#End Region
End Class