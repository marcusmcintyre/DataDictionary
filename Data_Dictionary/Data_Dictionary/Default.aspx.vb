Imports System.Data.OleDb
Imports System.Data.Sql
Imports System
Imports System.Drawing.Color

Public Class _Default
    Inherits System.Web.UI.Page

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not (IsPostBack()) Then
            ViewState("ps") = 0
            EnableDisableForm("False")
            populateTableDropDown("")
            'filterDataDictionaryByTable()
            populateDataDropDown("")
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
        'filterDataDictionaryByTable()
    End Sub


    Private Sub ddlColumnType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlColumnType.SelectedIndexChanged
        checkDataDropDown()
    End Sub
#End Region
#Region "GridView Events"
    Protected Sub lvDictionary_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gvDictionary.SelectedIndexChanged
        If (gvDictionary.SelectedIndex <> -1) Then
            updateStatus(0)
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
        Page.ClientScript.RegisterStartupScript(GetType(Page), "script", "success();", True)
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        updateStatus(1)
        clearForm()
        tbTableName.Text = ddlTable.SelectedValue
        ddlColumnType.SelectedIndex = -1
        gvDictionary.SelectedIndex = -1
        EnableDisableForm("True")
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If (gvDictionary.SelectedIndex = -1) Then
            EnableDisableForm("False")
        End If
        clearForm()
        filterDataDictionaryByTable()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        modifyEntry("deleteCommand")
        gvDictionary.SelectedIndex = -1
        EnableDisableForm("False")
        clearForm()
        filterDataDictionaryByTable()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ddlTable.SelectedIndex = -1
        tbSearch.Text = ""
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim sel As String = "SELECT * FROM [DATA_DICTIONARY]"
        Dim clause As String = ""
        Dim whereClause As String = " WHERE "
        Dim col As String = ""
        Dim tbl As String = ""
        Dim parm As String = ""

        For i = 0 To (SQLServer.SelectParameters.Count - 1)
            Dim oldParm = SQLServer.SelectParameters.Item(0)
            SQLServer.SelectParameters.Remove(oldParm)
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
            SQLServer.SelectParameters.Add(table)
            tbl = "[TABLE_NAME] LIKE @table"
            clause += tbl
        End If

        If Not tbSearch.Text.Equals("") Then
            Dim column = New Parameter
            column.Name = "search"
            column.DefaultValue = tbSearch.Text
            SQLServer.SelectParameters.Add(column)
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

        SQLServer.SelectCommand = sel + whereClause + ";"
        gvDictionary.DataBind()
    End Sub
#End Region
#Region "Helper Subs/Functions"
    Private Sub filterDataDictionaryByTable()
        Dim parm = New ControlParameter
        parm.Name = "Table"
        parm.ControlID = "ddlTable"
        parm.PropertyName = "SelectedValue"
        parm.DefaultValue = ddlTable.SelectedValue

        If (SQLServer.SelectParameters.Count > 0) Then
            Dim oldParm = SQLServer.SelectParameters.Item(0)
            SQLServer.SelectParameters.Remove(oldParm)
        End If

        SQLServer.SelectParameters.Add(parm)

        If (ddlTable.SelectedIndex <> 0) Then
            SQLServer.SelectCommand = ConfigurationManager.AppSettings("selectCommand")
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
            SQLServer.DeleteParameters.Add("ID", id)
            SQLServer.DeleteCommand = ConfigurationManager.AppSettings(comm)
            SQLServer.Delete()
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
            params.Add("NULLABILITY", chkNullable.Checked)
            params.Add("KEY", tbKey.Text)
            params.Add("DESCRIPTION", tbDescription.Text)

            ' UPDATE COMMAND
            If gvDictionary.SelectedIndex <> -1 Then
                For i = 0 To params.Count - 1
                    SQLServer.UpdateParameters.Add(params.Item(i))
                Next
                SQLServer.UpdateParameters.Add("ID", id)
                SQLServer.UpdateCommand = ConfigurationManager.AppSettings(comm)
                SQLServer.Update()
            Else
                For i = 0 To params.Count - 1
                    SQLServer.InsertParameters.Add(params.Item(i))
                Next
                SQLServer.InsertCommand = ConfigurationManager.AppSettings(comm)
                SQLServer.Insert()
            End If

            End If

            populateTableDropDown(tbTableName.Text)
            filterDataDictionaryByTable()

        If (gvDictionary.SelectedIndex = -1) Then
            Dim oldCmd As String = SQLServer.SelectCommand
            SQLServer.SelectCommand = "SELECT [ID] FROM [DATA_DICTIONARY] ORDER BY [ID] DESC"
            Dim reader As DataView = SQLServer.Select(DataSourceSelectArguments.Empty)
            If reader.Count > 0 Then
                id = reader(0).ToString()
            End If
            SQLServer.SelectCommand = oldCmd
        End If

            selectlv(gvDictionary, id)
    End Sub

    Private Sub selectlv(ByRef gv As GridView, ByVal str As String)
        Dim x, i As Integer
        Dim lblID As String
        Dim intRowFound, intPageFound As Integer
        'EACH PAGE
        For i = 0 To gv.PageCount
            gv.PageIndex = i
            gv.DataBind()
            If gv.Rows.Count > 0 Then
                'EACH ROW
                For x = 0 To gv.Rows.Count - 1
                    lblID = gv.Rows(x).Cells(1).Text
                    'FOUND IT
                    If lblID.Equals(str) Then
                        intRowFound = x
                        intPageFound = gv.PageIndex
                    End If
                Next
            End If
        Next

        gv.SelectedIndex = intRowFound
        gv.DataBind()

    End Sub

    Private Sub updateStatus(ByVal mode As Integer)
        If (mode = 1) Then
            lblStatus.Text = "Adding New Entry"
            lblCurrID.Text = ""
        Else
            Dim id As String = gvDictionary.SelectedRow.Cells(1).Text
            Dim table As String = gvDictionary.SelectedRow.Cells(2).Text
            Dim column As String = gvDictionary.SelectedRow.Cells(3).Text
            lblStatus.Text = "Working on: " & table & ", " & column
            lblCurrID.Text = "ID: " & id
        End If
    End Sub

    Private Sub fillSelectedDetails()
        clearForm()
        Dim lbl As String
        Dim chk As CheckBox

        lbl = gvDictionary.SelectedRow.Cells(2).Text
        tbTableName.Text = lbl

        lbl = gvDictionary.SelectedRow.Cells(3).Text
        tbColumnName.Text = lbl

        lbl = gvDictionary.SelectedRow.Cells(4).Text
        ddlColumnType.SelectedIndex = populateDataDropDown(lbl)

        lbl = gvDictionary.SelectedRow.Cells(5).Text
        tbColumnSize.Text = lbl

        lbl = gvDictionary.SelectedRow.Cells(6).Text
        tbPrecision.Text = lbl

        lbl = gvDictionary.SelectedRow.Cells(7).Text
        tbScale.Text = lbl

        lbl = gvDictionary.SelectedRow.Cells(8).Text
        chkNullable.Checked = lbl

        lbl = gvDictionary.SelectedRow.Cells(9).Text
        tbKey.Text = lbl

        lbl = gvDictionary.SelectedRow.Cells(10).Text
        If Not (lbl.Equals("&nbsp;")) Then
            tbDescription.Text = lbl
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
        Dim oldCmd As String = SQLServer.SelectCommand

        arrTable.Add("")
        SQLServer.SelectCommand = "SELECT DISTINCT [TABLE_NAME] FROM [DATA_DICTIONARY] ORDER BY [TABLE_NAME]"

        Dim read As DataView = SQLServer.Select(DataSourceSelectArguments.Empty)
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
        SQLServer.SelectCommand = oldCmd
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