Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class frmRoles

    'Dim sqlConn As String = "DRIVER={MySQL ODBC 3.51 Driver};" _
    '                          & "SERVER=localhost;" _
    '                          & "UID=root;" _
    '                          & "PWD=password;" _
    '                          & "DATABASE=wg_tms_db_90;"

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim theCode As String = txtRoleCode.Text.ToString.ToUpper.Trim

        If txtRoleCode.Text = "" Then
            validation1.Visible = True
        Else
            validation1.Visible = False

            Try
                connect = New OdbcConnection(sqlConn)

                Dim inscommand As OdbcCommand
                Dim insadapter As OdbcDataAdapter

                Dim dC As String = txtRoleCode.Text.ToString.ToUpper.Trim
                Dim dN1 As String = txtRoleDesc.Text.ToString.ToUpper.Trim

                connect.Open()
                'INSERT THE NEW ROLE
                Dim insCust As String = "insert into s_roles(role_code, role_name) values(?,?)"
                inscommand = New OdbcCommand(insCust, connect)
                insadapter = New OdbcDataAdapter()

                inscommand.Parameters.Add("@role_code", OdbcType.VarChar, 10).Value = dC.ToString
                inscommand.Parameters.Add("@role_name", OdbcType.VarChar, 100).Value = dN1.ToString
                insadapter.InsertCommand = inscommand
                insadapter.InsertCommand.ExecuteNonQuery()

                'SELECT THE MENU
                Dim selApp As String = "select menu_code from s_application order by ID;"

                Dim selCommand As OdbcCommand
                Dim selAdpt As OdbcDataAdapter
                Dim selDS As DataSet

                selCommand = New OdbcCommand(selApp, connect)
                selAdpt = New OdbcDataAdapter()

                selDS = New DataSet()
                selDS.Tables.Clear()
                selAdpt = New OdbcDataAdapter()
                selAdpt.SelectCommand = selCommand
                selAdpt.Fill(selDS, "menu")

                'Dim dtRetrievedData As DataTable = myDataSet.Tables(0)
                Dim dtRetrievedData As DataTable = selDS.Tables("menu")
                'Dim dtData As New DataTable
                'Dim dtDataRows As DataRow
                Dim dtData As New DataTable
                Dim dtDataRows As DataRow

                'dtData.TableName = "roles"
                'dtData.Columns.Add("Username")
                'dtData.Columns.Add("Password")
                'dtData.Columns.Add("Role")
                dtData.TableName = "menu"
                dtData.Columns.Add("menu_code")

                'For Each dtDataRows In dtRetrievedData.Rows

                '    Dim theName As String = dtDataRows("theUName").ToString.Trim()
                '    Dim thePW As String = dtDataRows("theUPass").ToString.Trim()

                '    For Each i As Char In thePW
                '        Select Case i
                '            Case ""
                '            Case Else
                '                thePW = thePW.Replace(i, "*")
                '        End Select
                '    Next

                '    Dim roLe As String = dtDataRows("theRole").ToString.Trim()

                '    dtData.Rows.Add(New Object() {theName.ToString.Trim(), thePW.ToString.Trim(), roLe.ToString.Trim()})
                'Next
                For Each dtDataRows In dtRetrievedData.Rows

                    Dim menu_code As String = dtDataRows("menu_code").ToString.Trim()

                    dtData.Rows.Add(New Object() {menu_code.ToString.Trim()})

                    'INSERT INTO s_roles_righs
                    Dim insMenu As String = "insert into s_roles_rights(role_code,menu_code) values(?,?)"

                    Dim inscommand2 As OdbcCommand
                    Dim insadapter2 As OdbcDataAdapter

                    inscommand2 = New OdbcCommand(insMenu, connect)
                    insadapter2 = New OdbcDataAdapter()

                    inscommand2.Parameters.Add("@role_code", OdbcType.VarChar, 10).Value = theCode.ToString.ToUpper.Trim()
                    inscommand2.Parameters.Add("@menu_code", OdbcType.VarChar, 15).Value = menu_code.ToString.Trim()
                    insadapter2.InsertCommand = inscommand2
                    insadapter2.InsertCommand.ExecuteNonQuery()

                Next

                txtRoleCode.Text = ""
                txtRoleDesc.Text = ""
                txtRoleCode.Enabled = False
                txtRoleDesc.Enabled = False
                'txtRoleCode.BackColor = Color.White
                'txtRoleDesc.BackColor = Color.White

                AddRoleToolStripMenuItem.Enabled = True
                EditRoleToolStripMenuItem.Enabled = True
                DeleteRoleToolStripMenuItem.Enabled = True
                ConfigureSystemRightsToolStripMenuItem.Enabled = True

                btnSave.Visible = False
                btnCancel.Visible = False

                'btnRefresh.PerformClick()
                GridView1.Enabled = True

                'REFRESH THE GRIDVIEW
                Dim command1 As OdbcCommand
                Dim myAdapter As OdbcDataAdapter
                Dim myDataSet As DataSet

                connect = New OdbcConnection(sqlConn)

                Try

                    Dim str_roles_Show As String = "SELECT ID, role_code, role_name from s_roles;"

                    connect.Open()
                    command1 = New OdbcCommand(str_roles_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "roles")

                    Dim dtRetrievedData2 As DataTable = myDataSet.Tables(0)

                    Dim dtData2 As New DataTable
                    Dim dtDataRows2 As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData2.TableName = "roles"
                    dtData2.Columns.Add("Role Code")
                    dtData2.Columns.Add("Description")

                    For Each dtDataRows2 In dtRetrievedData2.Rows

                        Dim Code As String = dtDataRows2("role_code").ToString.Trim()
                        Dim theName As String = dtDataRows2("role_name").ToString.Trim()

                        dtData2.Rows.Add(New Object() {Code.ToString.Trim(), theName.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData2)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("Role Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Role Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Role Code").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Role Code").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Role Code").SortMode = DataGridViewColumnSortMode.NotSortable
                    GridView1.Columns.Item("Role Code").Width = 150

                    GridView1.Columns.Item("Description").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Description").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Description").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Description").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Description").SortMode = DataGridViewColumnSortMode.NotSortable
                    GridView1.Columns.Item("Description").Width = 260

                    GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()


                    connect.Close()
                Catch ex As Exception
                    connect.Close()
                End Try

                'REFRESH TEXTBOX
                Dim iRowIndex As Integer
                Dim itm2 As String
                Dim itm3 As String

                iRowIndex = GridView1.CurrentRow.Index
                itm2 = GridView1.Item(0, iRowIndex).Value
                itm3 = GridView1.Item(1, iRowIndex).Value

                txtRoleCode.Text = itm2.ToString.ToUpper.Trim
                txtRoleDesc.Text = itm3.ToString.ToUpper.Trim

                connect.Close()
            Catch ex As Exception
                MessageBox.Show("Application Role Code Exists. Use another code.", "Add Application Role", MessageBoxButtons.OK)
                txtRoleCode.Focus()
                txtRoleCode.SelectAll()
            End Try

        End If

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        txtRoleCode.Text = ""
        txtRoleDesc.Text = ""
        txtRoleCode.Enabled = False
        txtRoleDesc.Enabled = False
        'txtRoleCode.BackColor = Color.White
        'txtRoleDesc.BackColor = Color.White

        AddRoleToolStripMenuItem.Enabled = True
        EditRoleToolStripMenuItem.Enabled = True
        DeleteRoleToolStripMenuItem.Enabled = True
        ConfigureSystemRightsToolStripMenuItem.Enabled = True

        btnEditSave.Visible = False
        btnSave.Visible = False
        btnCancel.Visible = False

        GridView1.Enabled = True
        GridView1.Refresh()
        validation1.Visible = False

        Dim iRowIndex As Integer
        Dim itm2 As String
        Dim itm3 As String

        iRowIndex = GridView1.CurrentRow.Index
        itm2 = GridView1.Item(0, iRowIndex).Value
        itm3 = GridView1.Item(1, iRowIndex).Value

        txtRoleCode.Text = itm2.ToString.ToUpper.Trim
        txtRoleDesc.Text = itm3.ToString.ToUpper.Trim

    End Sub

    Private Sub GridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles GridView1.CellMouseClick
        Dim iRowIndex As Integer
        Dim itm2 As String
        Dim itm3 As String

        Try
            iRowIndex = GridView1.CurrentRow.Index
            itm2 = GridView1.Item(0, iRowIndex).Value
            itm3 = GridView1.Item(1, iRowIndex).Value

            txtRoleCode.Text = itm2.ToString.ToUpper.Trim
            txtRoleDesc.Text = itm3.ToString.ToUpper.Trim

        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmRoles_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        GridView1.DataSource = ""
        GridView1.Refresh()
    End Sub

    Private Sub frmRoles_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Try


            Dim str_roles_Show As String = "SELECT ID, role_code, role_name from s_roles;"

            connect.Open()
            command1 = New OdbcCommand(str_roles_Show, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "roles")

            Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "roles"
            dtData.Columns.Add("Role Code")
            dtData.Columns.Add("Description")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim theCode As String = dtDataRows("role_code").ToString.Trim()
                Dim theName As String = dtDataRows("role_name").ToString.Trim()

                dtData.Rows.Add(New Object() {theCode.ToString.Trim(), theName.ToString.Trim()})
            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            GridView1.DataSource = finalDataSet.Tables(0)
            'GridView1.ColumnHeadersVisible = False
            GridView1.Columns.Item("Role Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Role Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Role Code").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Role Code").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Role Code").Width = 150

            GridView1.Columns.Item("Description").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Description").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Description").HeaderCell.Style.Font = headerFont
            GridView1.Columns.Item("Description").DefaultCellStyle.Font = detailFont
            GridView1.Columns.Item("Description").FillWeight = 260
            GridView1.Columns.Item("Description").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            GridView1.Columns.Item("Role Code").SortMode = DataGridViewColumnSortMode.NotSortable
            GridView1.Columns.Item("Description").SortMode = DataGridViewColumnSortMode.NotSortable

            GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
            GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Refresh()

            connect.Close()
        Catch ex As Exception
            connect.Close()
        End Try

        Dim iRowIndex As Integer
        Dim itm2 As String
        Dim itm3 As String

        iRowIndex = GridView1.CurrentRow.Index
        itm2 = GridView1.Item(0, iRowIndex).Value
        itm3 = GridView1.Item(0, iRowIndex).Value

        txtRoleCode.Text = itm2.ToString.ToUpper.Trim
        txtRoleDesc.Text = itm3.ToString.ToUpper.Trim

        txtRoleCode.Enabled = False
        txtRoleDesc.Enabled = False

        AddRoleToolStripMenuItem.Enabled = True
        EditRoleToolStripMenuItem.Enabled = True
        DeleteRoleToolStripMenuItem.Enabled = True

        btnSave.Visible = False
        btnCancel.Visible = False

    End Sub

    Private Sub btnEditSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditSave.Click

        Dim updtcommand As OdbcCommand
        Dim updtAdapter As OdbcDataAdapter

        connect = New OdbcConnection(sqlConn)

        Try

            connect = New OdbcConnection(sqlConn)
            connect.Open()

            Dim dC As String = txtRoleCode.Text.ToString.Trim
            Dim dN1 As String = txtRoleDesc.Text.ToString.Trim

            Dim updtRole As String = "update s_roles set role_name = ? where role_code = '" & dC.ToString.ToUpper.Trim & "';"

            updtcommand = New OdbcCommand(updtRole, connect)
            updtAdapter = New OdbcDataAdapter()

            updtcommand.Parameters.AddWithValue("@role_name", dN1.ToString.ToUpper.Trim)
            updtAdapter.UpdateCommand = updtcommand
            updtAdapter.UpdateCommand.ExecuteNonQuery()

            txtRoleCode.Enabled = False
            txtRoleDesc.Enabled = False
            'txtRoleCode.BackColor = Color.White
            'txtRoleDesc.BackColor = Color.White

            AddRoleToolStripMenuItem.Enabled = True
            EditRoleToolStripMenuItem.Enabled = True
            DeleteRoleToolStripMenuItem.Enabled = True
            ConfigureSystemRightsToolStripMenuItem.Enabled = True

            btnEditSave.Visible = False
            btnCancel.Visible = False

            'btnRefresh.PerformClick()
            GridView1.Enabled = True

            'REFRESH THE GRIDVIEW
            Dim command1 As OdbcCommand
            Dim myAdapter As OdbcDataAdapter
            Dim myDataSet As DataSet

            connect = New OdbcConnection(sqlConn)

            Try

                Dim str_roles_Show As String = "SELECT ID, role_code, role_name from s_roles;"

                connect.Open()
                command1 = New OdbcCommand(str_roles_Show, connect)

                myDataSet = New DataSet()
                myDataSet.Tables.Clear()
                myAdapter = New OdbcDataAdapter()
                myAdapter.SelectCommand = command1
                myAdapter.Fill(myDataSet, "roles")

                Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                Dim dtData As New DataTable
                Dim dtDataRows As DataRow

                Dim chkboxcol As New DataGridViewCheckBoxColumn

                dtData.TableName = "roles"
                dtData.Columns.Add("Role Code")
                dtData.Columns.Add("Description")

                For Each dtDataRows In dtRetrievedData.Rows

                    Dim Code As String = dtDataRows("role_code").ToString.Trim()
                    Dim theName As String = dtDataRows("role_name").ToString.Trim()

                    dtData.Rows.Add(New Object() {Code.ToString.Trim(), theName.ToString.Trim()})
                Next

                Dim finalDataSet As New DataSet
                finalDataSet.Tables.Add(dtData)

                GridView1.DataSource = finalDataSet.Tables(0)
                'GridView1.ColumnHeadersVisible = False
                GridView1.Columns.Item("Role Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Role Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Role Code").HeaderCell.Style.Font = headerFont
                GridView1.Columns.Item("Role Code").DefaultCellStyle.Font = detailFont
                GridView1.Columns.Item("Role Code").Width = 150

                GridView1.Columns.Item("Description").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Description").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Description").HeaderCell.Style.Font = headerFont
                GridView1.Columns.Item("Description").DefaultCellStyle.Font = detailFont
                GridView1.Columns.Item("Description").Width = 260

                GridView1.Columns.Item("Role Code").SortMode = DataGridViewColumnSortMode.NotSortable
                GridView1.Columns.Item("Description").SortMode = DataGridViewColumnSortMode.NotSortable

                GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Refresh()


                connect.Close()
            Catch ex As Exception
                connect.Close()
            End Try

            'REFRESH TEXTBOX
            Dim iRowIndex As Integer
            Dim itm2 As String
            Dim itm3 As String

            iRowIndex = GridView1.CurrentRow.Index
            itm2 = GridView1.Item(0, iRowIndex).Value
            itm3 = GridView1.Item(1, iRowIndex).Value

            txtRoleCode.Text = itm2.ToString.ToUpper.Trim
            txtRoleDesc.Text = itm3.ToString.ToUpper.Trim

            connect.Close()

        Catch ex As Exception
            'MessageBox.Show("Error in saving file", "Edit Debtor/Customer", MessageBoxButtons.OK)
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()

    End Sub

    Private Sub AddRoleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddRoleToolStripMenuItem.Click
        txtRoleCode.Text = ""
        txtRoleDesc.Text = ""
        'txtRoleCode.BackColor = Color.AntiqueWhite
        'txtRoleDesc.BackColor = Color.AntiqueWhite
        txtRoleCode.Enabled = True
        txtRoleDesc.Enabled = True

        txtRoleCode.Focus()

        AddRoleToolStripMenuItem.Enabled = False
        EditRoleToolStripMenuItem.Enabled = False
        DeleteRoleToolStripMenuItem.Enabled = False
        ConfigureSystemRightsToolStripMenuItem.Enabled = False

        btnSave.Visible = True
        btnCancel.Visible = True

        GridView1.Enabled = False
    End Sub

    Private Sub EditRoleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditRoleToolStripMenuItem.Click
        Dim iRowIndex As Integer
        Dim itm2 As String
        Dim itm3 As String

        Try
            iRowIndex = GridView1.CurrentRow.Index
            itm2 = GridView1.Item(0, iRowIndex).Value
            itm3 = GridView1.Item(1, iRowIndex).Value

            txtRoleCode.Text = itm2.ToString.ToUpper.Trim
            'txtRoleCode.BackColor = Color.White
            'txtRoleDesc.BackColor = Color.AntiqueWhite

            txtRoleDesc.Text = itm3.ToString.ToUpper.Trim

            txtRoleCode.Focus()

            AddRoleToolStripMenuItem.Enabled = False
            EditRoleToolStripMenuItem.Enabled = False
            DeleteRoleToolStripMenuItem.Enabled = False
            ConfigureSystemRightsToolStripMenuItem.Enabled = False

            btnEditSave.Visible = True
            btnCancel.Visible = True
            btnSave.Visible = False

            txtRoleCode.Enabled = False
            txtRoleDesc.Enabled = True

            GridView1.Enabled = False
        Catch ex As Exception
            MessageBox.Show("Select a record first", "Edit Application Role", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub DeleteRoleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRoleToolStripMenuItem.Click
        Try

            Dim result As Integer = MessageBox.Show("Confirm Delete Role?", "Delete Application Role", MessageBoxButtons.OKCancel)
            If result = DialogResult.OK Then

                Try

                    Dim delcommand As OdbcCommand
                    Dim delcommand2 As OdbcCommand
                    Dim deladapter As OdbcDataAdapter
                    Dim deladapter2 As OdbcDataAdapter

                    connect = New OdbcConnection(sqlConn)
                    connect.Open()

                    Dim del_cust As String = "delete from s_roles where role_code = '" & txtRoleCode.Text.ToString.ToUpper.Trim & "';"
                    Dim del_rights As String = "delete from s_roles_rights where role_code = '" & txtRoleCode.Text.ToString.ToUpper.Trim & "';"

                    delcommand = New OdbcCommand(del_cust, connect)
                    delcommand2 = New OdbcCommand(del_rights, connect)
                    deladapter = New OdbcDataAdapter()
                    deladapter2 = New OdbcDataAdapter()


                    deladapter.DeleteCommand = delcommand
                    deladapter.DeleteCommand.ExecuteNonQuery()
                    deladapter2.DeleteCommand = delcommand2
                    deladapter2.DeleteCommand.ExecuteNonQuery()

                    MessageBox.Show("Application Role Deleted", "Delete Application Role", MessageBoxButtons.OK)
                Catch ex As Exception
                    MessageBox.Show("There are existing users in this Role. Delete not allowed." & vbCrLf & vbCrLf & "Error Message : " & ex.ToString, "Error Deleting Role", MessageBoxButtons.OK)
                End Try

                'REFRESH THE GRIDVIEW
                Dim command1 As OdbcCommand
                Dim myAdapter As OdbcDataAdapter
                Dim myDataSet As DataSet

                connect = New OdbcConnection(sqlConn)

                Try

                    Dim str_roles_Show As String = "SELECT ID, role_code, role_name from s_roles;"

                    connect.Open()
                    command1 = New OdbcCommand(str_roles_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "roles")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData.TableName = "roles"
                    dtData.Columns.Add("Role Code")
                    dtData.Columns.Add("Description")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim Code As String = dtDataRows("role_code").ToString.Trim()
                        Dim theName As String = dtDataRows("role_name").ToString.Trim()

                        dtData.Rows.Add(New Object() {Code.ToString.Trim(), theName.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    GridView1.Columns.Item("Role Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Role Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Role Code").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Role Code").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Role Code").Width = 150

                    GridView1.Columns.Item("Description").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Description").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Description").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Description").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Description").Width = 260

                    GridView1.Columns.Item("Role Code").SortMode = DataGridViewColumnSortMode.NotSortable
                    GridView1.Columns.Item("Description").SortMode = DataGridViewColumnSortMode.NotSortable

                    GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()

                    connect.Close()
                Catch ex As Exception
                    connect.Close()
                End Try

                'REFRESH TEXTBOX
                Dim iRowIndex As Integer
                Dim itm2 As String
                Dim itm3 As String

                iRowIndex = GridView1.CurrentRow.Index
                itm2 = GridView1.Item(0, iRowIndex).Value
                itm3 = GridView1.Item(1, iRowIndex).Value

                txtRoleCode.Text = itm2.ToString.ToUpper.Trim
                txtRoleDesc.Text = itm3.ToString.ToUpper.Trim

                connect.Close()

            ElseIf result = DialogResult.Cancel Then

            End If

        Catch ex As Exception
            MessageBox.Show("Select a record first", "Edit Application Role", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub ConfigureSystemRightsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfigureSystemRightsToolStripMenuItem.Click
        'frmRoleRights.Close()
        Dim dCode As String = txtRoleCode.Text
        Dim dDesc As String = txtRoleDesc.Text
        frmRoleRights.StartPosition = FormStartPosition.CenterScreen
        frmRoleRights.txtRoleCode.Text = dCode.ToString.ToUpper.Trim()
        frmRoleRights.txtRoleDesc.Text = dDesc.ToString.ToUpper.Trim()
        frmRoleRights.DataGridView1.ReadOnly = False
        frmRoleRights.ShowDialog()
    End Sub
End Class