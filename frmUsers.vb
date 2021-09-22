Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl
Imports System.Text.RegularExpressions

Public Class frmUsers

    'Dim sqlConn As String = "DRIVER={MySQL ODBC 3.51 Driver};" _
    '                          & "SERVER=localhost;" _
    '                          & "UID=root;" _
    '                          & "PWD=password;" _
    '                          & "DATABASE=wg_tms_db_90;"

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Arial", 8, FontStyle.Bold)
    Dim detailFont As New Font("Calibri", 6, FontStyle.Regular)

    Private Sub GridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles GridView1.CellMouseClick

        Dim iRowIndex As Integer
        Dim itm2 As String
        Dim itm3 As String
        Dim itm4 As String

        Try
            iRowIndex = GridView1.CurrentRow.Index
            itm2 = GridView1.Item(0, iRowIndex).Value
            itm3 = GridView1.Item(1, iRowIndex).Value
            itm4 = GridView1.Item(2, iRowIndex).Value

            txtRoleCode.Text = itm2.ToString.ToUpper.Trim
            txtRoleDesc.Text = itm3.ToString.Trim
            cmbRoles.Text = itm4.ToString.ToUpper.Trim
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub frmUsers_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        GridView1.DataSource = ""
        GridView1.Refresh()
    End Sub

    Private Sub frmUsers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Me.BackColor = Color.White
        'Me.ForeColor = Color.Black

        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Try

            Dim str_roles_Show As String = "SELECT S.ID AS theID, S.username AS theUName, S.password AS theUPass, A.role_code AS theRole, A.role_name AS theRName from s_users S, s_roles A where S.role_code = A.role_code;"

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
            dtData.Columns.Add("Username")
            dtData.Columns.Add("Password")
            dtData.Columns.Add("Role")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim theName As String = dtDataRows("theUName").ToString.Trim()
                Dim thePW As String = dtDataRows("theUPass").ToString.Trim()

                For Each i As Char In thePW
                    Select Case i
                        Case ""
                        Case Else
                            thePW = thePW.Replace(i, "*")
                    End Select
                Next

                Dim roLe As String = dtDataRows("theRole").ToString.Trim()

                dtData.Rows.Add(New Object() {theName.ToString.Trim(), thePW.ToString.Trim(), roLe.ToString.Trim()})
                'dtData.Rows.Add(New Object() {theName.ToString.Trim(), roLe.ToString.Trim()})
            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            GridView1.DataSource = finalDataSet.Tables(0)
            'GridView1.ColumnHeadersVisible = False
            GridView1.Columns.Item("Username").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Username").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Username").Width = 150

            GridView1.Columns.Item("Password").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Password").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Password").Width = 150

            GridView1.Columns.Item("Role").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Role").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            GridView1.Columns.Item("Role").FillWeight = 150
            GridView1.Columns.Item("Role").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            GridView1.Columns.Item("Username").SortMode = DataGridViewColumnSortMode.NotSortable
            GridView1.Columns.Item("Password").SortMode = DataGridViewColumnSortMode.NotSortable
            GridView1.Columns.Item("Role").SortMode = DataGridViewColumnSortMode.NotSortable

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
        Dim itm4 As String

        iRowIndex = GridView1.CurrentRow.Index
        itm2 = GridView1.Item(0, iRowIndex).Value
        itm3 = GridView1.Item(1, iRowIndex).Value
        itm4 = GridView1.Item(2, iRowIndex).Value

        txtRoleCode.Text = itm2.ToString.ToUpper.Trim
        txtRoleDesc.Text = itm3.ToString.Trim
        cmbRoles.Text = itm4.ToString.ToUpper.Trim

        txtRoleCode.Enabled = False
        txtRoleDesc.Enabled = False
        cmbRoles.Enabled = False

        AddToolStripMenuItem.Enabled = True
        EditToolStripMenuItem.Enabled = True
        DeleteToolStripMenuItem.Enabled = True
        ModifyUserDetailsToolStripMenuItem.Enabled = True

        btnSave.Visible = False
        btnCancel.Visible = False
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        txtRoleCode.Text = ""
        txtRoleDesc.Text = ""
        cmbRoles.Text = ""
        txtRoleCode.Enabled = False
        txtRoleDesc.Enabled = False
        cmbRoles.Enabled = False
        'txtRoleCode.BackColor = Color.White
        'txtRoleDesc.BackColor = Color.White
        'cmbRoles.BackColor = Color.White


        AddToolStripMenuItem.Enabled = True
        EditToolStripMenuItem.Enabled = True
        DeleteToolStripMenuItem.Enabled = True
        ModifyUserDetailsToolStripMenuItem.Enabled = True

        btnEditSave.Visible = False
        btnSave.Visible = False
        btnCancel.Visible = False

        GridView1.Enabled = True
        btnRefresh.Enabled = True
        GridView1.Refresh()
        validation1.Visible = False


        Dim iRowIndex As Integer
        Dim itm2 As String
        Dim itm3 As String
        Dim itm4 As String

        iRowIndex = GridView1.CurrentRow.Index
        itm2 = GridView1.Item(0, iRowIndex).Value
        itm3 = GridView1.Item(1, iRowIndex).Value
        itm4 = GridView1.Item(2, iRowIndex).Value

        txtRoleCode.Text = itm2.ToString.ToUpper.Trim
        txtRoleDesc.Text = itm3.ToString.Trim
        cmbRoles.Text = itm4.ToString.ToUpper.Trim

    End Sub

    Private Sub btnEditSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditSave.Click
        Dim updtcommand As OdbcCommand
        Dim updtAdapter As OdbcDataAdapter

        connect = New OdbcConnection(sqlConn)

        Try

            connect = New OdbcConnection(sqlConn)
            connect.Open()

            Dim dC As String = txtRoleCode.Text.ToString.ToUpper.Trim
            Dim dN1 As String = txtRoleDesc.Text.ToString.Trim
            Dim dN2 As String = cmbRoles.SelectedValue.ToString.ToUpper.Trim

            Dim updtRole As String = "update s_users set password = ?, role_code = ? where username = '" & dC.ToString.ToUpper.Trim & "';"

            updtcommand = New OdbcCommand(updtRole, connect)
            updtAdapter = New OdbcDataAdapter()

            updtcommand.Parameters.AddWithValue("@password", dN1.ToString.Trim)
            updtcommand.Parameters.AddWithValue("@role_code", dN2.ToString.ToUpper.Trim)
            updtAdapter.UpdateCommand = updtcommand
            updtAdapter.UpdateCommand.ExecuteNonQuery()

            txtRoleCode.Enabled = False
            txtRoleDesc.Enabled = False
            cmbRoles.Enabled = False

            'txtRoleCode.BackColor = Color.White
            'txtRoleDesc.BackColor = Color.White
            'cmbRoles.BackColor = Color.White

            AddToolStripMenuItem.Enabled = True
            EditToolStripMenuItem.Enabled = True
            DeleteToolStripMenuItem.Enabled = True
            ModifyUserDetailsToolStripMenuItem.Enabled = True

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

                Dim str_roles_Show As String = "SELECT S.ID AS theID, S.username AS theUName, S.password AS theUPass, A.role_code AS theRole, A.role_name AS theRName from s_users S, s_roles A where S.role_code = A.role_code;"

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
                dtData.Columns.Add("Username")
                dtData.Columns.Add("Password")
                dtData.Columns.Add("Role")

                For Each dtDataRows In dtRetrievedData.Rows

                    Dim theName As String = dtDataRows("theUName").ToString.Trim()
                    Dim thePW As String = dtDataRows("theUPass").ToString.Trim()

                    For Each i As Char In thePW
                        Select Case i
                            Case ""
                            Case Else
                                thePW = thePW.Replace(i, "*")
                        End Select
                    Next

                    Dim roLe As String = dtDataRows("theRole").ToString.Trim()

                    dtData.Rows.Add(New Object() {theName.ToString.Trim(), thePW.ToString.Trim(), roLe.ToString.Trim()})
                Next

                Dim finalDataSet As New DataSet
                finalDataSet.Tables.Add(dtData)

                GridView1.DataSource = finalDataSet.Tables(0)
                'GridView1.ColumnHeadersVisible = False
                GridView1.Columns.Item("Username").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Username").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Username").Width = 150

                GridView1.Columns.Item("Password").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Password").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Password").Width = 150

                GridView1.Columns.Item("Role").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Role").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                GridView1.Columns.Item("Role").Width = 150

                GridView1.Columns.Item("Username").SortMode = DataGridViewColumnSortMode.NotSortable
                GridView1.Columns.Item("Password").SortMode = DataGridViewColumnSortMode.NotSortable
                GridView1.Columns.Item("Role").SortMode = DataGridViewColumnSortMode.NotSortable

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
            Dim itm4 As String

            iRowIndex = GridView1.CurrentRow.Index
            itm2 = GridView1.Item(0, iRowIndex).Value
            itm3 = GridView1.Item(1, iRowIndex).Value
            itm4 = GridView1.Item(2, iRowIndex).Value

            txtRoleCode.Text = itm2.ToString.ToUpper.Trim
            txtRoleDesc.Text = itm3.ToString.Trim
            cmbRoles.Text = itm4.ToString.ToUpper.Trim

            connect.Close()

        Catch ex As Exception
            'MessageBox.Show("Error in saving file", "Edit Debtor/Customer", MessageBoxButtons.OK)
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim theCode As String = txtRoleCode.Text.ToString.ToUpper.Trim

        If txtRoleCode.Text = "" Then
            validation1.Visible = True
        Else
            validation1.Visible = False
            If cmbRoles.SelectedIndex = -1 Then
                validation2.Visible = True
            Else
                validation2.Visible = False

                Try
                    connect = New OdbcConnection(sqlConn)

                    Dim inscommand As OdbcCommand
                    Dim insadapter As OdbcDataAdapter

                    Dim inscommand1 As OdbcCommand
                    Dim insadapter1 As OdbcDataAdapter


                    Dim dC As String = txtRoleCode.Text.ToString.ToUpper.Trim
                    Dim dN1 As String = txtRoleDesc.Text.ToString.Trim
                    Dim dN2 As String = cmbRoles.SelectedValue.ToString.Trim
                    Dim wage As Double = 0.0

                    connect.Open()
                    Dim insCust As String = "insert into s_users(username, password, role_code) values(?,?,?)"

                    inscommand = New OdbcCommand(insCust, connect)
                    insadapter = New OdbcDataAdapter()

                    inscommand.Parameters.Add("@username", OdbcType.VarChar, 15).Value = dC.ToString
                    inscommand.Parameters.Add("@password", OdbcType.VarChar, 20).Value = dN1.ToString
                    inscommand.Parameters.Add("@role_code", OdbcType.VarChar, 10).Value = dN2.ToString
                    insadapter.InsertCommand = inscommand
                    insadapter.InsertCommand.ExecuteNonQuery()

                    Dim insCust1 As String = "insert into m_employee(username, emp_wage) values(?,?)"

                    inscommand1 = New OdbcCommand(insCust1, connect)
                    insadapter1 = New OdbcDataAdapter()

                    inscommand1.Parameters.Add("@username", OdbcType.VarChar, 15).Value = dC.ToString
                    inscommand1.Parameters.Add("@emp_wage", OdbcType.Double).Value = wage.ToString
                    insadapter1.InsertCommand = inscommand1
                    insadapter1.InsertCommand.ExecuteNonQuery()

                    txtRoleCode.Text = ""
                    txtRoleDesc.Text = ""
                    txtRoleCode.Enabled = False
                    txtRoleDesc.Enabled = False
                    cmbRoles.Enabled = False

                    'txtRoleCode.BackColor = Color.White
                    'txtRoleDesc.BackColor = Color.White
                    'cmbRoles.BackColor = Color.White

                    AddToolStripMenuItem.Enabled = True
                    EditToolStripMenuItem.Enabled = True
                    DeleteToolStripMenuItem.Enabled = True
                    ModifyUserDetailsToolStripMenuItem.Enabled = True

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

                        Dim str_roles_Show As String = "SELECT S.ID AS theID, S.username AS theUName, S.password AS theUPass, A.role_code AS theRole, A.role_name AS theRName from s_users S, s_roles A where S.role_code = A.role_code;"

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
                        dtData.Columns.Add("Username")
                        dtData.Columns.Add("Password")
                        dtData.Columns.Add("Role")

                        For Each dtDataRows In dtRetrievedData.Rows

                            Dim theName As String = dtDataRows("theUName").ToString.Trim()
                            Dim thePW As String = dtDataRows("theUPass").ToString.Trim()

                            For Each i As Char In thePW
                                Select Case i
                                    Case ""
                                    Case Else
                                        thePW = thePW.Replace(i, "*")
                                End Select
                            Next

                            Dim roLe As String = dtDataRows("theRole").ToString.Trim()

                            dtData.Rows.Add(New Object() {theName.ToString.Trim(), thePW.ToString.Trim(), roLe.ToString.Trim()})
                        Next

                        Dim finalDataSet As New DataSet
                        finalDataSet.Tables.Add(dtData)

                        GridView1.DataSource = finalDataSet.Tables(0)
                        'GridView1.ColumnHeadersVisible = False
                        GridView1.Columns.Item("Username").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Username").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Username").Width = 150

                        GridView1.Columns.Item("Password").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Password").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Password").Width = 150

                        GridView1.Columns.Item("Role").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Role").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Role").Width = 150

                        GridView1.Columns.Item("Username").SortMode = DataGridViewColumnSortMode.NotSortable
                        GridView1.Columns.Item("Password").SortMode = DataGridViewColumnSortMode.NotSortable
                        GridView1.Columns.Item("Role").SortMode = DataGridViewColumnSortMode.NotSortable

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
                    Dim itm4 As String

                    iRowIndex = GridView1.CurrentRow.Index
                    itm2 = GridView1.Item(0, iRowIndex).Value
                    itm3 = GridView1.Item(1, iRowIndex).Value
                    itm4 = GridView1.Item(2, iRowIndex).Value

                    txtRoleCode.Text = itm2.ToString.ToUpper.Trim
                    txtRoleDesc.Text = itm3.ToString.Trim
                    cmbRoles.Text = itm4.ToString.ToUpper.Trim

                    connect.Close()

                Catch ex As Exception
                    'MessageBox.Show("Error in saving file", "Edit Debtor/Customer", MessageBoxButtons.OK)
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()
            End If

        End If
    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        'txtRoleCode.Text = ""
        'txtRoleDesc.Text = ""
        'cmbRoles.
        'txtRoleCode.Enabled = True
        'txtRoleDesc.Enabled = True

        'txtRoleCode.Focus()

        'btnAdd.Enabled = False
        'btnEdit.Enabled = False
        'btnDelete.Enabled = False

        'btnSave.Visible = True
        'btnCancel.Visible = True

        'GridView1.Enabled = False
        'btnRefresh.Enabled = False

        txtRoleCode.Text = ""
        txtRoleDesc.Text = ""
        cmbRoles.Text = ""

        'txtRoleCode.BackColor = Color.AntiqueWhite
        'txtRoleDesc.BackColor = Color.AntiqueWhite
        'cmbRoles.BackColor = Color.AntiqueWhite

        txtRoleCode.Focus()

        AddToolStripMenuItem.Enabled = False
        EditToolStripMenuItem.Enabled = False
        DeleteToolStripMenuItem.Enabled = False
        ModifyUserDetailsToolStripMenuItem.Enabled = False

        btnEditSave.Visible = False
        btnCancel.Visible = True
        btnSave.Visible = True

        txtRoleCode.Enabled = True
        txtRoleDesc.Enabled = True
        cmbRoles.Enabled = True

        btnRefresh.Enabled = False
        GridView1.Enabled = False

        Dim dCmd As OdbcCommand
        Dim dAdpt As OdbcDataAdapter
        Dim ds As DataSet

        connect = New OdbcConnection(sqlConn)

        Try
            Dim str_roles_Show As String = "SELECT role_code, role_name from s_roles;"

            connect.Open()
            dCmd = New OdbcCommand(str_roles_Show, connect)

            ds = New DataSet()
            ds.Tables.Clear()
            dAdpt = New OdbcDataAdapter()
            dAdpt.SelectCommand = dCmd
            dAdpt.Fill(ds, "roles")

            cmbRoles.DataSource = ds.Tables("roles")
            cmbRoles.DisplayMember = "role_code"
            cmbRoles.ValueMember = "role_code"
            cmbRoles.SelectedIndex = -1
            cmbRoles.DropDownStyle = ComboBoxStyle.DropDownList

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        connect.Close()
    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Dim iRowIndex As Integer
        Dim itm2 As String
        Dim itm3 As String
        Dim itm4 As String

        Try
            iRowIndex = GridView1.CurrentRow.Index
            itm2 = GridView1.Item(0, iRowIndex).Value
            itm3 = GridView1.Item(1, iRowIndex).Value
            itm4 = GridView1.Item(2, iRowIndex).Value

            txtRoleCode.Text = itm2.ToString.ToUpper.Trim
            txtRoleDesc.Text = itm3.ToString.Trim
            cmbRoles.Text = itm4.ToString.ToUpper.Trim

            txtRoleCode.Focus()

            If itm2.ToString.ToUpper.Trim = "SUPERVISOR" Then
                cmbRoles.Enabled = False
            Else
                cmbRoles.Enabled = True
            End If

            AddToolStripMenuItem.Enabled = False
            EditToolStripMenuItem.Enabled = False
            DeleteToolStripMenuItem.Enabled = False
            ModifyUserDetailsToolStripMenuItem.Enabled = False

            btnEditSave.Visible = True
            btnCancel.Visible = True
            btnSave.Visible = False

            txtRoleCode.Enabled = False
            txtRoleDesc.Enabled = True
            'txtRoleDesc.BackColor = Color.AntiqueWhite
            'cmbRoles.BackColor = Color.AntiqueWhite

            btnRefresh.Enabled = False
            GridView1.Enabled = False

            Dim dCmd As OdbcCommand
            Dim dAdpt As OdbcDataAdapter
            Dim ds As DataSet

            connect = New OdbcConnection(sqlConn)

            Try
                Dim str_roles_Show As String = "SELECT role_code, role_name from s_roles;"

                connect.Open()
                dCmd = New OdbcCommand(str_roles_Show, connect)

                ds = New DataSet()
                ds.Tables.Clear()
                dAdpt = New OdbcDataAdapter()
                dAdpt.SelectCommand = dCmd
                dAdpt.Fill(ds, "roles")

                cmbRoles.DataSource = ds.Tables("roles")
                cmbRoles.DisplayMember = "role_code"
                cmbRoles.ValueMember = "role_code"
                cmbRoles.SelectedValue = itm4.ToString
                cmbRoles.DropDownStyle = ComboBoxStyle.DropDownList

                connect.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            connect.Close()

        Catch ex As Exception
            MessageBox.Show("Select a record first", "Edit User Account", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim iRowIndex As Integer
        Dim itm2 As String

        iRowIndex = GridView1.CurrentRow.Index
        itm2 = GridView1.Item(0, iRowIndex).Value

        txtRoleCode.Text = itm2.ToString.ToUpper.Trim

        If itm2.ToString.ToUpper.Trim = "SUPERVISOR" Then
            MessageBox.Show("This User Account cannot be deleted", "Delete User Account", MessageBoxButtons.OK)
        Else

            Try

                Dim result As Integer = MessageBox.Show("Confirm Delete " & itm2.ToString.ToUpper.Trim & "?", "Delete User Account", MessageBoxButtons.OKCancel)
                If result = DialogResult.OK Then

                    Try

                        Dim delcommand As OdbcCommand
                        Dim deladapter As OdbcDataAdapter

                        connect = New OdbcConnection(sqlConn)
                        connect.Open()

                        Dim del_cust As String = "delete from s_users where username = '" & txtRoleCode.Text.ToString.ToUpper.Trim & "';"

                        delcommand = New OdbcCommand(del_cust, connect)
                        deladapter = New OdbcDataAdapter()

                        deladapter.UpdateCommand = delcommand
                        deladapter.UpdateCommand.ExecuteNonQuery()

                        MessageBox.Show("User Account Deleted", "Delete User Account", MessageBoxButtons.OK)
                    Catch ex As Exception
                        MessageBox.Show("There are existing transactions for this User Account. Delete not allowed." & vbCrLf & vbCrLf & "Error Message : " & ex.ToString, "Error Deleting User", MessageBoxButtons.OK)
                    End Try

                    Dim command1 As OdbcCommand
                    Dim myAdapter As OdbcDataAdapter
                    Dim myDataSet As DataSet

                    connect = New OdbcConnection(sqlConn)

                    Try

                        Dim str_roles_Show As String = "SELECT S.ID AS theID, S.username AS theUName, S.password AS theUPass, A.role_code AS theRole, A.role_name AS theRName from s_users S, s_roles A where S.role_code = A.role_code;"

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

                        'Dim chkboxcol As New DataGridViewCheckBoxColumn

                        dtData.TableName = "roles"
                        dtData.Columns.Add("Username")
                        dtData.Columns.Add("Password")
                        dtData.Columns.Add("Role")

                        For Each dtDataRows In dtRetrievedData.Rows

                            Dim theName As String = dtDataRows("theUName").ToString.Trim()
                            Dim thePW As String = dtDataRows("theUPass").ToString.Trim()

                            For Each i As Char In thePW
                                Select Case i
                                    Case ""
                                    Case Else
                                        thePW = thePW.Replace(i, "*")
                                End Select
                            Next

                            Dim roLe As String = dtDataRows("theRole").ToString.Trim()

                            dtData.Rows.Add(New Object() {theName.ToString.Trim(), thePW.ToString.Trim(), roLe.ToString.Trim()})
                        Next

                        Dim finalDataSet As New DataSet
                        finalDataSet.Tables.Add(dtData)

                        GridView1.DataSource = finalDataSet.Tables(0)
                        'GridView1.ColumnHeadersVisible = False
                        GridView1.Columns.Item("Username").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Username").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Username").Width = 150

                        GridView1.Columns.Item("Password").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Password").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Password").Width = 150

                        GridView1.Columns.Item("Role").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Role").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Columns.Item("Role").Width = 150

                        GridView1.Columns.Item("Username").SortMode = DataGridViewColumnSortMode.NotSortable
                        GridView1.Columns.Item("Password").SortMode = DataGridViewColumnSortMode.NotSortable
                        GridView1.Columns.Item("Role").SortMode = DataGridViewColumnSortMode.NotSortable

                        GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                        GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        GridView1.Refresh()

                        connect.Close()
                    Catch ex As Exception
                        connect.Close()
                    End Try

                    'REFRESH TEXTBOX
                    Dim iRowIndex2 As Integer = 0
                    Dim itm22 As String
                    Dim itm32 As String
                    Dim itm42 As String

                    itm22 = GridView1.Item(0, iRowIndex2).Value
                    itm32 = GridView1.Item(1, iRowIndex2).Value
                    itm42 = GridView1.Item(2, iRowIndex2).Value

                    txtRoleCode.Text = itm22.ToString.ToUpper.Trim
                    txtRoleDesc.Text = itm32.ToString.Trim
                    cmbRoles.Text = itm42.ToString.ToUpper.Trim

                    connect.Close()

                ElseIf result = DialogResult.Cancel Then

                End If

            Catch ex As Exception
                MessageBox.Show("Select a record first" & vbCrLf & vbCrLf & ex.ToString, "Edit Application Role", MessageBoxButtons.OK)
            End Try

        End If
    End Sub

    Private Sub ModifyUserDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ModifyUserDetailsToolStripMenuItem.Click
        diaEmployees.Close()

        Dim iRowIndex As Integer
        Dim itm1 As String

        Try
            iRowIndex = GridView1.CurrentRow.Index

            itm1 = GridView1.Item(0, iRowIndex).Value

            diaEmployees.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
            diaEmployees.txt1.Enabled = False
            diaEmployees.StartPosition = FormStartPosition.CenterScreen
            diaEmployees.Show()

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.ToString, "Edit Customer", MessageBoxButtons.OK)
        End Try
    End Sub
End Class