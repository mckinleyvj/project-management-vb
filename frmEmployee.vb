Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class frmEmployee

    'Dim sqlConn As String = "DRIVER={MySQL ODBC 3.51 Driver};" _
    '                          & "SERVER=192.168.100.155;" _
    '                          & "UID=root;" _
    '                          & "PWD=password;" _
    '                          & "DATABASE=wg_tms_db_90;"

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private Sub frmEmployee_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        GridView1.DataSource = ""
        GridView1.Refresh()
        txtDB.Text = ""
    End Sub

    Private Sub frmEmployee_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub frmEmployee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If cmbType.Items.Count <> 0 Then
            cmbType.SelectedIndex = 0
            lbl5.Text = "Code"
        End If

        If Checkbox1.Items.Count <> 0 Then
            Checkbox1.SelectedIndex = 1
        End If

        txtDB.Text = ""
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim searchText As String
        'Dim typeBox As String

        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        'TO LOAD CUSTOMERS IF TEXT BOX IS EMPTY

        If cmbType.Text = Nothing Then
            MessageBox.Show("Please select Search Type", "Search Employee", MessageBoxButtons.OK)
        ElseIf cmbType.Text = "Code" Then

            If txtDB.Text = "" Then

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Employee_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee ORDER BY username;"
                        str_Employee_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where deleted = 0 ORDER BY username;"
                        str_Employee_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where deleted = 1 ORDER BY username;"
                        str_Employee_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER

                    connect.Open()
                    command1 = New OdbcCommand(str_Employee_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData.TableName = "EmpTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("ID")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Telephone")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim empCode = dtDataRows("username").ToString().Trim()
                        Dim empName As String = dtDataRows("emp_name1").ToString.Trim()
                        Dim empTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
                        Dim empHP As String = dtDataRows("emp_hp").ToString.Trim()
                        Dim empDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {empDel.ToString.Trim(), empCode.ToString.Trim(), empName.ToString.Trim(), empTel1.ToString.Trim(), empHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 20

                    GridView1.Columns.Item("ID").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").Width = 100

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").FillWeight = 490
                    GridView1.Columns.Item("Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").Width = 100

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 100

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()


                    connect.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()
            End If

            'TO LOAD CUSTOMERS WITH TEXT INSIDE SEARCH BOX
            If txtDB.Text <> "" Then
                searchText = txtDB.Text.ToUpper.Trim.ToString
                'Label1.Text = searchText.ToString

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Employee_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE username LIKE '" & searchText.ToString & "%' ORDER BY username;"
                        str_Employee_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE username LIKE '" & searchText.ToString & "%' and deleted = 0 ORDER BY username;"
                        str_Employee_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE username LIKE '" & searchText.ToString & "%' and deleted = 1 ORDER BY username;"
                        str_Employee_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Employee_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_employee WHERE dbcode LIKE '%" & searchText.ToString & "%' and deleted = 0 ORDER BY username;"

                   connect.Open()
                    command1 = New OdbcCommand(str_Employee_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData.TableName = "EmpTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("ID")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Telephone")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim empCode = dtDataRows("username").ToString().Trim()
                        Dim empName As String = dtDataRows("emp_name1").ToString.Trim()
                        Dim empTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
                        Dim empHP As String = dtDataRows("emp_hp").ToString.Trim()
                        Dim empDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {empDel.ToString.Trim(), empCode.ToString.Trim(), empName.ToString.Trim(), empTel1.ToString.Trim(), empHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 20

                    GridView1.Columns.Item("ID").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").Width = 100

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").FillWeight = 490
                    GridView1.Columns.Item("Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").Width = 100

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 100

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()


                    connect.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()


            End If

        ElseIf cmbType.Text = "Name 1" Then

            If txtDB.Text = "" Then

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Employee_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee ORDER BY username;"
                        str_Employee_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where deleted = 0 ORDER BY username;"
                        str_Employee_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where deleted = 1 ORDER BY username;"
                        str_Employee_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster where deleted = 0 ORDER BY dbcode;"

                   connect.Open()
                    command1 = New OdbcCommand(str_Employee_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData.TableName = "EmpTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("ID")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Telephone")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim empCode = dtDataRows("username").ToString().Trim()
                        Dim empName As String = dtDataRows("emp_name1").ToString.Trim()
                        Dim empTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
                        Dim empHP As String = dtDataRows("emp_hp").ToString.Trim()
                        Dim empDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {empDel.ToString.Trim(), empCode.ToString.Trim(), empName.ToString.Trim(), empTel1.ToString.Trim(), empHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 20

                    GridView1.Columns.Item("ID").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").Width = 100

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").FillWeight = 490
                    GridView1.Columns.Item("Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").Width = 100

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 100

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()


                    connect.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()

            End If

            'TO LOAD CUSTOMERS WITH TEXT INSIDE SEARCH BOX
            If txtDB.Text <> "" Then
                searchText = txtDB.Text.ToUpper.Trim.ToString
                'Label1.Text = searchText.ToString

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Employee_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name1 LIKE '%" & searchText.ToString.Trim() & "%' ORDER BY username;"
                        str_Employee_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name1 LIKE '%" & searchText.ToString.Trim() & "%' and deleted = 0 ORDER BY username;"
                        str_Employee_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name1 LIKE '%" & searchText.ToString.Trim() & "%' and deleted = 1 ORDER BY username;"
                        str_Employee_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster where deleted = 0 ORDER BY dbcode;"

                    connect.Open()
                    command1 = New OdbcCommand(str_Employee_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData.TableName = "EmpTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("ID")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Telephone")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim empCode = dtDataRows("username").ToString().Trim()
                        Dim empName As String = dtDataRows("emp_name1").ToString.Trim()
                        Dim empTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
                        Dim empHP As String = dtDataRows("emp_hp").ToString.Trim()
                        Dim empDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {empDel.ToString.Trim(), empCode.ToString.Trim(), empName.ToString.Trim(), empTel1.ToString.Trim(), empHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 20

                    GridView1.Columns.Item("ID").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").Width = 100

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").FillWeight = 490
                    GridView1.Columns.Item("Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").Width = 100

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 100

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()


                    connect.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()

            End If

        ElseIf cmbType.Text = "Name 2" Then

            If txtDB.Text = "" Then

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Employee_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee ORDER BY username;"
                        str_Employee_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where deleted = 0 ORDER BY username;"
                        str_Employee_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where deleted = 1 ORDER BY username;"
                        str_Employee_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster where deleted = 0 ORDER BY dbcode;"

                    connect.Open()
                    command1 = New OdbcCommand(str_Employee_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData.TableName = "EmpTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("ID")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Telephone")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim empCode = dtDataRows("username").ToString().Trim()
                        Dim empName As String = dtDataRows("emp_name1").ToString.Trim()
                        Dim empTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
                        Dim empHP As String = dtDataRows("emp_hp").ToString.Trim()
                        Dim empDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {empDel.ToString.Trim(), empCode.ToString.Trim(), empName.ToString.Trim(), empTel1.ToString.Trim(), empHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 20

                    GridView1.Columns.Item("ID").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").Width = 100

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").FillWeight = 490
                    GridView1.Columns.Item("Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").Width = 100

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 100

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()

                    connect.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()

            End If

            'TO LOAD CUSTOMERS WITH TEXT INSIDE SEARCH BOX
            If txtDB.Text <> "" Then
                searchText = txtDB.Text.ToUpper.Trim.ToString
                'Label1.Text = searchText.ToString

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Employee_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_name2 LIKE '%" & searchText.ToString.Trim() & "%' ORDER BY username;"
                        str_Employee_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_name2 LIKE '%" & searchText.ToString.Trim() & "%' and deleted = 0 ORDER BY username;"
                        str_Employee_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT username, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_name2 LIKE '%" & searchText.ToString.Trim() & "%' and deleted = 1 ORDER BY username;"
                        str_Employee_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster where deleted = 0 ORDER BY dbcode;"

                    connect.Open()
                    command1 = New OdbcCommand(str_Employee_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData.TableName = "EmpTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("ID")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Telephone")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim empCode = dtDataRows("username").ToString().Trim()
                        Dim empName As String = dtDataRows("emp_name1").ToString.Trim()
                        Dim empTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
                        Dim empHP As String = dtDataRows("emp_hp").ToString.Trim()
                        Dim empDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {empDel.ToString.Trim(), empCode.ToString.Trim(), empName.ToString.Trim(), empTel1.ToString.Trim(), empHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 20

                    GridView1.Columns.Item("ID").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("ID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("ID").Width = 100

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").FillWeight = 490
                    GridView1.Columns.Item("Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Telephone").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Telephone").Width = 100

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 100

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()


                    connect.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
                connect.Close()

            End If
        End If
        txtDB.Focus()
        txtDB.SelectAll()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        'diaEmployees.Close()
        'diaEmployees.StartPosition = FormStartPosition.CenterParent
        ''diaEmployees.txtUser.ReadOnly = False
        ''diaEmployees.txtPass.ReadOnly = False
        ''diaEmployees.btnEditUsr.Visible = False
        'diaEmployees.ShowDialog()
        'frmUsers.Show()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim iRowIndex As Integer

        Try

            iRowIndex = GridView1.CurrentRow.Index

            Dim del As String
            del = GridView1.Item(0, iRowIndex).Value
            If del = True Then
                MessageBox.Show("Employee already Deleted", "Error", MessageBoxButtons.OK)
                Exit Try
            Else
                Dim result As Integer = MessageBox.Show("Confirm Delete Employee?", "Delete Employee", MessageBoxButtons.OKCancel)
                If result = DialogResult.OK Then
                    'Dim updtcommand As OdbcCommand
                    'Dim updtadapter As OdbcDataAdapter

                    'connect = New OdbcConnection(sqlConn)
                    'connect.Open()

                    'Dim del_emp As String = "update m_employee set deleted = 1 where username = '" & GridView1.Item(1, iRowIndex).Value & "';"

                    'updtcommand = New OdbcCommand(del_emp, connect)
                    'updtadapter = New OdbcDataAdapter()

                    'updtadapter.UpdateCommand = updtcommand
                    'updtadapter.UpdateCommand.ExecuteNonQuery()

                    'MessageBox.Show("Employee Deleted", "Delete Employee", MessageBoxButtons.OK)
                ElseIf result = DialogResult.Cancel Then
                    'Do nothing
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Please select an Employee first", "Delete Employee", MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub openEmployeeDialog()
        diaEmployees.TopLevel = False
        diaEmployees.TopMost = True
        frmMain.Panel1.Controls.Add(diaEmployees)
        diaEmployees.Location = New Point(Convert.ToInt32(frmMain.Panel1.Size.Width / 2 - diaCustomers.Width / 2),
                                   Convert.ToInt32(frmMain.Panel1.Size.Height / 2 - diaCustomers.Height / 2))
        diaEmployees.Show()
        diaEmployees.BringToFront()

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Dim iRowIndex As Integer
        Dim itm1 As String
        'Dim itm2 As String

        Dim del As Boolean

        Try
            iRowIndex = GridView1.CurrentRow.Index

            If cmbType.Text = "Code" Then
                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    MessageBox.Show("Employee already deleted", "Error", MessageBoxButtons.OK)
                    Exit Try
                Else
                    diaEmployees.Close()
                    itm1 = GridView1.Item(1, iRowIndex).Value
                    'itm2 = GridView1.Item(2, iRowIndex).Value

                    diaEmployees.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                    openEmployeeDialog()
                    'diaEmployees.txt1.Enabled = False
                    'diaEmployees.StartPosition = FormStartPosition.CenterScreen
                    'diaEmployees.Show()
                End If

            End If

            If cmbType.Text = "Name 1" Then
                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    MessageBox.Show("Employee already deleted", "Error", MessageBoxButtons.OK)
                    Exit Try
                Else
                    diaEmployees.Close()
                    itm1 = GridView1.Item(1, iRowIndex).Value
                    'itm2 = GridView1.Item(2, iRowIndex).Value

                    diaEmployees.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                    openEmployeeDialog()
                    'diaEmployees.StartPosition = FormStartPosition.CenterScreen
                    'diaEmployees.Show()
                End If

            End If

            If cmbType.Text = "Name 2" Then
                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    MessageBox.Show("Employee already deleted", "Error", MessageBoxButtons.OK)
                    Exit Try
                Else
                    diaEmployees.Close()
                    itm1 = GridView1.Item(1, iRowIndex).Value
                    'itm2 = GridView1.Item(2, iRowIndex).Value

                    diaEmployees.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                    openEmployeeDialog()
                    'diaEmployees.txt1.Enabled = False
                    'diaEmployees.StartPosition = FormStartPosition.CenterScreen
                    'diaEmployees.Show()
                End If

            End If

        Catch ex As Exception
            MessageBox.Show("Please select an Employee", "Edit Employee", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        If cmbType.SelectedIndex = 0 Then
            lbl5.Text = "Code"
        End If

        If cmbType.SelectedIndex = 1 Then
            lbl5.Text = "Name 1"
        End If
    End Sub

    Private Sub VIEWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VIEWToolStripMenuItem.Click
        viewDetail()
    End Sub

    Private Sub viewDetail()

    End Sub

    Private Sub GridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridView1.CellDoubleClick
        viewDetail()
    End Sub
End Class

'ElseIf cmbType.Text = "Name 1" Then

'            If txtDB.Text = "" Then

'                Try
''ADDED FOR NON-ACTIVE FILTER
'Dim str_Employee_Show As String

'                    If Checkbox1.SelectedIndex.ToString = 0 Then
'Dim show1 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_code <> 'SUPERVISOR' ORDER BY emp_code;"
'                        str_Employee_Show = show1.ToString
'                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
'Dim show2 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_code <> 'SUPERVISOR' and deleted = 0 ORDER BY dbcode;"
'                        str_Employee_Show = show2.ToString
'                    Else
'Dim show3 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_code <> 'SUPERVISOR' and deleted = 1 ORDER BY dbcode;"
'                        str_Employee_Show = show3.ToString
'                    End If
''END OF ADDED FOR NON-ACTIVE FILTER
''Dim str_Employee_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_employee where deleted = 0 ORDER BY dbcode;"

'                    connect.Open()
'                    command1 = New OdbcCommand(str_Employee_Show, connect)

'                    myDataSet = New DataSet()
'                    myDataSet.Tables.Clear()
'                    myAdapter = New OdbcDataAdapter()
'                    myAdapter.SelectCommand = command1
'                    myAdapter.Fill(myDataSet, "Db")

'Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

'Dim dtData As New DataTable
'Dim dtDataRows As DataRow

'                    dtData.TableName = "CustTable"
'                    dtData.Columns.Add("Code")
'                    dtData.Columns.Add("Name")
'                    dtData.Columns.Add("Name 2")
'                    dtData.Columns.Add("Phone 1")
'                    dtData.Columns.Add("Phone 2")
'                    dtData.Columns.Add("Mobile")
'                    dtData.Columns.Add("Deleted")

'                    For Each dtDataRows In dtRetrievedData.Rows

'Dim cusCode = dtDataRows("emp_code").ToString().Trim()
'Dim cusName As String = dtDataRows("emp_name1").ToString.Trim()
'Dim cusName2 As String = dtDataRows("emp_name2").ToString.Trim()
'Dim cusTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
'Dim cusTel2 As String = dtDataRows("emp_tel2").ToString.Trim()
'Dim cusHP As String = dtDataRows("emp_hp").ToString.Trim()
'Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

'                        dtData.Rows.Add(New Object() {cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim(), cusDel.ToString.Trim()})
'                    Next

'Dim finalDataSet As New DataSet
'                    finalDataSet.Tables.Add(dtData)

'                    GridView1.DataSource = finalDataSet.Tables(0)
''GridView1.ColumnHeadersVisible = False
'                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Code").Width = 55

'                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name").Width = 240

'                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name 2").Width = 240

'                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 1").Width = 75

'                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 2").Width = 75

'                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Mobile").Width = 75

'                    GridView1.Columns.Item("Deleted").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Deleted").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Deleted").Width = 60

'                    GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
'                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Refresh()


'                    connect.Close()
'                Catch ex As Exception
'                    connect.Close()
'                End Try

'            End If

''TO LOAD CUSTOMERS WITH TEXT INSIDE SEARCH BOX
'            If txtDB.Text <> "" Then
'                searchText = txtDB.Text.ToUpper.Trim.ToString
''Label1.Text = searchText.ToString

'                Try
''ADDED FOR NON-ACTIVE FILTER
'Dim str_Employee_Show As String

'                    If Checkbox1.SelectedIndex.ToString = 0 Then
'Dim show1 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name1 LIKE '%" & searchText.ToString & "%' and emp_code <> 'SUPERVISOR' ORDER BY emp_code;"
'                        str_Employee_Show = show1.ToString
'                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
'Dim show2 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name1 LIKE '%" & searchText.ToString & "%' and emp_code <> 'SUPERVISOR' and deleted = 0 ORDER BY emp_code;"
'                        str_Employee_Show = show2.ToString
'                    Else
'Dim show3 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name1 LIKE '%" & searchText.ToString & "%' and emp_code <> 'SUPERVISOR' and deleted = 1 ORDER BY emp_code;"
'                        str_Employee_Show = show3.ToString
'                    End If
''END OF ADDED FOR NON-ACTIVE FILTER
''Dim str_Employee_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_employee WHERE db_name1 LIKE '%" & searchText.ToString & "%' and deleted = 0 ORDER BY emp_code;"

'                    connect.Open()
'                    command1 = New OdbcCommand(str_Employee_Show, connect)

'                    myDataSet = New DataSet()
'                    myDataSet.Tables.Clear()
'                    myAdapter = New OdbcDataAdapter()
'                    myAdapter.SelectCommand = command1
'                    myAdapter.Fill(myDataSet, "Db")

'Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

'Dim dtData As New DataTable
'Dim dtDataRows As DataRow

'                    dtData.TableName = "CustTable"
'                    dtData.Columns.Add("Code")
'                    dtData.Columns.Add("Name")
'                    dtData.Columns.Add("Name 2")
'                    dtData.Columns.Add("Phone 1")
'                    dtData.Columns.Add("Phone 2")
'                    dtData.Columns.Add("Mobile")
'                    dtData.Columns.Add("Deleted")

'                    For Each dtDataRows In dtRetrievedData.Rows

'Dim cusCode = dtDataRows("emp_code").ToString().Trim()
'Dim cusName As String = dtDataRows("emp_name1").ToString.Trim()
'Dim cusName2 As String = dtDataRows("emp_name2").ToString.Trim()
'Dim cusTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
'Dim cusTel2 As String = dtDataRows("emp_tel2").ToString.Trim()
'Dim cusHP As String = dtDataRows("emp_hp").ToString.Trim()
'Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

'                        dtData.Rows.Add(New Object() {cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim(), cusDel.ToString.Trim()})
'                    Next

'Dim finalDataSet As New DataSet
'                    finalDataSet.Tables.Add(dtData)

'                    GridView1.DataSource = finalDataSet.Tables(0)
''GridView1.ColumnHeadersVisible = False
'                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Code").Width = 55

'                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name").Width = 240

'                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name 2").Width = 240

'                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 1").Width = 75

'                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 2").Width = 75

'                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Mobile").Width = 75

'                    GridView1.Columns.Item("Deleted").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Deleted").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Deleted").Width = 60

'                    GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
'                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Refresh()


'                    connect.Close()
'                Catch ex As Exception
'                    connect.Close()
'                End Try


'            End If

'        ElseIf cmbType.Text = "Name 2" Then

'            If txtDB.Text = "" Then

'                Try
''ADDED FOR NON-ACTIVE FILTER
'Dim str_Employee_Show As String

'                    If Checkbox1.SelectedIndex.ToString = 0 Then
'Dim show1 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_code <> 'SUPERVISOR' ORDER BY emp_code;"
'                        str_Employee_Show = show1.ToString
'                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
'Dim show2 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_code <> 'SUPERVISOR' and deleted = 0 ORDER BY emp_code;"
'                        str_Employee_Show = show2.ToString
'                    Else
'Dim show3 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee where emp_code <> 'SUPERVISOR' and deleted = 1 ORDER BY emp_code;"
'                        str_Employee_Show = show3.ToString
'                    End If
''END OF ADDED FOR NON-ACTIVE FILTER
''Dim str_Employee_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_employee where deleted = 0 ORDER BY emp_code;"

'                    connect.Open()
'                    command1 = New OdbcCommand(str_Employee_Show, connect)

'                    myDataSet = New DataSet()
'                    myDataSet.Tables.Clear()
'                    myAdapter = New OdbcDataAdapter()
'                    myAdapter.SelectCommand = command1
'                    myAdapter.Fill(myDataSet, "Db")

'Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

'Dim dtData As New DataTable
'Dim dtDataRows As DataRow

'                    dtData.TableName = "CustTable"
'                    dtData.Columns.Add("Code")
'                    dtData.Columns.Add("Name")
'                    dtData.Columns.Add("Name 2")
'                    dtData.Columns.Add("Phone 1")
'                    dtData.Columns.Add("Phone 2")
'                    dtData.Columns.Add("Mobile")
'                    dtData.Columns.Add("Deleted")

'                    For Each dtDataRows In dtRetrievedData.Rows

'Dim cusCode = dtDataRows("emp_code").ToString().Trim()
'Dim cusName As String = dtDataRows("emp_name1").ToString.Trim()
'Dim cusName2 As String = dtDataRows("emp_name2").ToString.Trim()
'Dim cusTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
'Dim cusTel2 As String = dtDataRows("emp_tel2").ToString.Trim()
'Dim cusHP As String = dtDataRows("emp_hp").ToString.Trim()
'Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

'                        dtData.Rows.Add(New Object() {cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim(), cusDel.ToString.Trim()})
'                    Next

'Dim finalDataSet As New DataSet
'                    finalDataSet.Tables.Add(dtData)

'                    GridView1.DataSource = finalDataSet.Tables(0)
''GridView1.ColumnHeadersVisible = False
'                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Code").Width = 55

'                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name").Width = 240

'                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name 2").Width = 240

'                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 1").Width = 75

'                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 2").Width = 75

'                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Mobile").Width = 75

'                    GridView1.Columns.Item("Deleted").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Deleted").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Deleted").Width = 60

'                    GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
'                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Refresh()


'                    connect.Close()
'                Catch ex As Exception
'                    connect.Close()
'                End Try

'            End If

''TO LOAD CUSTOMERS WITH TEXT INSIDE SEARCH BOX
'            If txtDB.Text <> "" Then
'                searchText = txtDB.Text.ToUpper.Trim.ToString
''Label1.Text = searchText.ToString

'                Try
''ADDED FOR NON-ACTIVE FILTER
'Dim str_Employee_Show As String

'                    If Checkbox1.SelectedIndex.ToString = 0 Then
'Dim show1 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name2 LIKE '%" & searchText.ToString & "%' and emp_code <> 'SUPERVISOR' ORDER BY emp_code;"
'                        str_Employee_Show = show1.ToString
'                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
'Dim show2 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name2 LIKE '%" & searchText.ToString & "%' and emp_code <> 'SUPERVISOR' and deleted = 0 ORDER BY emp_code;"
'                        str_Employee_Show = show2.ToString
'                    Else
'Dim show3 As String = "SELECT emp_code, emp_name1, emp_name2, emp_tel1, emp_tel2, emp_hp, emp_wage, deleted FROM m_employee WHERE emp_name2 LIKE '%" & searchText.ToString & "%' and emp_code <> 'SUPERVISOR' and deleted = 1 ORDER BY emp_code;"
'                        str_Employee_Show = show3.ToString
'                    End If
''END OF ADDED FOR NON-ACTIVE FILTER
''Dim str_Employee_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_employee WHERE db_name2 LIKE '%" & searchText.ToString & "%' and deleted = 0 ORDER BY emp_code;"

'                    connect.Open()
'                    command1 = New OdbcCommand(str_Employee_Show, connect)

'                    myDataSet = New DataSet()
'                    myDataSet.Tables.Clear()
'                    myAdapter = New OdbcDataAdapter()
'                    myAdapter.SelectCommand = command1
'                    myAdapter.Fill(myDataSet, "Db")

'Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

'Dim dtData As New DataTable
'Dim dtDataRows As DataRow

'                    dtData.TableName = "CustTable"
'                    dtData.Columns.Add("Code")
'                    dtData.Columns.Add("Name")
'                    dtData.Columns.Add("Name 2")
'                    dtData.Columns.Add("Phone 1")
'                    dtData.Columns.Add("Phone 2")
'                    dtData.Columns.Add("Mobile")
'                    dtData.Columns.Add("Deleted")

'                    For Each dtDataRows In dtRetrievedData.Rows

'Dim cusCode = dtDataRows("emp_code").ToString().Trim()
'Dim cusName As String = dtDataRows("emp_name1").ToString.Trim()
'Dim cusName2 As String = dtDataRows("emp_name2").ToString.Trim()
'Dim cusTel1 As String = dtDataRows("emp_tel1").ToString.Trim()
'Dim cusTel2 As String = dtDataRows("emp_tel2").ToString.Trim()
'Dim cusHP As String = dtDataRows("emp_hp").ToString.Trim()
'Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

'                        dtData.Rows.Add(New Object() {cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim(), cusDel.ToString.Trim()})
'                    Next

'Dim finalDataSet As New DataSet
'                    finalDataSet.Tables.Add(dtData)

'                    GridView1.DataSource = finalDataSet.Tables(0)
''GridView1.ColumnHeadersVisible = False
'                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Code").Width = 55

'                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name").Width = 240

'                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Name 2").Width = 240

'                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 1").Width = 75

'                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Phone 2").Width = 75

'                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Mobile").Width = 75

'                    GridView1.Columns.Item("Deleted").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Deleted").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Columns.Item("Deleted").Width = 60

'                    GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
'                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
'                    GridView1.Refresh()


'                    connect.Close()
'                Catch ex As Exception
'                    connect.Close()
'                End Try

'            End If