Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class frmCustomers

    'Dim sqlConn As String = "DRIVER={MySQL ODBC 3.51 Driver};" _
    '                          & "SERVER=localhost;" _
    '                          & "UID=root;" _
    '                          & "PWD=password;" _
    '                          & "DATABASE=wg_tms_db_90;"

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        Dim searchText As String
        'Dim typeBox As String

        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        'TO LOAD CUSTOMERS IF TEXT BOX IS EMPTY

        If cmbType.Text = Nothing Then
            MessageBox.Show("Please select Search Type", "Search Customer", MessageBoxButtons.OK)
        ElseIf cmbType.Text = "Code" Then

            If txtDB.Text = "" Then

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Customers_Show As String

                    'If Checkbox1.SelectedIndex.ToString = 0 Then
                    '    Dim show1 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster where dbcode <> '-' ORDER BY dbcode;"
                    '    str_Customers_Show = show1.ToString
                    'ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                    '    Dim show2 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster where dbcode <> '-'and deleted = 0 ORDER BY dbcode;"
                    '    str_Customers_Show = show2.ToString
                    'Else
                    '    Dim show3 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster where dbcode <> '-' and deleted = 1 ORDER BY dbcode;"
                    '    str_Customers_Show = show3.ToString
                    'End If

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT * FROM m_armaster where dbcode <> '-' ORDER BY dbcode;"
                        str_Customers_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT * FROM m_armaster where dbcode <> '-'and deleted = 0 ORDER BY dbcode;"
                        str_Customers_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT * FROM m_armaster where dbcode <> '-' and deleted = 1 ORDER BY dbcode;"
                        str_Customers_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER

                    connect.Open()
                    command1 = New OdbcCommand(str_Customers_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    Dim chkboxcol As New DataGridViewCheckBoxColumn

                    dtData.TableName = "CustTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("Code")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Name 2")
                    dtData.Columns.Add("Address")
                    dtData.Columns.Add("Phone 1")
                    dtData.Columns.Add("Phone 2")
                    dtData.Columns.Add("Mobile")
                    dtData.Columns.Add("Fax")
                    dtData.Columns.Add("Email")
                    'dtData.Columns.Add("Deleted")


                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim cusCode = dtDataRows("dbcode").ToString().Trim()
                        Dim cusName As String = dtDataRows("db_name1").ToString.Trim()
                        Dim cusName2 As String = dtDataRows("db_name2").ToString.Trim()
                        Dim cusAdd As String = dtDataRows("db_add").ToString.Trim()
                        Dim cusTel1 As String = dtDataRows("db_tel1").ToString.Trim()
                        Dim cusTel2 As String = dtDataRows("db_tel2").ToString.Trim()
                        Dim cusHP As String = dtDataRows("db_hp").ToString.Trim()
                        Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        'dtData.Rows.Add(New Object() {cusDel.ToString.Trim(), cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim()})
                        dtData.Rows.Add(New Object() {cusDel.ToString.Trim(), cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusAdd.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 20

                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").Width = 50

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").Width = 300
                    'GridView1.Columns.Item("Name").FillWeight = 248
                    'GridView1.Columns.Item("Name").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").Width = 300
                    GridView1.Columns.Item("Name 2").Visible = False

                    GridView1.Columns.Item("Address").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Address").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Address").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Address").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    'GridView1.Columns.Item("Address").Width = 248
                    GridView1.Columns.Item("Address").FillWeight = 248
                    GridView1.Columns.Item("Address").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").Width = 81

                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").Width = 81
                    GridView1.Columns.Item("Phone 2").Visible = False

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 81

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()


                    connect.Close()
                Catch ex As Exception
                    connect.Close()
                End Try

            End If

            'TO LOAD CUSTOMERS WITH TEXT INSIDE SEARCH BOX
            If txtDB.Text <> "" Then
                searchText = txtDB.Text.ToUpper.Trim.ToString
                'Label1.Text = searchText.ToString

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Customers_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE dbcode LIKE '" & searchText.ToString & "%' ORDER BY dbcode;"
                        str_Customers_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE dbcode LIKE '" & searchText.ToString & "%' and deleted = 0 ORDER BY dbcode;"
                        str_Customers_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE dbcode LIKE '" & searchText.ToString & "%' and deleted = 1 ORDER BY dbcode;"
                        str_Customers_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster WHERE dbcode LIKE '%" & searchText.ToString & "%' and deleted = 0 ORDER BY dbcode;"

                    connect.Open()
                    command1 = New OdbcCommand(str_Customers_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    dtData.TableName = "CustTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("Code")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Name 2")
                    dtData.Columns.Add("Phone 1")
                    dtData.Columns.Add("Phone 2")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim cusCode = dtDataRows("dbcode").ToString().Trim()
                        Dim cusName As String = dtDataRows("db_name1").ToString.Trim()
                        Dim cusName2 As String = dtDataRows("db_name2").ToString.Trim()
                        Dim cusTel1 As String = dtDataRows("db_tel1").ToString.Trim()
                        Dim cusTel2 As String = dtDataRows("db_tel2").ToString.Trim()
                        Dim cusHP As String = dtDataRows("db_hp").ToString.Trim()
                        Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {cusDel.ToString.Trim(), cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 19

                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").Width = 60

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").Width = 248

                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").Width = 248

                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").Width = 81

                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").Width = 81

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 81

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()


                    connect.Close()
                Catch ex As Exception
                    connect.Close()
                End Try

            End If

        ElseIf cmbType.Text = "Name 1" Then

            If txtDB.Text = "" Then

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Customers_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster ORDER BY dbcode;"
                        str_Customers_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster where deleted = 0 ORDER BY dbcode;"
                        str_Customers_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster where deleted = 1 ORDER BY dbcode;"
                        str_Customers_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster where deleted = 0 ORDER BY dbcode;"

                    connect.Open()
                    command1 = New OdbcCommand(str_Customers_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    dtData.TableName = "CustTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("Code")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Name 2")
                    dtData.Columns.Add("Phone 1")
                    dtData.Columns.Add("Phone 2")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim cusCode = dtDataRows("dbcode").ToString().Trim()
                        Dim cusName As String = dtDataRows("db_name1").ToString.Trim()
                        Dim cusName2 As String = dtDataRows("db_name2").ToString.Trim()
                        Dim cusTel1 As String = dtDataRows("db_tel1").ToString.Trim()
                        Dim cusTel2 As String = dtDataRows("db_tel2").ToString.Trim()
                        Dim cusHP As String = dtDataRows("db_hp").ToString.Trim()
                        Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {cusDel.ToString.Trim(), cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 19

                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").Width = 60

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").Width = 248

                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").Width = 248

                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").Width = 81

                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").Width = 81

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 81

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()

                    connect.Close()
                Catch ex As Exception
                    connect.Close()
                End Try

            End If

            'TO LOAD CUSTOMERS WITH TEXT INSIDE SEARCH BOX
            If txtDB.Text <> "" Then
                searchText = txtDB.Text.ToUpper.Trim.ToString
                'Label1.Text = searchText.ToString

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Customers_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE db_name1 LIKE '%" & searchText.ToString & "%' ORDER BY dbcode;"
                        str_Customers_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE db_name1 LIKE '%" & searchText.ToString & "%' and deleted = 0 ORDER BY dbcode;"
                        str_Customers_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE db_name1 LIKE '%" & searchText.ToString & "%' and deleted = 1 ORDER BY dbcode;"
                        str_Customers_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster WHERE db_name1 LIKE '%" & searchText.ToString & "%' and deleted = 0 ORDER BY dbcode;"

                    connect.Open()
                    command1 = New OdbcCommand(str_Customers_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    dtData.TableName = "CustTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("Code")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Name 2")
                    dtData.Columns.Add("Phone 1")
                    dtData.Columns.Add("Phone 2")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim cusCode = dtDataRows("dbcode").ToString().Trim()
                        Dim cusName As String = dtDataRows("db_name1").ToString.Trim()
                        Dim cusName2 As String = dtDataRows("db_name2").ToString.Trim()
                        Dim cusTel1 As String = dtDataRows("db_tel1").ToString.Trim()
                        Dim cusTel2 As String = dtDataRows("db_tel2").ToString.Trim()
                        Dim cusHP As String = dtDataRows("db_hp").ToString.Trim()
                        Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {cusDel.ToString.Trim(), cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 19

                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").Width = 60

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").Width = 248

                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").Width = 248

                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").Width = 81

                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").Width = 81

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 81

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()

                    connect.Close()
                Catch ex As Exception
                    connect.Close()
                End Try

            End If

        ElseIf cmbType.Text = "Name 2" Then

            If txtDB.Text = "" Then

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Customers_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster ORDER BY dbcode;"
                        str_Customers_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster where deleted = 0 ORDER BY dbcode;"
                        str_Customers_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster where deleted = 1 ORDER BY dbcode;"
                        str_Customers_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster where deleted = 0 ORDER BY dbcode;"

                    connect.Open()
                    command1 = New OdbcCommand(str_Customers_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    dtData.TableName = "CustTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("Code")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Name 2")
                    'dtData.Columns.Add("Address")
                    dtData.Columns.Add("Phone 1")
                    dtData.Columns.Add("Phone 2")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim cusCode = dtDataRows("dbcode").ToString().Trim()
                        Dim cusName As String = dtDataRows("db_name1").ToString.Trim()
                        Dim cusName2 As String = dtDataRows("db_name2").ToString.Trim()
                        'Dim cusAdd As String = dtDataRows("db_add").ToString.Trim()
                        Dim cusTel1 As String = dtDataRows("db_tel1").ToString.Trim()
                        Dim cusTel2 As String = dtDataRows("db_tel2").ToString.Trim()
                        Dim cusHP As String = dtDataRows("db_hp").ToString.Trim()
                        Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        'dtData.Rows.Add(New Object() {cusDel.ToString.Trim(), cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusAdd.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim()})
                        dtData.Rows.Add(New Object() {cusDel.ToString.Trim(), cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 19

                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").Width = 60

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").Width = 248

                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").Width = 248

                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").Width = 81

                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").Width = 81

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 81

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()

                    connect.Close()
                Catch ex As Exception
                    connect.Close()
                End Try

            End If

            'TO LOAD CUSTOMERS WITH TEXT INSIDE SEARCH BOX
            If txtDB.Text <> "" Then
                searchText = txtDB.Text.ToUpper.Trim.ToString
                'Label1.Text = searchText.ToString

                Try
                    'ADDED FOR NON-ACTIVE FILTER
                    Dim str_Customers_Show As String

                    If Checkbox1.SelectedIndex.ToString = 0 Then
                        Dim show1 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE db_name2 LIKE '%" & searchText.ToString & "%' ORDER BY dbcode;"
                        str_Customers_Show = show1.ToString
                    ElseIf Checkbox1.SelectedIndex.ToString = 1 Then
                        Dim show2 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE db_name2 LIKE '%" & searchText.ToString & "%' and deleted = 0 ORDER BY dbcode;"
                        str_Customers_Show = show2.ToString
                    Else
                        Dim show3 As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp, deleted FROM m_armaster WHERE db_name2 LIKE '%" & searchText.ToString & "%' and deleted = 1 ORDER BY dbcode;"
                        str_Customers_Show = show3.ToString
                    End If
                    'END OF ADDED FOR NON-ACTIVE FILTER
                    'Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2, db_tel1, db_tel2, db_hp FROM m_armaster WHERE db_name2 LIKE '%" & searchText.ToString & "%' and deleted = 0 ORDER BY dbcode;"

                    connect.Open()
                    command1 = New OdbcCommand(str_Customers_Show, connect)

                    myDataSet = New DataSet()
                    myDataSet.Tables.Clear()
                    myAdapter = New OdbcDataAdapter()
                    myAdapter.SelectCommand = command1
                    myAdapter.Fill(myDataSet, "Db")

                    Dim dtRetrievedData As DataTable = myDataSet.Tables(0)

                    Dim dtData As New DataTable
                    Dim dtDataRows As DataRow

                    dtData.TableName = "CustTable"
                    dtData.Columns.Add("X", GetType(Boolean))
                    dtData.Columns.Add("Code")
                    dtData.Columns.Add("Name")
                    dtData.Columns.Add("Name 2")
                    dtData.Columns.Add("Phone 1")
                    dtData.Columns.Add("Phone 2")
                    dtData.Columns.Add("Mobile")

                    For Each dtDataRows In dtRetrievedData.Rows

                        Dim cusCode = dtDataRows("dbcode").ToString().Trim()
                        Dim cusName As String = dtDataRows("db_name1").ToString.Trim()
                        Dim cusName2 As String = dtDataRows("db_name2").ToString.Trim()
                        Dim cusTel1 As String = dtDataRows("db_tel1").ToString.Trim()
                        Dim cusTel2 As String = dtDataRows("db_tel2").ToString.Trim()
                        Dim cusHP As String = dtDataRows("db_hp").ToString.Trim()
                        Dim cusDel As Boolean = dtDataRows("deleted").ToString.Trim()

                        dtData.Rows.Add(New Object() {cusDel.ToString.Trim(), cusCode.ToString.Trim(), cusName.ToString.Trim(), cusName2.ToString.Trim(), cusTel1.ToString.Trim(), cusTel2.ToString.Trim(), cusHP.ToString.Trim()})
                    Next

                    Dim finalDataSet As New DataSet
                    finalDataSet.Tables.Add(dtData)

                    GridView1.DataSource = finalDataSet.Tables(0)
                    'GridView1.ColumnHeadersVisible = False
                    GridView1.Columns.Item("X").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("X").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    GridView1.Columns.Item("X").Resizable = DataGridViewTriState.False
                    GridView1.Columns.Item("X").Width = 19

                    GridView1.Columns.Item("Code").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Code").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Code").Width = 60

                    GridView1.Columns.Item("Name").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name").Width = 248

                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Name 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Name 2").Width = 248

                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 1").Width = 81

                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Phone 2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Phone 2").Width = 81

                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").HeaderCell.Style.Font = headerFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Font = detailFont
                    GridView1.Columns.Item("Mobile").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Columns.Item("Mobile").Width = 81

                    'GridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
                    GridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    GridView1.Refresh()

                    connect.Close()
                Catch ex As Exception
                    connect.Close()
                End Try

            End If
        End If
        txtDB.Focus()
        txtDB.SelectAll()
    End Sub

    Private Sub frmCustomers_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        GridView1.DataSource = ""
        GridView1.Refresh()
        txtDB.Text = ""
    End Sub
    'THIS WILL PICK KEYPRESSES ON WHOLE FORM
    Private Sub frmCustomers_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub frmCustomers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If cmbType.Items.Count <> 0 Then
            cmbType.SelectedIndex = 0
            lbl4.Text = "Code"
        End If

        If Checkbox1.Items.Count <> 0 Then
            Checkbox1.SelectedIndex = 1
        End If

        txtDB.Text = ""

    End Sub

    'Private Sub cmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged

    '    Dim iRowIndex As Integer

    '    iRowIndex = cmbType.SelectedIndex.ToString
    '    'MessageBox.Show(iRowIndex.ToString.Trim)

    '    If iRowIndex.ToString.Trim = 0 Then
    '        Dim command1 As OdbcCommand
    '        Dim mydataset As DataSet
    '        Dim myadapter As OdbcDataAdapter

    '        connect = New OdbcConnection(sqlConn)
    '        Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2 FROM m_armaster where dbcode <> '' and deleted = 0;"

    '        connect.Open()
    '        command1 = New OdbcCommand(str_Customers_Show, connect)

    '        mydataset = New DataSet()
    '        mydataset.Tables.Clear()
    '        myadapter = New OdbcDataAdapter()
    '        myadapter.SelectCommand = command1
    '        myadapter.Fill(mydataset, "Db")

    '        txtDB.DataSource = mydataset.Tables("Db")
    '        txtDB.DisplayMember = "dbcode"
    '        txtDB.ValueMember = "dbcode"
    '        'txtDB.AutoCompleteMode = AutoCompleteMode.SuggestAppend
    '        'txtDB.AutoCompleteSource = AutoCompleteSource.ListItems
    '        connect.Close()

    '        txtDB.SelectedIndex = -1
    '    End If

    '    If iRowIndex.ToString.Trim = 1 Then
    '        Dim command1 As OdbcCommand
    '        Dim mydataset As DataSet
    '        Dim myadapter As OdbcDataAdapter

    '        connect = New OdbcConnection(sqlConn)
    '        Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2 FROM m_armaster where db_name1 <> '' and deleted = 0;"

    '        connect.Open()
    '        command1 = New OdbcCommand(str_Customers_Show, connect)

    '        mydataset = New DataSet()
    '        mydataset.Tables.Clear()
    '        myadapter = New OdbcDataAdapter()
    '        myadapter.SelectCommand = command1
    '        myadapter.Fill(mydataset, "Db")

    '        txtDB.DataSource = mydataset.Tables("Db")
    '        txtDB.DisplayMember = "db_name1"
    '        txtDB.ValueMember = "db_name1"
    '        'txtDB.AutoCompleteMode = AutoCompleteMode.SuggestAppend
    '        'txtDB.AutoCompleteSource = AutoCompleteSource.ListItems
    '        connect.Close()

    '        txtDB.SelectedIndex = -1
    '    End If

    '    If iRowIndex.ToString.Trim = 2 Then
    '        Dim command1 As OdbcCommand
    '        Dim mydataset As DataSet
    '        Dim myadapter As OdbcDataAdapter

    '        connect = New OdbcConnection(sqlConn)
    '        Dim str_Customers_Show As String = "SELECT dbcode, db_name1, db_name2 FROM m_armaster where db_name2 <> '' and deleted = 0;"

    '        connect.Open()
    '        command1 = New OdbcCommand(str_Customers_Show, connect)

    '        mydataset = New DataSet()
    '        mydataset.Tables.Clear()
    '        myadapter = New OdbcDataAdapter()
    '        myadapter.SelectCommand = command1
    '        myadapter.Fill(mydataset, "Db")

    '        txtDB.DataSource = mydataset.Tables("Db")
    '        txtDB.DisplayMember = "db_name2"
    '        txtDB.ValueMember = "db_name2"
    '        'txtDB.AutoCompleteMode = AutoCompleteMode.SuggestAppend
    '        'txtDB.AutoCompleteSource = AutoCompleteSource.ListItems
    '        connect.Close()

    '        txtDB.SelectedIndex = -1
    '    End If

    'End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click

        openCustomerDialog()

    End Sub

    Private Sub openCustomerDialog()
        diaCustomers.TopLevel = False
        diaCustomers.TopMost = True
        frmMain.Panel1.Controls.Add(diaCustomers)
        diaCustomers.Location = New Point(Convert.ToInt32(frmMain.Panel1.Size.Width / 2 - diaCustomers.Width / 2),
                                   Convert.ToInt32(frmMain.Panel1.Size.Height / 2 - diaCustomers.Height / 2))
        diaCustomers.Show()
        diaCustomers.BringToFront()

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Dim iRowIndex As Integer
        Dim itm1 As String
        Dim del As Boolean

        Try
            iRowIndex = GridView1.CurrentRow.Index

            If cmbType.Text = "Code" Then
                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    MessageBox.Show("Customer already deleted", "Error", MessageBoxButtons.OK)
                    Exit Try
                Else
                    diaCustomers.Close()
                    itm1 = GridView1.Item(1, iRowIndex).Value

                    diaCustomers.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                    openCustomerDialog()
                End If

            End If

            If cmbType.Text = "Name 1" Then
                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    MessageBox.Show("Customer already deleted", "Error", MessageBoxButtons.OK)
                    Exit Try
                Else
                    itm1 = GridView1.Item(1, iRowIndex).Value

                    diaCustomers.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                    openCustomerDialog()
                End If

            End If

            If cmbType.Text = "Name 2" Then
                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    MessageBox.Show("Customer already deleted", "Error", MessageBoxButtons.OK)
                    Exit Try
                Else
                    itm1 = GridView1.Item(1, iRowIndex).Value

                    diaCustomers.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                    openCustomerDialog()
                End If

            End If

        Catch ex As Exception
            MessageBox.Show("Please Select a Customer", "Edit Debtor/Customer", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim iRowIndex As Integer

        Try

            iRowIndex = GridView1.CurrentRow.Index

            Dim del As String
            del = GridView1.Item(0, iRowIndex).Value
            If del = True Then
                MessageBox.Show("Customer already deleted", "Error", MessageBoxButtons.OK)
                Exit Try
            Else
                Dim result As Integer = MessageBox.Show("Confirm Delete Customer?", "Delete Debtor/Customer", MessageBoxButtons.OKCancel)
                If result = DialogResult.OK Then

                    ''MAKE CONDITION TO CHECK IF GOT TRANSACTION HERE

                    'IF THERE IS TRANSACTION THEN
                    'DO NOT ALLOW DELETE
                    'ELSE
                    Dim updtcommand As OdbcCommand
                    Dim updtadapter As OdbcDataAdapter

                    connect = New OdbcConnection(sqlConn)
                    connect.Open()

                    Dim del_cust As String = "update m_armaster set deleted = 1 where dbcode = '" & GridView1.Item(1, iRowIndex).Value & "';"

                    updtcommand = New OdbcCommand(del_cust, connect)
                    updtadapter = New OdbcDataAdapter()

                    updtadapter.UpdateCommand = updtcommand
                    updtadapter.UpdateCommand.ExecuteNonQuery()

                    MessageBox.Show("Customer Deleted.", "Delete Debtor/Customer", MessageBoxButtons.OK)
                    'END IF

                ElseIf result = DialogResult.Cancel Then

                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Please Select a Customer", "Delete Debtor/Customer", MessageBoxButtons.OK)
            'MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub cmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        If cmbType.SelectedIndex = 0 Then
            lbl4.Text = "Code"
            txtDB.Select()
            Exit Sub
        End If

        If cmbType.SelectedIndex = 1 Then
            lbl4.Text = "Name 1"
            txtDB.Select()
            Exit Sub
        End If

        If cmbType.SelectedIndex = 2 Then
            lbl4.Text = "Name 2"
            txtDB.Select()
            Exit Sub
        End If
    End Sub

    Private Sub ViewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewToolStripMenuItem.Click
        viewDetail()
    End Sub

    Private Sub GridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles GridView1.CellDoubleClick
        viewDetail()
    End Sub

    Private Sub viewDetail()
        Dim iRowIndex As Integer
        Dim itm1 As String
        Dim del As Boolean

        Try
            iRowIndex = GridView1.CurrentRow.Index

            If cmbType.Text = "Code" Then
                diaCustomers.Close()
                itm1 = GridView1.Item(1, iRowIndex).Value

                diaCustomers.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                diaCustomers.SaveToolStripMenuItem.Visible = False
                diaCustomers.CancelToolStripMenuItem.Visible = False
                diaCustomers.txt1.Enabled = False
                diaCustomers.txt2.Enabled = False
                diaCustomers.txt3.Enabled = False
                diaCustomers.txt4.Enabled = False
                diaCustomers.txt5.Enabled = False
                diaCustomers.txt6.Enabled = False
                diaCustomers.txt7.Enabled = False
                diaCustomers.txt8.Enabled = False
                diaCustomers.txt9.Enabled = False
                diaCustomers.txt10.Enabled = False

                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    diaCustomers.EditToolStripMenuItem.Visible = False
                Else
                    diaCustomers.EditToolStripMenuItem.Visible = True
                End If
                openCustomerDialog()

            End If

            If cmbType.Text = "Name 1" Then
                diaCustomers.Close()
                itm1 = GridView1.Item(1, iRowIndex).Value

                diaCustomers.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                diaCustomers.SaveToolStripMenuItem.Visible = False
                diaCustomers.CancelToolStripMenuItem.Visible = False
                diaCustomers.txt1.Enabled = False
                diaCustomers.txt2.Enabled = False
                diaCustomers.txt3.Enabled = False
                diaCustomers.txt4.Enabled = False
                diaCustomers.txt5.Enabled = False
                diaCustomers.txt6.Enabled = False
                diaCustomers.txt7.Enabled = False
                diaCustomers.txt8.Enabled = False
                diaCustomers.txt9.Enabled = False
                diaCustomers.txt10.Enabled = False

                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    diaCustomers.EditToolStripMenuItem.Visible = False
                Else
                    diaCustomers.EditToolStripMenuItem.Visible = True
                End If
                openCustomerDialog()
            End If

            If cmbType.Text = "Name 2" Then
                diaCustomers.Close()
                itm1 = GridView1.Item(1, iRowIndex).Value

                diaCustomers.ValidateExisting.Text = itm1.ToString.ToUpper.Trim
                diaCustomers.SaveToolStripMenuItem.Visible = False
                diaCustomers.CancelToolStripMenuItem.Visible = False
                diaCustomers.txt1.Enabled = False
                diaCustomers.txt2.Enabled = False
                diaCustomers.txt3.Enabled = False
                diaCustomers.txt4.Enabled = False
                diaCustomers.txt5.Enabled = False
                diaCustomers.txt6.Enabled = False
                diaCustomers.txt7.Enabled = False
                diaCustomers.txt8.Enabled = False
                diaCustomers.txt9.Enabled = False
                diaCustomers.txt10.Enabled = False

                del = GridView1.Item(0, iRowIndex).Value
                If del = True Then
                    diaCustomers.EditToolStripMenuItem.Visible = False
                Else
                    diaCustomers.EditToolStripMenuItem.Visible = True
                End If
                openCustomerDialog()
            End If

        Catch ex As Exception
            MessageBox.Show("Please Select a Customer", "Edit Debtor/Customer", MessageBoxButtons.OK)
        End Try
    End Sub

End Class