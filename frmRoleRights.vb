Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl
Imports System.Text.RegularExpressions

Public Class frmRoleRights

    'Dim sqlConn As String = "DRIVER={MySQL ODBC 3.51 Driver};" _
    '                          & "SERVER=localhost;" _
    '                          & "UID=root;" _
    '                          & "PWD=password;" _
    '                          & "DATABASE=wg_tms_db_90;"

    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Private Sub frmRoleRights_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim theCode As String = txtRoleCode.Text.ToString.ToUpper.Trim()

        Dim command1 As OdbcCommand
        Dim myAdapter As OdbcDataAdapter
        Dim myDataSet As DataSet

        connect = New OdbcConnection(sqlConn)

        Try

            Dim str_rights As String = "SELECT A.role_code AS theRole, A.menu_code AS theMenu, B.menu_desc AS theMenuName, A.enabled AS tick FROM s_roles_rights A, s_application B WHERE A.menu_code = B.menu_code AND A.role_code = '" & theCode.ToString & "' ORDER BY B.menu_code;"
            'Dim str_rights As String = "SELECT B.menu_code AS theMenu, B.menu_desc AS theMenuName FROM s_application B;"

            connect.Open()
            command1 = New OdbcCommand(str_rights, connect)

            myDataSet = New DataSet()
            myDataSet.Tables.Clear()
            myAdapter = New OdbcDataAdapter()
            myAdapter.SelectCommand = command1
            myAdapter.Fill(myDataSet, "app")

            Dim dtRetrievedData As DataTable = myDataSet.Tables("app")

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            'Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "app"
            dtData.Columns.Add("Allow", GetType(Boolean))
            dtData.Columns.Add("Application Menu")
            dtData.Columns.Add("Menu")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim ticked As String = dtDataRows("tick").ToString.Trim()
                Dim finalTick As Boolean

                If ticked.ToString.Trim() = "1" Then
                    finalTick = True
                ElseIf ticked.ToString.Trim() = "0" Then
                    finalTick = False
                Else
                    finalTick = False
                End If

                Dim theMenu_Desc As String = dtDataRows("theMenuName").ToString.Trim()

                Dim theMenu_code As String = dtDataRows("theMenu").ToString.Trim()

                dtData.Rows.Add(New Object() {finalTick.ToString(), theMenu_Desc.ToString.Trim(), theMenu_code.ToString.Trim()})

            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            DataGridView1.DataSource = finalDataSet.Tables(0)
            'DataGridView1.Columns(0).ReadOnly = True
            'For i As Int32 = 1 To DataGridView1.Columns.Count - 1
            '    DataGridView1.Columns(i).ReadOnly = False
            'Next

            DataGridView1.Columns.Item("Allow").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns.Item("Allow").HeaderCell.Style.Font = headerFont
            DataGridView1.Columns.Item("Allow").DefaultCellStyle.Font = detailFont
            DataGridView1.Columns.Item("Allow").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns.Item("Allow").Width = 58

            DataGridView1.Columns.Item("Application Menu").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            DataGridView1.Columns.Item("Application Menu").HeaderCell.Style.Font = headerFont
            DataGridView1.Columns.Item("Application Menu").DefaultCellStyle.Font = detailFont
            DataGridView1.Columns.Item("Application Menu").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DataGridView1.Columns.Item("Application Menu").FillWeight = 500
            DataGridView1.Columns.Item("Application Menu").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            DataGridView1.Columns.Item("Menu").Visible = False

            DataGridView1.Columns.Item("Allow").SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns.Item("Application Menu").SortMode = DataGridViewColumnSortMode.NotSortable

            'DataGridView1.ColumnHeadersDefaultCellStyle.Font = headerFont
            DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            DataGridView1.Columns.Item("Application Menu").ReadOnly = True
            DataGridView1.Refresh()

            connect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connect.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    'Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)

    '    For Each r As DataGridViewRow In DataGridView1.Rows
    '        r.Cells(1).ReadOnly = False
    '    Next

    'End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        DataGridView1.DataSource = ""
        DataGridView1.Refresh()
        Me.Close()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Dim theCode As String = txtRoleCode.Text.ToString.ToUpper.Trim()

        Try
            connect.Open()
            For Each row As DataGridViewRow In DataGridView1.Rows
                Dim g As String

                Dim req As String = row.Cells("Allow").Value
                Dim reh As String = row.Cells("Application Menu").Value
                Dim rec As String = row.Cells("Menu").Value

                If row.Cells("Allow").Value = True Then
                    g = "1"
                Else
                    g = "0"
                End If

                Dim updtcommand As OdbcCommand
                Dim updtAdapter As OdbcDataAdapter


                Dim query = "update s_roles_rights SET enabled = '" & g.ToString.Trim() & "' WHERE role_code = '" & theCode.ToString.ToUpper.Trim() & "' and menu_code = '" & rec.ToString.Trim() & "';"

                updtcommand = New OdbcCommand(query, connect)
                updtAdapter = New OdbcDataAdapter()

                updtAdapter.UpdateCommand = updtcommand
                updtAdapter.UpdateCommand.ExecuteNonQuery()


            Next
            connect.Close()

            MessageBox.Show("Application Role for " & theCode.ToUpper.Trim() & " has been modified successfully", "Application Role Rights", MessageBoxButtons.OK)
            DataGridView1.ReadOnly = True
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        connect.Close()
    End Sub
End Class