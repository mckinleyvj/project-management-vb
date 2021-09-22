Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl
Imports System.Deployment.Application

Public Class frmMain

    Dim sqlConn As String = frmConfig.theConnection

    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 9, FontStyle.Bold)
    Dim detailFont As New Font("Segoe UI", 8, FontStyle.Regular)

    Public theUser As String

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.BackColor = Color.White
        Me.ForeColor = Color.Black
        MenuStrip1.BackColor = Color.Gainsboro
        MenuStrip1.ForeColor = Color.Black

        Me.WindowState = FormWindowState.Maximized
        'Me.Top = 0
        'Me.Left = 0
        'Me.Height = Screen.PrimaryScreen.WorkingArea.Height
        'Me.Width = Screen.PrimaryScreen.WorkingArea.Width

        PictureBox1.Visible = False

        connect = New OdbcConnection(sqlConn)

        Try

            connect.Open()

            Dim selCommand As OdbcCommand
            Dim selAdapt As OdbcDataAdapter
            Dim selDS As DataSet

            Dim selString As String = "SELECT U.username AS uName, U.role_code AS rCode, R.menu_code AS mCode, S.menu_app AS theMenu, R.enabled AS rEnable FROM s_users U, s_roles_rights R, s_application S WHERE U.role_code = R.role_code AND S.menu_code = R.menu_code AND U.username = '" & theUser.ToString.ToUpper.Trim() & "';"

            selCommand = New OdbcCommand(selString, connect)

            selDS = New DataSet()
            selDS.Tables.Clear()
            selAdapt = New OdbcDataAdapter()
            selAdapt.SelectCommand = selCommand
            selAdapt.Fill(selDS, "login")

            Dim dtRetrievedData As DataTable = selDS.Tables(0)

            Dim dtData As New DataTable
            Dim dtDataRows As DataRow

            dtData.TableName = "LoginTable"
            dtData.Columns.Add("Username")
            dtData.Columns.Add("Role Code")
            dtData.Columns.Add("Menu Code")
            dtData.Columns.Add("The Menu")
            dtData.Columns.Add("Enabled")

            For Each dtDataRows In dtRetrievedData.Rows

                Dim uCode As String = dtDataRows("uName").ToString().Trim()
                Dim rCode As String = dtDataRows("rCode").ToString.Trim()
                Dim mCode As String = dtDataRows("mCode").ToString.Trim()
                Dim theMenu As String = dtDataRows("theMenu").ToString.Trim()
                Dim enab As Integer = dtDataRows("rEnable").ToString()

                dtData.Rows.Add(New Object() {uCode.ToString.Trim(), rCode.ToString.Trim(), mCode.ToString.Trim(), theMenu.ToString.Trim(), enab.ToString})

                Dim toolstripitem As ToolStripMenuItem
                toolstripitem = CType(MenuStrip1.Items(theMenu.ToString.Trim()), ToolStripItem)

                Dim transmenuitem As ToolStripMenuItem
                transmenuitem = CType(TransactionToolStripMenuItem.DropDownItems(theMenu.ToString.Trim()), ToolStripItem)

                Dim setupmenuitem As ToolStripMenuItem
                setupmenuitem = CType(FileSetupToolStripMenuItem.DropDownItems(theMenu.ToString.Trim()), ToolStripItem)

                Dim reportsitem As ToolStripMenuItem
                reportsitem = CType(ReportsToolStripMenuItem.DropDownItems(theMenu.ToString.Trim()), ToolStripItem)

                If toolstripitem Is Nothing Then
                    'Do nothing
                Else
                    Dim openMenu As String = toolstripitem.Name.ToString.Trim()

                    If enab = 1 Then
                        MenuStrip1.Items(openMenu).Visible = True
                    Else
                        MenuStrip1.Items(openMenu).Visible = False
                    End If
                End If

                If transmenuitem Is Nothing Then
                    'DO NOTHING
                Else
                    Dim openMenu As String = transmenuitem.Name.ToString.Trim()
                    'MessageBox.Show(transmenuitem.Name.ToString, "test", MessageBoxButtons.OK)
                    If enab = 1 Then
                        TransactionToolStripMenuItem.DropDownItems(openMenu).Visible = True
                    Else
                        TransactionToolStripMenuItem.DropDownItems(openMenu).Visible = False
                    End If
                End If

                If setupmenuitem Is Nothing Then
                    'DO NOTHING
                Else
                    Dim openMenu As String = setupmenuitem.Name.ToString.Trim()
                    'MessageBox.Show(transmenuitem.Name.ToString, "test", MessageBoxButtons.OK)
                    If enab = 1 Then
                        FileSetupToolStripMenuItem.DropDownItems(openMenu).Visible = True
                    Else
                        FileSetupToolStripMenuItem.DropDownItems(openMenu).Visible = False
                    End If
                End If

                If reportsitem Is Nothing Then
                    'DO NOTHING
                Else
                    Dim openMenu As String = reportsitem.Name.ToString.Trim()
                    'MessageBox.Show(transmenuitem.Name.ToString, "test", MessageBoxButtons.OK)
                    If enab = 1 Then
                        ReportsToolStripMenuItem.DropDownItems(openMenu).Visible = True
                    Else
                        ReportsToolStripMenuItem.DropDownItems(openMenu).Visible = False
                    End If
                End If

            Next

            Dim finalDataSet As New DataSet
            finalDataSet.Tables.Add(dtData)

            'DataGridView1.DataSource = finalDataSet.Tables(0)
            'DataGridView1.Refresh()

            connect.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        If theUser = "" Then
            ToolStripStatusLabel1.Text = "PROSERV MANAGEMENT SYSTEM " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString()
            Me.Text = "PROSERV MANAGEMENT SYSTEM " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString()
        Else
            ToolStripStatusLabel1.Text = "PROSERV MANAGEMENT SYSTEM " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString() & " - " & theUser
            Me.Text = "PROSERV MANAGEMENT SYSTEM " & ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString() & " - " & theUser
        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
        End
    End Sub

    Private Sub CustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerToolStripMenuItem.Click

        frmCustomers.TopLevel = False
        frmCustomers.TopMost = True
        Me.Panel1.Controls.Add(frmCustomers)
        frmCustomers.Show()

        'frmCustomers.Close()
        'frmCustomers.StartPosition = FormStartPosition.CenterParent
        'frmCustomers.ShowDialog()
    End Sub

    Private Sub EmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmployeeToolStripMenuItem.Click

        frmEmployee.TopLevel = False
        frmEmployee.TopMost = True
        Me.Panel1.Controls.Add(frmEmployee)
        frmEmployee.Show()

        'frmEmployee.Close()
        'frmEmployee.StartPosition = FormStartPosition.CenterParent
        'frmEmployee.ShowDialog()
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click

        frmLogin.Close()
        diaCustomers.Close()
        diaEmployees.Close()
        diaProjType.Close()
        diaRemarks.Close()
        frmConfig.Close()
        frmCustomers.Close()
        frmEmployee.Close()
        'Me.Close()
        frmPROJECTS.Close()
        frmProjectTypes.Close()
        frmRoleRights.Close()
        frmRoles.Close()
        frmUsers.Close()

        frmLogin.StartPosition = FormStartPosition.CenterScreen
        frmLogin.Show()
        Me.Dispose(True)


    End Sub

    Private Sub RolesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RolesToolStripMenuItem.Click
        frmRoles.Close()
        frmRoles.StartPosition = FormStartPosition.CenterScreen
        frmRoles.ShowDialog()
    End Sub

    Private Sub UserAccessToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserAccessToolStripMenuItem.Click
        frmUsers.Close()
        frmUsers.StartPosition = FormStartPosition.CenterScreen
        frmUsers.ShowDialog()
    End Sub

    Private Sub ProjectTypeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProjectTypeToolStripMenuItem.Click
        frmProjectTypes.Close()
        frmProjectTypes.StartPosition = FormStartPosition.CenterScreen
        frmProjectTypes.ShowDialog()
    End Sub

    Private Sub ProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProjectToolStripMenuItem.Click

        frmPROJECTS.Close()
        frmPROJECTS.StartPosition = FormStartPosition.CenterScreen
        frmPROJECTS.ShowDialog()

    End Sub

    Private Sub JobSheetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JobSheetToolStripMenuItem.Click
        frmJOBSHEET.Close()
        frmJOBSHEET.StartPosition = FormStartPosition.CenterScreen
        frmJOBSHEET.ShowDialog()
    End Sub

    Private Sub JobMasterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JobMasterToolStripMenuItem.Click
        frmJobTypes.Close()
        frmJobTypes.StartPosition = FormStartPosition.CenterScreen
        frmJobTypes.ShowDialog()
    End Sub

    Private Sub JobsheetReportListingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JobsheetReportListingToolStripMenuItem.Click
        diaJSListing.Close()
        diaJSListing.StartPosition = FormStartPosition.CenterScreen
        diaJSListing.ShowDialog()
    End Sub
End Class
