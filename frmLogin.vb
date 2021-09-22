Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class frmLogin

    'DIM sqlConn As String = "DRIVER={MySQL ODBC 3.51 Driver};" _
    '                         & "SERVER=192.168.100.155;" _
    '                         & "UID=root;" _
    '                         & "PWD=password;" _
    '                         & "DATABASE=wg_tms_db_90;"

    Dim sqlConn As String = frmConfig.theConnection

    Dim connect As OdbcConnection

    Dim headerFont As New Font("Arial", 8, FontStyle.Bold)
    Dim detailFont As New Font("Calibri", 6, FontStyle.Regular)

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.BackColor = Color.Gainsboro
        Me.ForeColor = Color.Black
        Me.Text = "Service Management System"

        Dim line As String

        If File.Exists(frmConfig.myfile) = True Then
            Using reader As StreamReader = New StreamReader(frmConfig.myfile)
                line = reader.ReadLine
            End Using

            Dim theConfig As String = line.ToString.Trim()
            frmConfig.theConnection = theConfig.ToString.Trim()
            Label3.Text = frmConfig.theConnection.ToString.Trim()
            sqlConn = Label3.Text.ToString.Trim()
        Else
            Dim theConfiguration As String = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=192.168.100.2;UID=root;PWD=password;DATABASE=wg_tms_db_90;"
            Dim objWriter As New System.IO.StreamWriter(frmConfig.myfile)
            objWriter.Write(theConfiguration)
            objWriter.Close()

            Using reader As StreamReader = New StreamReader(frmConfig.myfile)
                line = reader.ReadLine
            End Using

            Dim theConfig As String = line.ToString.Trim()
            frmConfig.theConnection = theConfig.ToString.Trim()
            Label3.Text = frmConfig.theConnection.ToString.Trim()
            sqlConn = Label3.Text.ToString.Trim()
        End If

        txtUser.Focus()
        
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click

        Try

            Dim checkcommand As OdbcCommand
            Dim checkadapter As OdbcDataAdapter
            Dim checkDS As DataSet

            connect = New OdbcConnection(sqlConn)
            connect.Open()

            Dim usrnme As String = txtUser.Text.ToUpper.Trim.ToString
            Dim pw As String = txtPW.Text.ToUpper.Trim.ToString
            Dim str_check_login As String = "SELECT username, password from s_users where username = '" & usrnme.ToString.ToUpper.Trim & "';"

            checkcommand = New OdbcCommand(str_check_login, connect)

            checkDS = New DataSet()
            checkDS.Tables.Clear()
            checkadapter = New OdbcDataAdapter()
            checkadapter.SelectCommand = checkcommand
            checkadapter.Fill(checkDS, "login")

            Dim dtRetrievedData As DataTable = checkDS.Tables("login")

            Dim dtData As New DataTable
            'Dim dtDataRows As DataRow

            Dim chkboxcol As New DataGridViewCheckBoxColumn

            dtData.TableName = "LoginTable"
            dtData.Columns.Add("username")
            dtData.Columns.Add("password")

            If dtRetrievedData.Rows.Count <> 0 Then

                For Each dtDataRows In dtRetrievedData.Rows

                    Dim frmDB_usrnm As String = dtDataRows("username").ToString().Trim().ToUpper()
                    Dim frmDB_pw As String = dtDataRows("password").ToString.Trim().ToUpper()
                    'theUser = frmDB_usrnsm.ToString.ToUpper.Trim

                    dtData.Rows.Add(New Object() {frmDB_usrnm.ToUpper.ToString.Trim, frmDB_pw.ToString.Trim.ToUpper})

                    'MessageBox.Show("Username is : " & frmDB_usrnm.ToUpper.ToString.Trim & " and the password is : " & frmDB_pw.ToString.Trim.ToUpper & ".", "M-Z Time Management System", MessageBoxButtons.OK)
                    If pw.ToUpper.ToString.Trim = frmDB_pw.ToString.Trim.ToUpper Then
                        frmMain.theUser = frmDB_usrnm.ToString.ToUpper.Trim
                        connect.Close()
                        txtUser.Text = ""
                        txtPW.Text = ""
                        txtUser.Focus()
                        frmMain.StartPosition = FormStartPosition.CenterScreen
                        frmMain.Show()
                        Me.Dispose(True)
                    Else
                        MessageBox.Show("Password Combination is not correct.", "ProServ Management System", MessageBoxButtons.OK)
                        txtPW.Focus()
                        txtPW.SelectAll()
                        Exit Sub
                    End If

                Next
                'MessageBox.Show("There are " & dtRetrievedData.Rows.Count.ToString & " users found.", "M-Z Time Management System", MessageBoxButtons.OK)
            Else
                MessageBox.Show("Username not found.", "", MessageBoxButtons.OK)
                txtPW.Text = ""
                txtUser.Focus()
                txtUser.SelectAll()
                Exit Sub
            End If

        Catch ex As Exception
            MessageBox.Show("There's an error with your login username. Or the System is not connected to any Database. Please check with the below error message :" & vbCrLf & vbCrLf & ex.ToString, "ProServ Management System", MessageBoxButtons.OK)
            connect.Close()
        End Try
        connect.Close()

    End Sub

    Private Sub txtUser_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUser.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtPW.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            btnClear.PerformClick()
        End If

    End Sub

    Private Sub txtPW_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPW.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnLogin.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            txtPW.Text = ""
            txtUser.Focus()
            txtUser.SelectAll()
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click

        Dim result As Integer = MessageBox.Show("Exit Application Completely?", "ProServ Management System", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            Application.Exit()
        ElseIf result = DialogResult.No Then
            'Do Nothing
        End If

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        txtUser.Text = ""
        txtPW.Text = ""
        txtUser.Focus()
        frmConfig.StartPosition = FormStartPosition.CenterScreen
        frmConfig.Show()
        Me.Dispose(True)
    End Sub

    Private Sub btnLogin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnLogin.KeyDown
        If e.KeyCode = Keys.Escape Then
            btnClear.PerformClick()
        End If
    End Sub

End Class