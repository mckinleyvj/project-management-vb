Imports System.IO

Public Class frmConfig

    Dim MySQLServerHost As String
    Dim MySQLServerUName As String
    Dim MySQLServerUPass As String
    Dim MySQLServerDatabaseID As String

    Public Shared theConnection As String
    Public Shared globalDBname As String

    Public myfile As String = Application.StartupPath & "\Config.TMS"

    Private Sub Config_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmLogin.Close()
        frmLogin.StartPosition = FormStartPosition.CenterScreen
        frmLogin.Show()
        Me.Dispose(True)
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        txtHost.Enabled = False
        txtUName.Enabled = False
        txtPass.Enabled = False
        txtDatabase.Enabled = False

        txtHost.Text = ""
        txtUName.Text = ""
        txtPass.Text = ""
        txtDatabase.Text = ""

        btnLogin.Visible = False
        btnCancel.Visible = False
        btnSave.Visible = False
        btnEdit.Visible = True
        btnClose.Visible = True

        validation1.Visible = False
        validation2.Visible = False
        validation3.Visible = False
        validation4.Visible = False

        txtAdmin.Text = ""
        txtPassword.Text = ""
        txtAdmin.Visible = False
        txtPassword.Visible = False

        frmLogin.Close()
        frmLogin.StartPosition = FormStartPosition.CenterScreen
        frmLogin.Show()
        Me.Dispose(True)
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        If txtHost.Text = "" And txtUName.Text = "" And txtPass.Text = "" And txtDatabase.Text = "" Then
            validation1.Visible = True
            validation2.Visible = True
            validation3.Visible = True
            validation4.Visible = True
        End If

        If txtHost.Text <> "" And txtUName.Text = "" And txtPass.Text = "" And txtDatabase.Text = "" Then
            validation1.Visible = False
            validation2.Visible = True
            validation3.Visible = True
            validation4.Visible = True
        End If

        If txtHost.Text = "" And txtUName.Text <> "" And txtPass.Text = "" And txtDatabase.Text = "" Then
            validation1.Visible = True
            validation2.Visible = False
            validation3.Visible = True
            validation4.Visible = True
        End If

        If txtHost.Text = "" And txtUName.Text = "" And txtPass.Text <> "" And txtDatabase.Text = "" Then
            validation1.Visible = True
            validation2.Visible = True
            validation3.Visible = False
            validation4.Visible = True
        End If

        If txtHost.Text = "" And txtUName.Text = "" And txtPass.Text = "" And txtDatabase.Text <> "" Then
            validation1.Visible = True
            validation2.Visible = True
            validation3.Visible = True
            validation4.Visible = False
        End If

        If txtHost.Text <> "" And txtUName.Text <> "" And txtPass.Text = "" And txtDatabase.Text = "" Then
            validation1.Visible = False
            validation2.Visible = False
            validation3.Visible = True
            validation4.Visible = True
        End If

        If txtHost.Text <> "" And txtUName.Text = "" And txtPass.Text <> "" And txtDatabase.Text = "" Then
            validation1.Visible = False
            validation2.Visible = True
            validation3.Visible = False
            validation4.Visible = True
        End If

        If txtHost.Text <> "" And txtUName.Text = "" And txtPass.Text = "" And txtDatabase.Text <> "" Then
            validation1.Visible = False
            validation2.Visible = True
            validation3.Visible = True
            validation4.Visible = False
        End If

        If txtHost.Text = "" And txtUName.Text <> "" And txtPass.Text <> "" And txtDatabase.Text = "" Then
            validation1.Visible = True
            validation2.Visible = False
            validation3.Visible = False
            validation4.Visible = True
        End If

        If txtHost.Text = "" And txtUName.Text <> "" And txtPass.Text = "" And txtDatabase.Text <> "" Then
            validation1.Visible = True
            validation2.Visible = False
            validation3.Visible = True
            validation4.Visible = False
        End If

        If txtHost.Text = "" And txtUName.Text = "" And txtPass.Text <> "" And txtDatabase.Text <> "" Then
            validation1.Visible = True
            validation2.Visible = True
            validation3.Visible = False
            validation4.Visible = False
        End If

        If txtHost.Text <> "" And txtUName.Text <> "" And txtPass.Text = "" And txtDatabase.Text <> "" Then
            validation1.Visible = False
            validation2.Visible = False
            validation3.Visible = True
            validation4.Visible = False
        End If

        If txtHost.Text <> "" And txtUName.Text <> "" And txtPass.Text <> "" And txtDatabase.Text = "" Then
            validation1.Visible = False
            validation2.Visible = False
            validation3.Visible = False
            validation4.Visible = True
        End If

        If txtHost.Text = "" And txtUName.Text <> "" And txtPass.Text <> "" And txtDatabase.Text <> "" Then
            validation1.Visible = True
            validation2.Visible = False
            validation3.Visible = False
            validation4.Visible = False
        End If

        If txtHost.Text <> "" And txtUName.Text = "" And txtPass.Text <> "" And txtDatabase.Text <> "" Then
            validation1.Visible = False
            validation2.Visible = True
            validation3.Visible = False
            validation4.Visible = False
        End If

        If txtHost.Text <> "" And txtUName.Text <> "" And txtPass.Text <> "" And txtDatabase.Text <> "" Then
            validation1.Visible = False
            validation2.Visible = False
            validation3.Visible = False
            validation4.Visible = False
            Try

                MySQLServerHost = txtHost.Text
                MySQLServerUName = txtUName.Text
                MySQLServerUPass = txtPass.Text
                MySQLServerDatabaseID = txtDatabase.Text

                Dim theConfiguration As String = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & MySQLServerHost.ToString.Trim() _
                                                 & ";UID=" & MySQLServerUName.ToString.Trim() & ";PWD=" & MySQLServerUPass.ToString.Trim() _
                                                 & ";DATABASE=" & MySQLServerDatabaseID.ToString.Trim() & ";"

                globalDBname = MySQLServerDatabaseID.ToString.Trim

                Dim objWriter As New System.IO.StreamWriter(myfile)
                objWriter.Write(theConfiguration)
                objWriter.Close()

                txtHost.Text = ""
                txtUName.Text = ""
                txtPass.Text = ""
                txtDatabase.Text = ""
                txtHost.Focus()
                txtConfig.Text = ""

                theConnection = theConfiguration.ToString.Trim

                Dim i As Integer = MessageBox.Show("Configuration saved! " & vbCrLf & theConfiguration.ToString.Trim() & ".", "System Configuration Settings.", MessageBoxButtons.OK)
                If i = DialogResult.OK Then

                    txtConfig.Text = theConnection.ToString
                    btnLogin.Visible = False
                    btnEdit.Visible = True
                    btnClose.Visible = True
                    btnSave.Visible = False
                    btnCancel.Visible = False
                    txtAdmin.Visible = False
                    txtPassword.Visible = False
                    lblAdmin.Visible = False
                    lblPassword.Visible = False

                    txtHost.Enabled = False
                    txtUName.Enabled = False
                    txtPass.Enabled = False
                    txtDatabase.Enabled = False

                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If

        

    End Sub

    Private Sub Config_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim line As String

        If File.Exists(myfile) = True Then
            Using reader As StreamReader = New StreamReader(myfile)
                line = reader.ReadLine
            End Using

            Dim theConfig As String = line.ToString.Trim()
            theConnection = theConfig.ToString.Trim()

            txtConfig.Text = theConfig.ToString.Trim()
        Else

            Dim theConfiguration As String = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;UID=root;PWD=password;DATABASE=wg_tms_db_90;"
            Dim objWriter As New System.IO.StreamWriter(myfile)
            objWriter.Write(theConfiguration)
            objWriter.Close()

            Using reader As StreamReader = New StreamReader(myfile)
                line = reader.ReadLine
            End Using

            Dim theConfig As String = line.ToString.Trim()
            theConnection = theConfig.ToString.Trim()

            txtConfig.Text = theConfig.ToString.Trim()
        End If
    End Sub

    Private Sub btEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

        lblAdmin.Visible = True
        lblPassword.Visible = True
        txtAdmin.Visible = True
        txtPassword.Visible = True

        btnEdit.Visible = False
        btnSave.Visible = False
        'btnClose.Visible = False

        btnLogin.Visible = True
        btnCancel.Visible = True

        txtAdmin.Enabled = False
        txtAdmin.Text = "admin"
        txtPassword.Enabled = True

        txtPassword.Focus()

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        btnEdit.Visible = True
        btnClose.Visible = True
        btnSave.Visible = False

        txtAdmin.Visible = False
        txtPassword.Visible = False
        lblAdmin.Visible = False
        lblPassword.Visible = False

        btnLogin.Visible = False
        btnCancel.Visible = False

        txtAdmin.Text = ""
        txtPassword.Text = ""

        txtHost.Text = ""
        txtUName.Text = ""
        txtPass.Text = ""
        txtDatabase.Text = ""

        txtHost.Enabled = False
        txtUName.Enabled = False
        txtPass.Enabled = False
        txtDatabase.Enabled = False

        validation1.Visible = False
        validation2.Visible = False
        validation3.Visible = False
        validation4.Visible = False

        txtAdmin.Focus()
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click

        Dim username As String = "admin"
        Dim password As String = "adminl0g!n"

        If txtAdmin.Text = "" Then
            MessageBox.Show("Please enter login username", "Authorization Required.", MessageBoxButtons.OK)
            txtAdmin.Focus()
            txtAdmin.SelectAll()
        Else
            If txtPassword.Text = "" Then
                MessageBox.Show("Please enter password", "Authorization Required.", MessageBoxButtons.OK)
                txtPassword.Focus()
                txtPassword.SelectAll()
            Else
                If txtAdmin.Text.ToString.Trim() <> username.ToString.Trim() Then
                    MessageBox.Show("Username not found. Please contact Administrator", "Authorization Required.", MessageBoxButtons.OK)
                    txtAdmin.Focus()
                    txtAdmin.SelectAll()
                Else
                    If txtPassword.Text.ToString.Trim() <> password.ToString.Trim() Then
                        MessageBox.Show("Password does not match", "Authorization Required.", MessageBoxButtons.OK)
                        txtPassword.Focus()
                        txtPassword.SelectAll()
                    Else
                        MessageBox.Show("Log on successfully. Proceed to Configurations.", "Authorization Complete.", MessageBoxButtons.OK)
                        txtAdmin.Text = ""
                        txtPassword.Text = ""

                        txtAdmin.Visible = False
                        txtPassword.Visible = False
                        lblAdmin.Visible = False
                        lblPassword.Visible = False

                        btnLogin.Visible = False
                        btnCancel.Visible = False

                        txtHost.Enabled = True
                        txtUName.Enabled = True
                        txtPass.Enabled = True
                        txtDatabase.Enabled = True

                        txtHost.Focus()
                        btnSave.Visible = True
                        btnCancel.Visible = True

                        validation1.Visible = False
                        validation2.Visible = False
                        validation3.Visible = False
                        validation4.Visible = False

                    End If
                End If
            End If
        End If

    End Sub

    Private Sub txtPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnLogin.Focus()
        End If
    End Sub

    Private Sub txtAdmin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAdmin.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtPassword.Focus()
            txtPassword.SelectAll()
        End If

        If e.KeyCode = Keys.Tab Then
            txtPassword.Focus()
            txtPassword.SelectAll()
        End If
    End Sub

    Private Sub txtHost_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtHost.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtUName.Focus()
        End If
    End Sub

    Private Sub txtUName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUName.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtPass.Focus()
        End If
    End Sub

    Private Sub txtPass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            txtDatabase.Focus()
        End If
    End Sub

    Private Sub txtDatabase_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDatabase.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnSave.Focus()
        End If
    End Sub

End Class