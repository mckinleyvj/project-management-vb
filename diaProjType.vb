Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class diaProjType

    'Dim sqlConn As String = "DRIVER={MySQL ODBC 3.51 Driver};" _
    '                          & "SERVER=localhost;" _
    '                          & "UID=root;" _
    '                          & "PWD=password;" _
    '                          & "DATABASE=wg_tms_db_90;"
    Dim sqlConn As String = frmConfig.theConnection
    Dim connect As OdbcConnection

    Dim headerFont As New Font("Segoe UI", 8, FontStyle.Regular)
    Dim detailFont As New Font("Segoe UI", 6, FontStyle.Regular)

    
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub diaProjType_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        diaPROJECTS.reload_cmbProj()
    End Sub

    Private Sub diaProjType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ValidateExisting.Text = "" Then
            Me.Text = "Projects Master Details Entry"
        Else
            Me.Text = "Projects Master Details Entry - " & ValidateExisting.Text
        End If


    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Dim updtCommand As OdbcCommand
        Dim updtAdapter As OdbcDataAdapter

        If ValidateExisting.Text = "" Then

            If txt1.Text = "" And txt7.Text = "" Then
                validation1.Visible = True
                validation3.Visible = True
            End If

            If txt1.Text <> "" And txt7.Text = "" Then
                validation1.Visible = False
                validation3.Visible = True
            End If

            If txt1.Text = "" And txt7.Text <> "" Then
                validation1.Visible = True
                validation3.Visible = False
            End If

            If txt1.Text <> "" And txt7.Text <> "" Then
                validation1.Visible = False
                validation3.Visible = False
                Try

                    connect = New OdbcConnection(sqlConn)

                    Dim inscommand2 As OdbcCommand
                    Dim insadapter2 As OdbcDataAdapter

                    Dim dC As String = txt1.Text.ToString.Trim
                    Dim dN1 As String = txt2.Text.ToString.Trim
                    Dim dw1 As String = txt7.Text.ToString
                    Dim dN2 As Decimal = Convert.ToDecimal(dw1.ToString)

                    connect.Open()
                    Dim insProj As String = "insert into m_project_type(project_code, project_name, project_price) values(?,?,?)"
                    inscommand2 = New OdbcCommand(insProj, connect)
                    insadapter2 = New OdbcDataAdapter()

                    inscommand2.Parameters.Add("@project_code", OdbcType.VarChar, 20).Value = dC.ToString
                    inscommand2.Parameters.Add("@project_name", OdbcType.VarChar, 250).Value = dN1.ToString
                    inscommand2.Parameters.Add("@project_price", OdbcType.Double).Value = dN2.ToString
                    insadapter2.InsertCommand = inscommand2
                    insadapter2.InsertCommand.ExecuteNonQuery()

                    MessageBox.Show("Projects detail saved.", "Add New Project", MessageBoxButtons.OK)

                    txt1.Clear()
                    txt2.Clear()
                    txt7.Clear()

                    txt1.Focus()
                    frmProjectTypes.RefreshListToolStripMenuItem_Click(Nothing, Nothing)
                    connect.Close()
                    'END OF SAVE FILE INTO DATABASE
                Catch ex As Exception
                    MessageBox.Show("Error Saving Project detail." & vbCrLf & vbCrLf & "Error Message : " & ex.ToString, "Add New Project", MessageBoxButtons.OK)
                    txt1.Focus()
                    txt1.SelectAll()
                    'MessageBox.Show(ex.ToString)
                End Try


            End If
        Else
            txt1.Enabled = False
            txt2.Focus()
            txt2.SelectAll()

            Dim result As Integer = MessageBox.Show("Confirm to Edit Project details?", "Edit Project", MessageBoxButtons.OKCancel)
            If result = DialogResult.OK Then
                Try

                    connect = New OdbcConnection(sqlConn)
                    connect.Open()

                    Dim dC As String = txt1.Text.ToString.Trim
                    Dim dN1 As String = txt2.Text.ToString.Trim
                    Dim dL As String = txt7.Text.ToString.Trim

                    Dim updtProj As String = "update m_project_type set project_name = ?, project_price = ? where project_code = '" & ValidateExisting.Text.ToString.ToUpper & "';"

                    updtCommand = New OdbcCommand(updtProj, connect)
                    updtAdapter = New OdbcDataAdapter()

                    updtCommand.Parameters.AddWithValue("@project_name", dN1.ToString)
                    updtCommand.Parameters.AddWithValue("@project_price", dL.ToString)

                    updtAdapter.UpdateCommand = updtCommand
                    updtAdapter.UpdateCommand.ExecuteNonQuery()

                    MessageBox.Show("Project details saved.", "Edit Project details", MessageBoxButtons.OK)
                    frmProjectTypes.RefreshListToolStripMenuItem_Click(Nothing, Nothing)

                    connect.Close()

                Catch ex As Exception
                    MessageBox.Show("Error Saving Project details." & vbCrLf & vbCrLf & "Error Message : " & ex.ToString, "Edit Project", MessageBoxButtons.OK)
                End Try
                connect.Close()

                txt1.Clear()
                txt2.Clear()
                txt7.Clear()

                Me.Close()

            ElseIf result = DialogResult.Cancel Then
                'DO NOTHING
            End If
        End If
    End Sub

    Private Sub CancelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelToolStripMenuItem.Click
        txt1.Text = ""
        txt2.Text = ""
        txt7.Text = ""
        txt1.Focus()
        Me.Close()
    End Sub
End Class