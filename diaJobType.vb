Imports System.IO
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.Sql
Imports System.Xml
Imports System.Data.SqlClient
'Imports System.Windows.Forms.PrintPreviewDialog
Imports System.Xml.Xsl

Public Class diaJobType

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

    Private Sub diaJobType_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        diaPROJECTS.reload_cmbProj()
    End Sub

    Private Sub diaJobType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If ValidateExisting.Text = "" Then
            Me.Text = "Tasks/Jobs Details Entry"
        Else
            Me.Text = "Tasks/Jobs Details Entry - " & ValidateExisting.Text
        End If

        'cmbSymbol.SelectedIndex = 0

    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Dim updtCommand As OdbcCommand
        Dim updtAdapter As OdbcDataAdapter

        'added on 14/08
        Dim adjust_type As String = ""

        If cmbSymbol.SelectedItem = "" Then
            adjust_type = " "
        Else
            adjust_type = cmbSymbol.SelectedItem.ToString()
        End If

        If ValidateExisting.Text = "" Then

            If txt1.Text = "" Then
                validation1.Visible = True
            End If

            If txt1.Text <> "" Then
                validation1.Visible = False
                Try

                    connect = New OdbcConnection(sqlConn)

                    Dim inscommand2 As OdbcCommand
                    Dim insadapter2 As OdbcDataAdapter

                    Dim dC As String = txt1.Text.ToString.Trim
                    Dim dN1 As String = txt2.Text.ToString.Trim

                    connect.Open()
                    Dim insProj As String = "insert into m_jobs_type(job_type, job_name, adjust_type) values(?,?,?)"
                    inscommand2 = New OdbcCommand(insProj, connect)
                    insadapter2 = New OdbcDataAdapter()

                    inscommand2.Parameters.Add("@job_type", OdbcType.VarChar, 20).Value = dC.ToString
                    inscommand2.Parameters.Add("@job_name", OdbcType.VarChar, 250).Value = dN1.ToString
                    inscommand2.Parameters.Add("@adjust_type", OdbcType.VarChar, 1).Value = adjust_type.ToString
                    insadapter2.InsertCommand = inscommand2
                    insadapter2.InsertCommand.ExecuteNonQuery()

                    MessageBox.Show("Job Type Detail Saved.", "Add New Job Type", MessageBoxButtons.OK)

                    txt1.Clear()
                    txt2.Clear()
                    cmbSymbol.SelectedIndex = -1

                    txt1.Focus()
                    frmJobTypes.RefreshListToolStripMenuItem_Click(Nothing, Nothing)
                    connect.Close()
                    'END OF SAVE FILE INTO DATABASE
                Catch ex As Exception
                    MessageBox.Show("Error Saving into Job Type Table." & vbCrLf & vbCrLf & "Error Message : " & ex.ToString, "Add New Job Type", MessageBoxButtons.OK)
                    txt1.Focus()
                    txt1.SelectAll()
                    'MessageBox.Show(ex.ToString)
                End Try


            End If
        Else
            txt1.Enabled = False
            txt2.Focus()
            txt2.SelectAll()

            Dim result As Integer = MessageBox.Show("Confirm to Edit Job Type Details?", "Edit Job Type", MessageBoxButtons.OKCancel)
            If result = DialogResult.OK Then
                Try

                    connect = New OdbcConnection(sqlConn)
                    connect.Open()

                    Dim dC As String = txt1.Text.ToString.Trim
                    Dim dN1 As String = txt2.Text.ToString.Trim

                    Dim updtProj As String = "update m_jobs_type set job_name = ?, adjust_type = ? where job_type = '" & ValidateExisting.Text.ToString.ToUpper & "';"

                    updtCommand = New OdbcCommand(updtProj, connect)
                    updtAdapter = New OdbcDataAdapter()

                    updtCommand.Parameters.AddWithValue("@job_name", dN1.ToString)
                    updtCommand.Parameters.AddWithValue("@adjust_type", adjust_type.ToString)

                    updtAdapter.UpdateCommand = updtCommand
                    updtAdapter.UpdateCommand.ExecuteNonQuery()

                    MessageBox.Show("Job Type Details Saved.", "Edit Job Type", MessageBoxButtons.OK)
                    frmProjectTypes.RefreshListToolStripMenuItem_Click(Nothing, Nothing)

                    connect.Close()

                Catch ex As Exception
                    MessageBox.Show("Error Saving into Job Type Table." & vbCrLf & vbCrLf & "Error Message : " & ex.ToString, "Edit Job Type", MessageBoxButtons.OK)
                End Try
                connect.Close()

                txt1.Clear()
                txt2.Clear()

                Me.Close()

            ElseIf result = DialogResult.Cancel Then
                'DO NOTHING
            End If
        End If

        frmJobTypes.RefreshListToolStripMenuItem_Click(Nothing, Nothing)
    End Sub

    Private Sub CancelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelToolStripMenuItem.Click
        txt1.Text = ""
        txt2.Text = ""
        txt1.Focus()
        Me.Close()
    End Sub
End Class