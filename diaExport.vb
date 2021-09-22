Public Class diaExport

    Public expPath As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim FolderBrowserDialog As New FolderBrowserDialog
        Dim SaveFileDialog As New SaveFileDialog

        'With FolderBrowserDialog
        '    .SelectedPath = "C:\"
        '    .Description = "Select the source directory"
        '    If .ShowDialog = DialogResult.OK Then
        '        TextBox1.Text = (.SelectedPath)
        '    End If
        'End With

        With SaveFileDialog
            .DefaultExt = "*.xml"
            .Filter = "XML Files|*.xml"
            '.CreatePrompt = True
            .InitialDirectory = "C:\"
            If .ShowDialog = DialogResult.OK Then
                TextBox1.Text = .RestoreDirectory
            End If
        End With

    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        expPath = TextBox1.Text.ToString
        diaJSListing.DATAGRIDVIEW_TO_EXCEL((diaJSListing.dgvReport))
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class