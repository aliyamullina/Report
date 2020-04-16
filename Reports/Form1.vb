Imports System.IO
Public Class Form1
    Private openfile As OpenFileDialog 'window to open files 
    Private Sub FileOpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileOpenToolStripMenuItem.Click
        Try
            'set the openfiledialog properties
            openfile = New OpenFileDialog With {
                .Filter = "No",
                .Title = "Select files",
                .CheckFileExists = True,
                .Multiselect = True,
                .RestoreDirectory = True
            }

            If openfile.ShowDialog = Windows.Forms.DialogResult.OK Then
                openfile.SafeFileNames.Count()
                For i As Integer = 0 To openfile.SafeFileNames.Count() - 1
                    ListBox1.Items.Add(openfile.SafeFileNames(i))
                    ListBox2.Items.Add(openfile.FileName(i))
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class