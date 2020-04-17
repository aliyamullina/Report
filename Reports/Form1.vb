Imports System
Imports System.IO
Imports System.Text
Public Class Form1
    Private openfile As OpenFileDialog 'window to open files 
    Private Sub FileOpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileOpenToolStripMenuItem.Click

        Try
            openfile = New OpenFileDialog With {
                .Title = "Select files",
                .CheckFileExists = True,
                .Multiselect = True,
                .RestoreDirectory = True
            }

            If openfile.ShowDialog = Windows.Forms.DialogResult.OK Then
                openfile.SafeFileNames.Count()
                For i As Integer = 0 To openfile.SafeFileNames.Count() - 1
                    Dim vName = ListBox1.Items.Add(openfile.SafeFileNames(i))
                    Dim vPath = ListBox2.Items.Add(openfile.FileNames(i))
                    Dim vReport = ListBox3.Items.AddRange(File.ReadAllLines(vPath))
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