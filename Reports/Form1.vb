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
                Dim vName = openfile.SafeFileNames.Count()
                For i As Integer = 0 To openfile.SafeFileNames.Count() - 1
                    ListBox1.Items.Add(openfile.SafeFileNames(i))
                    Dim vPath As String = openfile.FileNames(i)
                    ListBox2.Items.Add(vPath)

                    Dim vReport As TextReader = New StreamReader(vPath, encoding:=System.Text.Encoding.Default)
                    TextBox1.Text &= vReport.ReadToEnd()
                    'TextBox1.Text = String.Join(vbCrLf,  )

                    vReport.Close()

                    ' If TextBox1.Text.StartsWith("-") = True Then
                    ' TextBox1.Text = TextBox2.Lines(i)
                    'End If

                    '.StartsWith("-")

                    'Dim vSearch As String = "	ERROR" '"	OK"
                    'Dim vComparison As StringComparison = StringComparison
                Next
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

End Class