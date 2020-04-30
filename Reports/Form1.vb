﻿Imports System
Imports System.IO
Imports System.Text
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListView1.View = View.Details
        ListView1.GridLines = True
        ListView1.Columns.Add("Device", 200)
        ListView1.Columns.Add("Errors", 300)
        ListView1.FullRowSelect = True
    End Sub

    Private openfile As OpenFileDialog

    Private Sub FileOpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileOpenToolStripMenuItem.Click

        Try
            openfile = New OpenFileDialog With {
                .Title = "Select files",
                .CheckFileExists = True,
                .Multiselect = True,
                .RestoreDirectory = True
            }

            If openfile.ShowDialog = Windows.Forms.DialogResult.OK Then

                For i As Integer = 0 To openfile.SafeFileNames.Count() - 1
                    ListBox1.Items.Add(openfile.SafeFileNames(i))
                    Dim vPath As String = openfile.FileNames(i)

                    TextBox3.Text = Directory.GetParent(vPath).Name

                    ListBox2.Items.Add(vPath)

                    Dim allTxt() As String = IO.File.ReadAllLines(vPath, Encoding.Default)

                    'Как сделать сортировку файла по ошибкам?

                    Dim addTxt = From vLine1 In allTxt Where (vLine1.StartsWith(vbTab & "OK") _
                                                           And vLine1 Like "*Закрытие*") OrElse vLine1.StartsWith(vbTab & "ERROR")

                    If addTxt.Count = 0 Then Continue For
                    'TextBox1.AppendText(openfile.SafeFileNames(i))
                    'TextBox1.AppendText(String.Join(vbCrLf, addTxt.ToArray) & vbCrLf)

                    'ListView1.Items.Add(openfile.SafeFileNames(i))
                    'ListView1.Items.Add(String.Join(vbCrLf, addTxt.ToArray) & vbCrLf)

                    ListView1.Items.Add(New ListViewItem({openfile.SafeFileNames(i), String.Join(vbCrLf, addTxt.ToArray)}))
                Next
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click

        If ListBox1.Items.Count > 0 And ListBox2.Items.Count > 0 Then

            TextBox2.Text = Application.StartupPath

            File.WriteAllText(TextBox2.Text & "\Отчет " & TextBox3.Text & ".txt", TextBox1.Text, Encoding.Default)
        Else
            Exit Sub
        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        TextBox1.Clear()
    End Sub


End Class