Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListView1.View = View.Details
        ListView1.GridLines = True
        ListView1.Columns.Add("Device", 200)
        ListView1.Columns.Add("Errors", 400)
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

            If openfile.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                'Как выбрать папку, а не файлы?

                For i As Integer = 0 To openfile.SafeFileNames.Count() - 1
                    ListBox1.Items.Add(openfile.SafeFileNames(i))
                    Dim vPath As String = openfile.FileNames(i)

                    TextBox3.Text = Directory.GetParent(vPath).Name

                    ListBox2.Items.Add(vPath)

                    Dim allTxt() As String = IO.File.ReadAllLines(vPath, Encoding.Default)

                    Dim addTxt = From vLine1 In allTxt Where (vLine1.StartsWith(vbTab & "OK") _
                                                           And vLine1 Like "*Закрытие*") _
                                                           OrElse vLine1.StartsWith(vbTab & "ERROR")

                    If addTxt.Count = 0 Then Continue For

                    Dim deviceName = openfile.SafeFileNames(i)
                    Dim deviceReport = String.Join(vbCrLf, addTxt.ToArray)

                    ListView1.Items.Add(New ListViewItem({deviceName, deviceReport}))

                Next

                findFolderName()
                exportToExcel()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListView1.Items.Clear()
    End Sub

    Public Sub findFolderName()
        Try
            If ListBox1.Items.Count > 0 And ListBox2.Items.Count > 0 Then

                TextBox2.Text = System.Windows.Forms.Application.StartupPath
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Folder name error " & ex.Message)
        End Try

    End Sub

    Public Sub exportToExcel()
        Try
            Dim objExcel As New Excel.Application
            Dim bkWorkBook As Excel.Workbook
            Dim shWorkSheet As Excel.Worksheet
            Dim i As Integer
            Dim j As Integer

            objExcel = New Excel.Application
            bkWorkBook = objExcel.Workbooks.Add
            shWorkSheet = CType(bkWorkBook.ActiveSheet, Excel.Worksheet)

            For i = 0 To Me.ListView1.Columns.Count - 1
                shWorkSheet.Cells(1, i + 1) = Me.ListView1.Columns(i).Text
            Next

            For i = 0 To Me.ListView1.Items.Count - 1
                For j = 0 To Me.ListView1.Items(i).SubItems.Count - 1
                    shWorkSheet.Cells(i + 2, j + 1) = Me.ListView1.Items(i).SubItems(j).Text
                Next
            Next

            objExcel.Columns.AutoFit()

            'Как сделать сортировку файла по столбцу B?
            ' Ok, порт, нет ответа
            'objExcel.Range("")

            With shWorkSheet
                'определяем диапазон со 2-й строки до последней заполненной строки в стобце А
                'Dim xlApp As New Excel.Application 'приложение Excel
                'Dim xlWB As Excel.Workbook 'книга
                'Dim xlSht As Excel.Worksheet 'лист
                Dim Rng As Excel.Range 'диапазон ячеек, который будем сотрировать

                'xlWB = xlApp.Workbooks.Open("G:\Excel file.xlsx") 'путь к нашему Excel файлу
                'xlSht = xlWB.Worksheets("Лист1") 'имя листа с данными


                Rng = .Range(.Cells(2, "A"), .Cells(.Rows.Count, "A").End(XlDirection.xlUp))
                'сама сортировка
                .Sort.SortFields.Clear()
                .Sort.SortFields.Add(Key:=Rng, SortOn:=XlSortOn.xlSortOnValues, Order:=XlSortOrder.xlDescending, DataOption:=XlSortDataOption.xlSortNormal) ' XlSortOrder.xlAscending - по возрастанию, XlSortOrder.xlDescending -  по убыванию
                With .Sort
                    .SetRange(Rng)
                    .Header = XlYesNoGuess.xlNo 'xlGuess
                    .MatchCase = False
                    .Orientation = Constants.xlTopToBottom 'XlSortOrientation.xlSortRows 
                    .SortMethod = XlSortMethod.xlPinYin
                    .Apply()
                End With
            End With

            bkWorkBook.SaveAs(TextBox2.Text & "\Отчет " & TextBox3.Text)
            bkWorkBook.Close()
            objExcel.Quit()

        Catch ex As Exception
            MsgBox("Export excel error " & ex.Message)
        End Try
    End Sub

End Class