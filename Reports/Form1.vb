Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Text
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form1


    ''' <summary>
    ''' ОК - предварительная загрузка окна с таблицей
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListView1.View = View.Details
        ListView1.GridLines = True
        ListView1.Columns.Add("Device", 200)
        ListView1.Columns.Add("Errors", 400)
        ListView1.FullRowSelect = True
    End Sub

    Private openfile As OpenFileDialog

    ''' <summary>
    ''' ОК - таблица с данными и файл в excel
    ''' 
    ''' Задачи:
    ''' Как выбрать папку, а не файлы?
    ''' Скрыть Таблицу     
    ''' Вывести названия папок 
    ''' Вывести системные сообщения
    ''' Открывать папку с файлом в конце
    ''' </summary>
    Private Sub FileOpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileOpenToolStripMenuItem.Click

        Try
            openfile = New OpenFileDialog With {
                .Title = "Select files",
                .CheckFileExists = True,
                .Multiselect = True,
                .RestoreDirectory = True
            }

            If openfile.ShowDialog = System.Windows.Forms.DialogResult.OK Then

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

                sortColumnError()
                findFolderName()
                exportToExcel()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub

    ''' <summary>
    ''' ОК - берет название папки выбранных файлов
    ''' </summary>
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

    ''' <summary>
    ''' ОК - выгрузка из таблицы в файл
    ''' </summary>
    Public Sub exportToExcel()
        Try
            Dim objExcel As New Excel.Application
            Dim bkWorkBook As Excel.Workbook
            Dim shWorkSheet As Excel.Worksheet
            Dim i As Integer, j As Integer

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

            bkWorkBook.SaveAs(TextBox2.Text & "\Отчет " & Format(Now, "yyyyMMdd") & " " & TextBox3.Text)
            bkWorkBook.Close()
            objExcel.Quit()

        Catch ex As Exception
            MsgBox("Export excel error " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ОК - очистка по кнопке для повторного исрользования
    ''' </summary>
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListView1.Items.Clear()
    End Sub

    ''' <summary>
    ''' ERROR - сортировка
    ''' </summary>
    Public Sub sortColumnError()
        Try
            Dim sortColumn As Integer = -1

            'ListView1.Sorting = SortOrder.Ascending
            'ListView1.Sort()



        Catch ex As Exception
            MsgBox("Dont sorting " & ex.Message)
        End Try
    End Sub

    ' Sorting Method
    Dim m_SortingColumn As New ColumnHeader
    Private Sub ListView1_ColumnClick(ByVal sender As _
        System.Object, ByVal e As _
        System.Windows.Forms.ColumnClickEventArgs) Handles _
        ListView1.ColumnClick

        ' Get the new sorting column.
        Dim new_sorting_column As ColumnHeader =
            ListView1.Columns(e.Column)

        ' Figure out the new sorting order.
        Dim sort_order As System.Windows.Forms.SortOrder
        If m_SortingColumn Is Nothing Then
            ' New column. Sort ascending.
            sort_order = SortOrder.Ascending
        Else
            ' See if this is the same column.
            If new_sorting_column.Equals(m_SortingColumn) Then
                ' Same column. Switch the sort order.
                If m_SortingColumn.Text.StartsWith("> ") Then
                    sort_order = SortOrder.Descending
                Else
                    sort_order = SortOrder.Ascending
                End If
            Else
                ' New column. Sort ascending.
                sort_order = SortOrder.Ascending
            End If

            ' Remove the old sort indicator.
            m_SortingColumn.Text =
                m_SortingColumn.Text.Substring(2)
        End If

        ' Display the new sort order.
        m_SortingColumn = new_sorting_column
        If sort_order = SortOrder.Ascending Then
            m_SortingColumn.Text = "> " & m_SortingColumn.Text
        Else
            m_SortingColumn.Text = "< " & m_SortingColumn.Text
        End If

        ' Create a comparer.
        ListView1.ListViewItemSorter = New _
            ListViewComparer(e.Column, sort_order)

        ' Sort.
        ListView1.Sort()

        ' Refresh resultTextBox
        For Each item In ListView1.CheckedItems
            Dim sitem As ListViewItem = item
            sitem.Checked = False
            sitem.Checked = True
        Next
    End Sub

End Class