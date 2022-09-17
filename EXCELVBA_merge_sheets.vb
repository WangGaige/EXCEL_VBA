EXCEL合并文件及合并工作表(工作薄)的通用方法摘要：文章：EXCEL合并文件及合并工作表(工作薄)的通用方法[原创] 摘要：使用MSOFFICEEXCEL的时候经常遇到：（1）需要将多个Excel文件进行合并；（2）需要将多个Sheet进行合并，发表于北京联高软件有限公司技术文章栏目，代码以高亮显示。
关键字：合并, 文件, excel, 原创, 通用, sheets, thisworkbook, filestoopen, for, count, sub, next, 功能, dim, end
使用MS OFFICE EXCEL的时候经常遇到：
（1）需要将多个 Excel 文件进行合并；
（2）需要将多个 Sheet 进行合并；
这里给出最佳答案。当然您得需要会使用宏(MICRO)。

功能一：合并Excel文件
Sub CombineWorkbooks()
Dim FilesToOpen, ft
Dim x As Integer
Application.ScreenUpdating = False
On Error GoTo errhandler

FilesToOpen = Application.GetOpenFilename _
(FileFilter:="Micrsofe Excel文件(*.xls), *.xls", _
MultiSelect:=True, Title:="要合并的文件")

If TypeName(FilesToOpen) = "boolean" Then
MsgBox "没有选定文件"
End If

x = 1
While x <= UBound(FilesToOpen)
Set wk = Workbooks.Open(Filename:=FilesToOpen(x))

wk.Sheets().Move after:=ThisWorkbook.Sheets _
(ThisWorkbook.Sheets.Count)
x = x + 1
Wend

MsgBox "合并成功完成！"

errhandler:
End Sub

功能二：合并任意的 Sheet
合并之前，请先创建一个空白的 Sheet 作为合并目标 Sheet ，这个 Sheet 必须是第一个 Sheet 哦。
如果不合并标题行（比如第一行）则 j=1 改为 j=2
如果数据不是从第一行，或者第一列开始的，请修改 j=1 及 k=2 两行的参数。
比如 j=2 k=3 表示从 第2行，第三列开始的数据。

Sub CombineSheet()

Dim i, j, k, n As Integer
n = 1
For i = 2 To ThisWorkbook.Sheets.Count
For j = 1 To ThisWorkbook.Sheets(i).UsedRange.Rows.Count
For k = 1 To ThisWorkbook.Sheets(i).UsedRange.Columns.Count
ThisWorkbook.Sheets(1).Cells(n, k).Value = ThisWorkbook.Sheets(i).Cells(j, k).Value
Next k
n = n + 1
Next j
Next i

End Sub

意外惊喜：合并 Sheet 的功能会自动去掉 超链接(HREF) 标记。
实际上，为了去掉 Excel 的所有超链接，也可以使用这个函数啊。
---------------------------excel文件里有多个sheet，怎样把每个sheet全部导出为单独的xlsexcel文件里有多个sheet，怎样把每个sheet全部导出为单独的xls，还是用原sheet名命名，一个一个的另存为太费劲，有太多sheet 1.Alt+F11 进入VBE2.菜单：插入-模块。3.复制下面的代码到光标处4.Alt+F11回到Excel5.Alt+F8 选Test,点击运  

'将工作簿所有工作表另存为单独的文件。
'路径为原工作簿路径，文件名为工作表名

Sub Test()
    Dim Sht As Worksheet
    For Each Sht In Sheets
        Sht.Copy
        ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & Sht.Name & ".xls"
        ActiveWorkbook.Close
    Next
End Sub 
