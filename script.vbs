Dim args, objExcel

Set args = WScript.Arguments
Set objExcel = CreateObject("Excel.Application")

objExcel.Workbooks.Open args(0)
objExcel.Visible = True

objExcel.Cells(2, 2).Value = args(1)
objExcel.Run "SayHello"

objExcel.Activeworkbook.Save
objExcel.Activeworkbook.Close(0)
objExcel.Quit