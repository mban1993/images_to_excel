#Include XL.ahk
xl.Application.DisplayAlerts := false
xl.Application.ScreenUpdating := false
xl.Application.Interactive := false


FileSelectFolder, WhichFolder
XL := ComObjCreate("Excel.Application")
XL.WorkBooks.Add
XL.Visible := true
WinMaximize, ahk_exe excel.exe
Sheet := XL.ActiveSheet
ComObjConnect(Sheet, Worksheet_Events)
Sheet.Columns("A").ColumnWidth := 37

Loop {
Loop Files, %WhichFolder%\*.*, R  
{
	FilePath := A_LoopFileFullPath
	SplitPath, A_LoopFileFullPath,,,, namefile
	xlApp := ComObjActive("Excel.Application")
	picRow := (A_Index -1) * 16 + 1 	   ;starting cell for image
 	txtRow := (A_Index -1) * 16 + 15 	   ;starting cell for text
	xlRng := xlApp.Range("A" picRow)                 
	xlApp.ScreenUpdating := false
	xlShape := xlApp.ActiveSheet.Shapes.AddPicture(FilePath, false, true, xlRng.Left, xlRng.Top, -1, -1)
	xlShape.LockAspectRatio := true
	xlShape.Width := 200			   ;size of image
	xlApp.ScreenUpdating := true
	Xl.Range("A" txtRow).Value := namefile         
	
}
break
}

xl.Application.DisplayAlerts := true
xl.Application.ScreenUpdating := true
xl.Application.Interactive := true