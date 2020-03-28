Option Explicit

Dim objExcel, objExcelWb, objExcelSht
Dim oFileObject
Dim oMstrCmp
Dim ErrMsg
Dim scriptdir
const xlup = -4162

Main

Sub Main()
    Dim rwNum
    Dim ParentValue, ParentDesc, strFolder 
    Dim ChildValue 
    dim ParentRow
    dim child
    'scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
    Set oFileObject = CreateObject("Scripting.FileSystemObject")
    strFolder = oFileObject.GetParentFolderName(WScript.ScriptFullName)

    Set objExcel = CreateObject("Excel.Application")
    Set objExcelWb = objExcel.Workbooks.Open(strFolder & "\rawfeed.xls")
    Set objExcelSht = objExcelWb.Worksheets("Sheet1")
    
    Dim i
    Dim arrayColstoDelete, arrayCellsToChange
    'Unnecessary Columns to delete.  Delete from right to left
    arrayColstoDelete = array(63,62,61,60,59,58,52,51,50,48,47,46,44,43,42,40,39,38,36,23,22,21,19,16,9,7)
    arrayCellsToChange= array(8,10,18,19,20,22) 'These are the columns from Parent to copy into Child rows

    strFolder = oFileObject.GetParentFolderName(WScript.ScriptFullName)
    'We get the Last row of sheet and work our way back up
    For each i in arrayColstoDelete 
        objExcelSht.cells(1,i).entireColumn.Delete
    next
    'We get the Last row of sheet and work our way back up
    rwNum = objExcelSht.Range("A" & objExcelSht.Rows.Count).End(xlUp).Row
      
    Do While rwNum > 1
        'Check if Row is a Parent and get its ID
        if objExcelSht.cells(rwNum,3).Value = "Parent" then
            ParentValue = objExcelSht.cells(rwNum,1).Value
                        
            ParentRow = rwNum
            'get next row up and see if it is a Child
            rwNum = rwNum -1 
                Do Until objExcelSht.cells(rwNum,3).Value <> "Child"
                        For each i in arrayCellsToChange 
                            'Update the Child Row with Parents Values
                            objExcelSht.Cells(rwNum,i).value = objExcelSht.cells(ParentRow,i).Value
                        Next 
                        'Update the Child Name with PArent Name & Suffix
                        if objExcelSht.Cells(rwNum,3).value = "Child" then objExcelSht.Cells(rwNum,9).value = objExcelSht.cells(rwNum,9).Value & " - " & mid(objExcelSht.cells(rwNum,5).Value,4) 
                    
                    rwNum = rwNum - 1
                Loop
        'Delete the Parent Row as we dont need it
        objExcelSht.cells(ParentRow,1).EntireRow.Delete
        else 
        rwNum = rwNum - 1
        End If
        
    Loop
    objExcelWb.SaveAs "" & strfolder & "\ProcessedFeed.xls"
    objExcelWb.close

End Sub

