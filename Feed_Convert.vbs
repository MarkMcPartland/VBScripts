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
    Dim arrayCellsToChange
    arrayCellsToChange= array(8,12,24,25,26,28) 'These are the columns from Parent to copy into Child rows

    strFolder = oFileObject.GetParentFolderName(WScript.ScriptFullName)
    'We get the Last row of sheet and work our way back up
    rwNum = objExcelSht.Range("A" & objExcelSht.Rows.Count).End(xlUp).Row
      
    Do While rwNum > 1
        'Check if Row is a Parent and get its ID
        if objExcelSht.cells(rwNum,3).Value = "Parent" then
            ParentValue = objExcelSht.cells(rwNum,1).Value
                        
            ParentRow = rwNum
            'get next row up
            rwNum = rwNum -1 
                Do Until objExcelSht.cells(rwNum,3).Value <> "Child"
                        For each i in arrayCellsToChange 
                            'Update the Child Row with Parents Values
                            objExcelSht.Cells(rwNum,i).value = objExcelSht.cells(ParentRow,i).Value
                        Next 
                        'Update the Child Name with PArent Name & Suffix
                        if objExcelSht.Cells(rwNum,3).value = "Child" then objExcelSht.Cells(rwNum,11).value = objExcelSht.cells(rwNum,11).Value & " - " & objExcelSht.cells(rwNum,5).Value 
                    
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

