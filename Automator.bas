Attribute VB_Name = "Module1"
Sub AutomateParser()

    Dim F As String
    Dim pauseRun As Boolean
    Dim roww As Long
    Dim ws As Worksheet
    Dim wbP As Workbook
    Dim wbA As Workbook
    Dim wbF As Workbook
    Dim Msg, Style, Response
    
    roww = 0
    pauseRun = False
    
    Dim FileLocSpec As String
    FileLocSpec = "C:\Users\Fernando\Dropbox\Data\Psion Data\OldParsed\*.xlsm"
    F = Dir(FileLocSpec)
    
    Do Until F = ""
        roww = roww + 1
        Cells(roww, 1).Value = F
        F = Dir
    Loop

    Set r = Range("A1")
    
    While r.Value <> "" And pauseRun <> True
        Set wbP = Workbooks.Open("C:\Users\Fernando\Dropbox\Data\Parser\CamposPhDParser.xlsm")
        Set wbF = Workbooks.Open("C:\Users\Fernando\Dropbox\Data\Psion Data\OldParsed\" & r.Value)
        
        Application.DisplayAlerts = False
        
        wbF.Worksheets(2).Cells.Copy
        
        Set ws = wbP.Worksheets(2)
        ws.Range("A1").PasteSpecial
        wbF.Close
        
        Application.Run "CamposPhDParser.xlsm!ProcessData"
        
        wbP.SaveAs ("C:\Users\Fernando\Dropbox\Data\Psion Data\NewParsed\" & Left(r.Value, Len(r.Value) - 4) & "xlsm")
        wbP.Close
        
        Set r = r.Offset(1, 0)
        
        
'        Msg = "Do you want to continue ?"
'        Style = vbYesNo + vbCritical + vbDefaultButton2
'        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
'        If Response = vbYes Then
'            pauseRun = False
'        Else
'            pauseRun = True
'        End If
        
    Wend
End Sub
