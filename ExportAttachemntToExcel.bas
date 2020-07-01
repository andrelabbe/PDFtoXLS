Attribute VB_Name = "ExportAttachemntToExcel"
Option Explicit
Public Sub ConvertAttachemntToExcel(Item As Outlook.MailItem)
    On Error Resume Next

    Dim objAtt As Outlook.Attachment
    Dim saveFolder
    Dim saveName
    Dim timeNow
    
    saveFolder = "\\cru-wks-02\ProDisk\GeorgePdfToXls\"

    Dim dspName
    
    For Each objAtt In Item.Attachments
        'below all could be on one line but it id easier to check what is going on
        timeNow = Format(Now, "yyyy-mm-dd-hh-mm-ss")
        saveName = saveFolder & timeNow & ".pdf"
        dspName = UCase(objAtt.DisplayName)
        ' add the check since attachment might include image from email signature, etc...
        If dspName Like "*.PDF" Then
            objAtt.SaveAsFile saveName
            Set objAtt = Nothing
        End If
        
        Next
    ' got the pdf now we convert
    ' convert the pdf to excel
    SaveAsXls


End Sub
Private Sub SaveAsXls()
    Dim pdfPath As String
    Dim excelPath As String

    pdfPath = "c:\ProDisk\GeorgePdfToXls"
    excelPath = "c:\ProDisk\GeorgePdfToXls\xls"

    Dim fsObject As New FileSystemObject
    Dim oneFile
    Dim oneFilePath
    Set oneFilePath = fsObject.GetFolder(pdfPath)

    Dim wordApps
    Dim doc
    Dim clipped
    

    Set wordApps = CreateObject("word.application")
    ' set Word to be hidden, if you want to see it change value to True
    wordApps.Visible = False

    Dim excelApps
    Set excelApps = CreateObject("Excel.Application")
    ' set Excel to be hidden, if you want to see it change value to True
    excelApps.Visible = False

    For Each oneFile In oneFilePath.Files
        Set doc = wordApps.Documents.Open(oneFile.Path, False, Format:="PDF Files")
        Set clipped = doc.Paragraphs(1).Range
        clipped.WholeStory
        Dim oneWorkbook, oneSheet
        Set oneWorkbook = excelApps.workbooks.Add
        Set oneSheet = oneWorkbook.sheets(1)
        clipped.Copy
        
        oneSheet.Paste
        oneWorkbook.SaveAs (excelPath & "\" & Replace(oneFile.Name, ".pdf", "") & ".xlsx")
        oneWorkbook.Close (False)
    Next
    ' close Word
    wordApps.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    wordApps.Quit
    ' close Excel
    excelApps.Quit

End Sub
' only there is needed odd charater(s) in filename
' not use right now
Function RemoveRubbish(sClean As String) As String
'Removing unwanted characters
sClean = Replace(sClean, Chr(9), Chr(32))
sClean = Replace(sClean, Chr(10), Chr(32))
sClean = Replace(sClean, Chr(13), Chr(32))
sClean = Replace(sClean, Chr(16), Chr(32))
sClean = Replace(sClean, ": ", "")
sClean = Replace(sClean, " ", "")
sClean = Replace(sClean, "FWNew", "")
RemoveRubbish = sClean
End Function

