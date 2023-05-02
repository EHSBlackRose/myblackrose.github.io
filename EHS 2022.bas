Attribute VB_Name = "Module1"
Option Explicit

Const FOLDER_SAVED As String = "B:\2022\REG_ASPAL_TAMBUN SELATAN_NO_U_PL_"
Const SOURCE_FILE_PATH As String = "B:\_PENGADAAN\2022\FIX\FIX 2022.xlsx"

Sub MailMergeToIndPDF()
Dim MainDoc As Document, TargetDoc As Document
Dim dbPath As String
Dim recordNumber As Long, totalRecord As Long

Set MainDoc = ActiveDocument
With MainDoc.MailMerge
    
        '// if you want to specify your data, insert a WHERE clause in the SQL statement
        .OpenDataSource Name:=SOURCE_FILE_PATH, sqlstatement:="SELECT * FROM [FIX PL PERMUKIMAN 2022 Kontrak $]"
            
        totalRecord = .DataSource.RecordCount

        For recordNumber = 637 To 637
        
            With .DataSource
                .ActiveRecord = recordNumber
                .FirstRecord = recordNumber
                .LastRecord = recordNumber
            End With
            
            .Destination = wdSendToNewDocument
            .Execute False
            
            Set TargetDoc = ActiveDocument

            TargetDoc.SaveAs2 FOLDER_SAVED & .DataSource.DataFields("LOKASI_PEKERJAAN").Value & ".docx", wdFormatDocumentDefault
            TargetDoc.ExportAsFixedFormat FOLDER_SAVED & .DataSource.DataFields("NO_URUT_PL").Value & ".pdf", exportformat:=wdExportFormatPDF
            TargetDoc.Close False
            Set TargetDoc = Nothing
        Next recordNumber
End With
    On Error Resume Next
    Kill FOLDER_SAVED & "*.docx"
    On Error GoTo 0
Set MainDoc = Nothing
End Sub

