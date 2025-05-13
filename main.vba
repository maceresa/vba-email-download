Option Explicit

Function Compras_EncontrarCarpetaCompulsa(compulsa As String, rootPath As String) As String

Dim tt As String
Dim ts As Variant
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim MyFolder As Object
Dim MySubFolder As Object

'    Dim fso As New FileSystemObject
'    Dim MyFolder As Folder
'    Dim mySubFolder As Folder

Set MyFolder = FSO.GetFolder(rootPath)
 For Each MySubFolder In MyFolder.subfolders
   tt = MySubFolder.Name
   ts = InStr(1, tt, compulsa, 1)
   If ts > 0 Then
    'MsgBox (tt)
    Compras_EncontrarCarpetaCompulsa = tt
    'Shell "Explorer.exe " & MySubFolder, vbNormalFocus
    Exit Function
   End If
 '  Debug.Print MySubFolder.Name
 Next MySubFolder
 
End Function

Public Sub Compras_GuardarAdjunto(Item As Outlook.MailItem)

Dim oAttachment As Outlook.Attachment
Dim sSaveFolder As String
Dim sFileName As String
Dim objDoc As Object
Dim proveedor As String

Const olMsg As Long = 3

Dim m As MailItem
Dim savePath As String
Dim anio As String
Dim raiz As String
Dim compulsa As String
Dim carpeta As String
Dim compulsaConCero As String
Dim CAB As String
Dim CABConCero As String
Dim subject As String
Dim largo As Integer
Dim aCortar As Integer
Dim nombreAdjunto As String
Dim extension As String

anio = Year(Now())
raiz = "\\ncparfs\UGLs\08\Sede\DeptoAdministrativo\Compras\" + anio + "\Compulsas\"
    
If TypeName(Item) <> "MailItem" Then Exit Sub

    Set m = Item
    
    If InStr(m.To, "Presupuestos_ugl08") = 0 Then Exit Sub '"Presupuestos_ugl08" Then Exit Sub
        m.subject = Replace(m.subject, "-", " ")
        m.subject = Replace(m.subject, "_", " ")
        m.subject = Replace(m.subject, "Nº", " ")
        m.subject = Replace(m.subject, "N°", " ")
        m.subject = Replace(m.subject, "Nro", " ")
        m.subject = Replace(m.subject, "   ", " ")
        m.subject = Replace(m.subject, "  ", " ")
        
        If InStr(m.subject, "CA ") <> 0 Then
            subject = Left(Mid(m.subject, InStr(m.subject, "CA ")), 7)
            compulsa = subject
            largo = Len(subject)
            aCortar = largo - 3
            If Not IsNumeric(Right(subject, 1)) Or Len(subject) < 7 Then
                compulsa = "CA 0" & Left(Right(subject, aCortar), 3)
            End If
        End If
        If InStr(m.subject, "CAB ") <> 0 Then
            subject = Left(Mid(m.subject, InStr(m.subject, "CAB ")), 8)
            compulsa = subject
            largo = Len(subject)
            aCortar = largo - 4
            If Not IsNumeric(Right(subject, 1)) Or Len(subject) < 8 Then
                compulsa = "CA 0" & Left(Right(subject, aCortar), 3)
            Else
                compulsa = "CA 0" & Left(Right(subject, aCortar - 1), 3)
            End If
        End If
        If InStr(UCase(m.subject), "COMPULSA ABREVIADA ") <> 0 Then
            subject = Left(Mid(m.subject, InStr(UCase(m.subject), "COMPULSA ABREVIADA ")), 23)
            compulsa = subject
            largo = Len(subject)
            aCortar = largo - 19
            If Not IsNumeric(Right(subject, 1)) Or Len(subject) < 23 Then
                compulsa = "CA 0" & Left(Right(subject, aCortar), 3)
            Else
                compulsa = "CA 0" & Left(Right(subject, aCortar - 1), 3)
            End If
        End If
        If InStr(m.subject, "CD ") <> 0 Then
            subject = Left(Mid(m.subject, InStr(m.subject, "CD ")), 7)
            compulsa = subject
            largo = Len(subject)
            aCortar = largo - 3
            If Not IsNumeric(Right(subject, 1)) Or Len(subject) < 7 Then
                compulsa = "CA 0" & Left(Right(subject, aCortar), 3)
            End If
        End If
        
        If carpeta = "" And compulsa <> "" Then
            carpeta = Compras_EncontrarCarpetaCompulsa(compulsa, raiz)
        End If
        
        If carpeta = "" Then Exit Sub
        
        'proveedor = Mid(m.SenderEmailAddress, InStr(m.SenderName, "@"))
        proveedor = m.Sender.Name '& "-" & Right(proveedor, 1)
            
        savePath = raiz & carpeta & "\"
        sFileName = savePath & Format(Now, "YYYYMMDD") & "-" & Format(Now, "hhmmss") & "-" & proveedor & "-Mail.pdf"
        
        'objDoc.ExportAsFixedFormat OutputFileName:=sFileName, _
                ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=0, To:=0, _
                Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=True, UseISO19005_1:=False

            ' And close once saved on disk
            'objDoc.Close (False)

'-----------------------------------------------
    Dim FSO As Object, TmpFolder As Object
    Dim tmpFileName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    tmpFileName = FSO.GetSpecialFolder(2)
    
    'ReplaceCharsForFileName sName, "-"
    tmpFileName = tmpFileName & "\" & proveedor & Format(Now, "YYYYMMDD") & "-" & Format(Now, "hhmmss") & ".mht"
    
    Item.SaveAs tmpFileName, olMHTML
    
    'Create a Word object
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Set wrdApp = CreateObject("Word.Application")
    
    Set wrdDoc = wrdApp.Documents.Open(FileName:=tmpFileName, Visible:=True)
  
    'Dim WshShell As Object
    Dim SpecialPath As String
    Dim strToSaveAs As String
    'Set WshShell = CreateObject("WScript.Shell")
    'MyDocs = WshShell.SpecialFolders(16)
       
    strToSaveAs = sFileName 'MyDocs & "\" & sName & ".pdf"
    ' check for duplicate filenames
    ' if matched, add the current time to the file name
    'If FSO.FileExists(strToSaveAs) Then
        'sName = sName & Format(Now, "hhmmss")
        'strToSaveAs = sFileName 'MyDocs & "\" & sName & ".pdf"
    'End If
  
    wrdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:= _
    strToSaveAs, ExportFormat:=wdExportFormatPDF, _
    OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
    Range:=wdExportAllDocument, From:=0, To:=0, Item:= _
    wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
    CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
    BitmapMissingFonts:=True, UseISO19005_1:=False
             
    wrdDoc.Close
    wrdApp.Quit
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    'Set WshShell = Nothing
'-----------------------------------------------
    'SaveMessageAsPDF (Item)
    'm.SaveAs savePath, olMsg
    
    For Each oAttachment In Item.Attachments
        extension = Right(Format(Now, "YYYYMMDD") & "-" & Format(Now, "hhmmss") & "-" & proveedor & "-" & oAttachment.DisplayName, 4)
        nombreAdjunto = Left(Format(Now, "YYYYMMDD") & "-" & Format(Now, "hhmmss") & "-" & proveedor & "-" & oAttachment.DisplayName, 55) & extension
        oAttachment.SaveAsFile savePath & nombreAdjunto
    Next
    
    
End Sub
