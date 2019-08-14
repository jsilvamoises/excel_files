Attribute VB_Name = "MDL_FILES"
Sub abrirArquivo()
 
MsgBox "Selecione o arquivo txt", vbOKOnly, "Seleção de Arquivo"
 
'ABRIR ARQUIVO
arquivo = " "
 
Dim fd As FileDialog
 
Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
Dim arquivo_temp As Variant
 
With fd
   .AllowMultiSelect = True
   If .Show = -1 Then
      For Each arquivo_temp In .SelectedItems
         arquivo = arquivo_temp
      Next arquivo_temp
   End If
End With
 
Set fd = Nothing
 
'Abaixo é um código para ajustar as colunas do .txt para o excel que varia conforme cada tipo de arquivo
'Para você saber os seus parametros ideais, uma dica é criar uma macro e abrir um .txt e definir as colunas
'E depois ver o código que foi gerado.
 
Workbooks.OpenText arquivo _
, Origin:=xlWindows, StartRow:=1, DataType:=xlFixedWidth, TextQualifier:=xlDoubleQuote _
, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=True, Comma:=False _
, Space:=False, Other:=False, FieldInfo:=Array(Array(0, 1), Array(38, 1), _
Array(91, 1)), TrailingMinusNumbers:=True
 
'Ajuste automatico de coluna do excel
Columns("B:B").EntireColumn.AutoFit
Columns("A:A").EntireColumn.AutoFit
 
End Sub
 
Sub WriteFile(ByRef conteudo As String, open_ As Boolean)
Dim Usuario As String
Usuario = Environ("username")
Dim path As String
path = "C:\" & Usuario & "\@QUERIES\"

If Dir(path, vbDirectory) = "" Then
 Shell ("cmd /c mkdir """ & path & """")
End If


ThisFile = path & "query.sql"
On Error Resume Next
    Kill (ThisFile)
On Error GoTo 0

Open ThisFile For Output As #1
Print #1, conteudo
Close #1
If open_ Then
  OpenInNotepad ThisFile
End If

End Sub
Sub TesteWriteFile()
Dim x As String
x = Environ("homedir")
WriteFile "Meu Jesuskkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkk", True
End Sub

' VBA For Creating A Text File
Sub TextFile_Create()
'PURPOSE: Create A New Text File
'SOURCE: www.TheSpreadsheetGuru.com

Dim TextFile As Integer
Dim FilePath As String

'What is the file path and name for the new text file?
  FilePath = "C:\Users\Usuario\Desktop\MyFile.txt"

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file
  Open FilePath For Output As TextFile

'Write some lines of text
  Print #TextFile, "Hello Everyone!"
  Print #TextFile, "I created this file with VBA."
  Print #TextFile, "Goodbye"
  
'Save & Close Text File
  Close TextFile

End Sub

' VBA For Extracting All The Text From A Text File
Sub TextFile_PullData()
'PURPOSE: Send All Data From Text File To A String Variable
'SOURCE: www.TheSpreadsheetGuru.com

Dim TextFile As Integer
Dim FilePath As String
Dim FileContent As String

'File Path of Text File
  FilePath = "C:\Users\" & Environ("username") & "\Desktop\MyFile.txt"

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file
  Open FilePath For Input As TextFile

'Store file content inside a variable
  FileContent = Input(LOF(TextFile), TextFile)

'Report Out Text File Contents
  MsgBox FileContent

'Close Text File
  Close TextFile

End Sub

' VBA For Modifying A Text File (With Find/Replace)

Sub TextFile_FindReplace()
'PURPOSE: Modify Contents of a text file using Find/Replace
'SOURCE: www.TheSpreadsheetGuru.com

Dim TextFile As Integer
Dim FilePath As String
Dim FileContent As String

'File Path of Text File
  FilePath = "C:\Users\" & Environ("username") & "\Desktop\MyFile.txt"

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file in a Read State
  Open FilePath For Input As TextFile

'Store file content inside a variable
  FileContent = Input(LOF(TextFile), TextFile)

'Clost Text File
  Close TextFile
  
'Find/Replace
  FileContent = Replace(FileContent, "Goodbye", "Cheers")

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file in a Write State
  Open FilePath For Output As TextFile
  
'Write New Text data to file
  Print #TextFile, FileContent

'Close Text File
  Close TextFile

End Sub


Sub TextFile_Create_Append()
'PURPOSE: Add More Text To The End Of A Text File
'SOURCE: www.TheSpreadsheetGuru.com

Dim TextFile As Integer
Dim FilePath As String

'What is the file path and name for the new text file?
  FilePath = "C:\Users\" & Environ("username") & "\Desktop\MyFile.txt"

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file
  Open FilePath For Append As TextFile

'Write some lines of text
  Print #TextFile, "Sincerely,"
  Print #TextFile, ""
  Print #TextFile, "Chris"
  
'Save & Close Text File
  Close TextFile

End Sub

' VBA For Fill Array With Delimited Data From Text File

Sub DelimitedTextFileToArray()
'PURPOSE: Load an Array variable with data from a delimited text file
'SOURCE: www.TheSpreadsheetGuru.com

Dim Delimiter As String
Dim TextFile As Integer
Dim FilePath As String
Dim FileContent As String
Dim LineArray() As String
Dim DataArray() As String
Dim TempArray() As String
Dim rw As Long, col As Long
Dim x, y As Variant

'Inputs
  Delimiter = ";"
  FilePath = "C:\Users\" & Environ("username") & "\Desktop\MyFile.txt"
  rw = 0
  
'Open the text file in a Read State
  TextFile = FreeFile
  Open FilePath For Input As TextFile
  
'Store file content inside a variable
  FileContent = Input(LOF(TextFile), TextFile)

'Close Text File
  Close TextFile
  
'Separate Out lines of data
  LineArray() = Split(FileContent, vbCrLf)

'Read Data into an Array Variable
  For x = LBound(LineArray) To UBound(LineArray)
    If Len(Trim(LineArray(x))) <> 0 Then
      'Split up line of text by delimiter
        TempArray = Split(LineArray(x), Delimiter)
      
      'Determine how many columns are needed
        col = UBound(TempArray)
      
      'Re-Adjust Array boundaries
        ReDim Preserve DataArray(col, rw)
      
      'Load line of data into Array variable
        For y = LBound(TempArray) To UBound(TempArray)
          DataArray(y, rw) = TempArray(y)
        Next y
    End If
    
    'Next line
      rw = rw + 1
    
  Next x

End Sub


' VBA For Deleting A Text File
Sub TextFile_Delete()
'PURPOSE: Delete a Text File from your computer
'SOURCE: www.TheSpreadsheetGuru.com

Dim FilePath As String

'File Path of Text File
  FilePath = "C:\Users\" & Environ("username") & "\Desktop\MyFile.txt"

'Delete File
  Kill FilePath

End Sub


Sub OpenInNotepad(filename As Variant)
Dim MyTxtFile
On Error GoTo NOTE_PAD
    MyTxtFile = Shell("C:\Program Files (x86)\Notepad++\notepad++.exe " & filename, 1)

NOTE_PAD:
Debug.Print "Erro: >> " & Err.Description
    If Err.Number > 0 Then
    MyTxtFile = Shell("C:\WINDOWS\notepad.exe " & filename, 1)
    End If
    
    


End Sub


Sub FileCopy_(file As String, destino As String, tag As String)
    'fileCopy file, destino
fileCopy file, destino
    ' listfiles "c:\"
End Sub

Sub testeFileCopy()
fileCopy "C:\Users\Usuario\Downloads\lotomania.xlsx", "C:\Users\Usuario\Dropbox\#JavaProjects\lotomania.xlsx"

End Sub

Function listfiles(ByVal sPath As String)

    Dim vaArray     As Variant
    Dim i           As Integer
    Dim oFile       As Object
    Dim oFSO        As Object
    Dim oFolder     As Object
    Dim oFiles      As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files

    If oFiles.Count = 0 Then Exit Function

    ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = oFile.Name
        Debug.Print oFile.Name
        i = i + 1
    Next

    listfiles = vaArray

End Function
