Option Explicit
' Reads tiff metadata into an Excel worksheet
' Insert a new module into the VBA project, copy this code into it, and then run the Main sub

Const rootFolderPath As String = "C:/path/to/your/tiffs"
Const BUFFER_SIZE As Long = 4096

Private Type IFDTag
  tag As Integer
  Type As Integer
  Count As Long
  ValueOffset As Long
End Type

Sub Main()
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Sheets("TIFF_Metadata")

  ' Clear previous data
  ws.Cells.Clear

  ' Write headers
  ws.Cells(1, 1).Value = "File"
  ws.Cells(1, 2).Value = "Width"
  ws.Cells(1, 3).Value = "Height"
  ws.Cells(1, 4).Value = "DPI X"
  ws.Cells(1, 5).Value = "DPI Y"

  ' Start processing files
  ProcessDirectory fso.GetFolder(rootFolderPath), ws
End Sub

Sub ProcessDirectory(folder As Object, ws As Worksheet)
  Dim file As Object
  Dim subFolder As Object
  Dim rowIndex As Long
  Dim loops As Long

  rowIndex = 2
  loops = 0
  For Each subFolder In folder.SubFolders
    For Each file In subFolder.Files
      loops = loops + 1
      If loops Mod 100 = 0 Then
        DoEvents ' keep excel from freezing
      End If

      If LCase(Right(file.Name, 4)) = ".tif" Then
        ProcessTiffFile file.Path, ws, rowIndex
        rowIndex = rowIndex + 1
      End If
    Next file
  Next subFolder
End Sub

Sub ProcessTiffFile(filePath As String, ws As Worksheet, rowIndex As Long)
  Dim fileNum As Integer
  Dim buffer() As Byte
  Dim isLittleEndian As Boolean
  Dim magicNumber As Integer
  Dim firstIFDOffset As Long
  Dim width As Long
  Dim height As Long
  Dim xResolution As Double
  Dim yResolution As Double
  Dim tags() As IFDTag
  Dim i As Integer

  fileNum = FreeFile
  ws.Cells(rowIndex, 1).value = filePath
  On Error Resume Next
  Open filePath For Binary Access Read As #fileNum

  ' Read TIFF header
  ReDim buffer(7)
  Get #fileNum, , buffer
  isLittleEndian = (Chr(buffer(0)) & Chr(buffer(1)) = "II")
  magicNumber = ReadUInt(buffer, 2, 2, isLittleEndian)
  firstIFDOffset = ReadUInt(buffer, 4, 4, isLittleEndian)

  If magicNumber <> 42 Then
    Close #fileNum
    MsgBox "Not a valid TIFF file: " & filePath
   Exit Sub
  End If

  ' Read IFD entries
  Seek #fileNum, firstIFDOffset + 1
  ReDim buffer(1)
  Get #fileNum, , buffer
  Dim entries As Integer
  entries = ReadUInt(buffer, 0, 2, isLittleEndian)

  ReDim tags(entries - 1)
  ReDim buffer(entries * 12 - 1)
  Get #fileNum, , buffer

  For i = 0 To entries - 1
    With tags(i)
      .tag = ReadUInt(buffer, i * 12, 2, isLittleEndian)
      .Type = ReadUInt(buffer, i * 12 + 2, 2, isLittleEndian)
      .Count = ReadUInt(buffer, i * 12 + 4, 4, isLittleEndian)
      .ValueOffset = ReadUInt(buffer, i * 12 + 8, 4, isLittleEndian)
    End With
  Next i

  ' Extract metadata
  For i = 0 To UBound(tags)
    Select Case tags(i).tag
     Case 256 ' ImageWidth
      width = ReadIFDValue(fileNum, tags(i), isLittleEndian)
     Case 257 ' ImageLength
      height = ReadIFDValue(fileNum, tags(i), isLittleEndian)
     Case 282 ' XResolution
      xResolution = ReadResolution(fileNum, tags(i).ValueOffset, isLittleEndian)
     Case 283 ' YResolution
      yResolution = ReadResolution(fileNum, tags(i).ValueOffset, isLittleEndian)
    End Select
  Next i

  Close #fileNum
  On Error Goto 0

    ' Write metadata To worksheet
    ws.Cells(rowIndex, 2).value = width
    ws.Cells(rowIndex, 3).value = height
    ws.Cells(rowIndex, 4).value = xResolution
    ws.Cells(rowIndex, 5).value = yResolution
End Sub

Function ReadUInt(buffer() As Byte, offset As Long, length As Integer, isLittleEndian As Boolean) As Long
  Dim i As Integer
  Dim result As Long

  result = 0
  If isLittleEndian Then
    For i = length - 1 To 0 Step -1
      result = result * 256 + buffer(offset + i)
    Next i
  Else
    For i = 0 To length - 1
      result = result * 256 + buffer(offset + i)
    Next i
  End If

  ReadUInt = result
End Function

Function ReadIFDValue(fileNum As Integer, tag As IFDTag, isLittleEndian As Boolean) As Long
  Dim buffer(3) As Byte
  Dim value As Long

  ' Check If value is stored directly
  If tag.Count = 1 And tag.Type = 3 Then ' SHORT
    value = tag.ValueOffset And &HFFFF&
  Elseif tag.Count = 1 And tag.Type = 4 Then ' LONG
    value = tag.ValueOffset
  Else
    Seek #fileNum, tag.ValueOffset + 1
    Get #fileNum, , buffer
    value = ReadUInt(buffer, 0, 4, isLittleEndian)
  End If

  ReadIFDValue = value
End Function

Function ReadResolution(fileNum As Integer, offset As Long, isLittleEndian As Boolean) As Double
  Dim buffer(7) As Byte
  Dim numerator As Long
  Dim denominator As Long

  Seek #fileNum, offset + 1
  Get #fileNum, , buffer

  numerator = ReadUInt(buffer, 0, 4, isLittleEndian)
  denominator = ReadUInt(buffer, 4, 4, isLittleEndian)

  ReadResolution = numerator / denominator
End Function
