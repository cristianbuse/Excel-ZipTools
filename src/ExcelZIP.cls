VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelZIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' Excel VBA Zip Tools
''' ----------------------------------------------
''' https://github.com/cristianbuse/Excel-ZipTools
''' ----------------------------------------------
''' MIT License
'''
''' Copyright (c) 2022 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit

Private m_bytes() As Byte
Private m_sBytes As String

Private Type FileHeader
    versionNeeded As Integer
    generalBitFlag As Integer
    compressionMethod As Integer
    lastModificationTime As Integer
    lastModificationDate As Integer
    crc32uncompressed As Long
    CompressedSize As Long
    uncompressedSize As Long
    nameLength As Integer
    extraFieldLength As Integer
    fileName As String
    dataStart As Long
    uncompressedData() As Byte
End Type

Private m_fileHeaders() As FileHeader
Private m_fileIndex As Collection

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Class Events
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
    ClearState
End Sub
Private Sub ClearState()
    ReDim m_fileHeaders(1 To 1)
    Set m_fileIndex = New Collection
End Sub

'*******************************************************************************
'Initialize the archive
'*******************************************************************************
Public Sub Init(ByVal zipPath As String)
    ClearState
    ReadBytes zipPath, m_bytes
    m_sBytes = m_bytes
    
    If LenB(m_sBytes) = 0 Or LenB(m_sBytes) > 2 ^ 31 - 1 Then Exit Sub
    '
    Const fileSignature As Long = &H4034B50
    Dim i As Long: i = 0
    Dim j As Long
    '
    Do
        If ReadLong(i) <> fileSignature Then Exit Do
        '
        j = j + 1
        If UBound(m_fileHeaders) < j Then
            ReDim Preserve m_fileHeaders(1 To UBound(m_fileHeaders) * 2)
        End If
        '
        With m_fileHeaders(j)
            .versionNeeded = ReadInt(i + 4)
            .generalBitFlag = ReadInt(i + 6)
            .compressionMethod = ReadInt(i + 8)
            .lastModificationTime = ReadInt(i + 10)
            .lastModificationDate = ReadInt(i + 12)
            .crc32uncompressed = ReadLong(i + 14)
            .CompressedSize = ReadLong(i + 18)
            .uncompressedSize = ReadLong(i + 22)
            .nameLength = ReadInt(i + 26)
            .extraFieldLength = ReadInt(i + 28)
            .fileName = StrConv(MidB$(m_sBytes, i + 31, .nameLength), vbUnicode)
            .dataStart = i + 30 + .nameLength + .extraFieldLength
            .uncompressedData = vbNullString
            i = .dataStart + .CompressedSize
            m_fileIndex.Add j, .fileName
        End With
    Loop
End Sub

'*******************************************************************************
'Utilities for reading integers out of the array of bytes
'*******************************************************************************
Private Function ReadInt(ByVal i As Long) As Integer
    Dim hiByte As Byte: hiByte = m_bytes(i + 1)
    ReadInt = m_bytes(i) Or (hiByte And &H7F) * &H100 Or (&H8000 * -CBool(hiByte And &H80))
End Function
Private Function ReadLong(ByVal i As Long) As Long
    Dim hiByte As Byte: hiByte = m_bytes(i + 3)
    ReadLong = m_bytes(i) Or m_bytes(i + 1) * &H100& Or m_bytes(i + 2) * &H10000 _
            Or (hiByte And &H7F) * &H1000000 Or (&H80000000 * -CBool(hiByte And &H80))
End Function

'*******************************************************************************
'Utility - Reads a file into an array of Bytes
'*******************************************************************************
Private Sub ReadBytes(ByVal filePath As String, ByRef result() As Byte)
    If Not IsFile(filePath) Then
        Erase result
        Exit Sub
    End If
    '
    Dim fileNumber As Long: fileNumber = FreeFile()
    '
    Open filePath For Binary Access Read As #fileNumber
    Dim Size As Long: Size = LOF(fileNumber)
    If Size > 0 Then
        ReDim result(0 To Size - 1)
        Get fileNumber, , result
    Else
        Erase result
    End If
    Close #fileNumber
End Sub
Private Function IsFile(ByVal filePath As String) As Boolean
    On Error Resume Next
    IsFile = ((GetAttr(filePath) And vbDirectory) <> vbDirectory)
    On Error GoTo 0
End Function

'===============================================================================
'Number of files in the archive
'===============================================================================
Public Function Count() As Long
    Count = m_fileIndex.Count
End Function

'*******************************************************************************
'Returns an array of all file names in the archive or a zero-length array
'*******************************************************************************
Public Function GetFileNames() As String()
    If m_fileIndex.Count = 0 Then
        GetFileNames = Split(vbNullString)
        Exit Function
    End If
    '
    Dim i As Long
    Dim res() As String
    '
    ReDim res(1 To m_fileIndex.Count)
    For i = 1 To m_fileIndex.Count
        res(i) = m_fileHeaders(i).fileName
    Next i
    GetFileNames = res
End Function

'*******************************************************************************
'Checks if a file exists within the archive
'*******************************************************************************
Public Function HasFile(ByRef fileName As String) As Boolean
    On Error Resume Next
    m_fileIndex.Item fileName
    HasFile = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Returns the uncompressed data for the required file
'*******************************************************************************
Public Sub ReadData(ByVal fileName As String, ByRef result() As Byte)
    If Not HasFile(fileName) Then
        Err.Raise 5, TypeName(Me) & ".ReadData", "File does not exist"
        Exit Sub
    End If
    '
    With m_fileHeaders(m_fileIndex(fileName))
        If UBound(.uncompressedData) - LBound(.uncompressedData) + 1 = 0 Then
            Inflate m_bytes, .dataStart, .uncompressedData, .uncompressedSize
            If Crc32(.uncompressedData) <> .crc32uncompressed Then
                Err.Raise 5, TypeName(Me) & ".ReadData", "Corrupted data"
            End If
        End If
        result = .uncompressedData
    End With
End Sub
