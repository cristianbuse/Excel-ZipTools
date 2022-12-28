Attribute VB_Name = "LibCRC"
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
Option Private Module

#If VBA7 Then
    Private Declare PtrSafe Function RtlComputeCrc32 Lib "ntdll" (ByVal dwInitial As Long, pData As Any, ByVal iLen As Long) As Long
#Else
    Private Declare Function RtlComputeCrc32 Lib "ntdll" (ByVal dwInitial As Long, pData As Any, ByVal iLen As Long) As Long
#End If

Public Function Crc32(ByRef bytes() As Byte) As Long
    #If Mac = 0 Then
        On Error Resume Next
        Crc32 = RtlComputeCrc32(0, bytes(0), UBound(bytes) - LBound(bytes) + 1)
        If Err.Number = 0 Then Exit Function
        On Error GoTo 0
    #End If
    '
    Static crcTable(0 To 255) As Long
    Dim res As Long: res = &HFFFFFFFF
    Dim i As Long
    '
    If crcTable(1) = 0 Then
        Const polynomial As Long = &HEDB88320
        Dim remainder As Long
        Dim j As Long
        '
        For i = LBound(crcTable) To UBound(crcTable)
            remainder = i
            For j = 1 To 8
                remainder = ShiftRight(remainder) Xor (-(remainder And 1&) And polynomial)
            Next j
            crcTable(i) = remainder
        Next i
    End If
    For i = LBound(bytes) To UBound(bytes)
        res = (res And &HFFFFFF00) \ &H100& And &HFFFFFF _
            Xor crcTable(CByte(res And &HFF&) Xor bytes(i))
    Next i
    Crc32 = Not res
End Function
Public Function ShiftRight(ByVal uLong As Long) As Long
    ShiftRight = (uLong And &HFFFFFFFE) \ &H2& And &H7FFFFFFF
End Function
Public Function ShiftRight8(ByVal uLong As Long) As Long
    ShiftRight8 = (uLong And &HFFFFFF00) \ &H100& And &HFFFFFF
End Function
