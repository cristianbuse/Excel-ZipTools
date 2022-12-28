Attribute VB_Name = "LibDeflate"
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

'<length, backward distance> pairs
Private m_lengths(257 To 285) As Long
Private m_lenBits(257 To 285) As Long
Private m_distances(0 To 29) As Long
Private m_distBits(0 To 29) As Long

Private Const MAX_CODE_LEN As Long = 15

Private m_2pwr(0 To MAX_CODE_LEN) As Long 'Single bit mask

Private Enum BlockCompressionType
    noCompression = 0
    fixedHuffmanCodes = 1
    dynamicHuffmanCodes = 2
    reserved = 3
End Enum

Private Type Symbols
    codes() As Long
End Type
Private Type FakeHuffmanTree
    minLenUsed As Long
    maxLenUsed As Long
    lengths(1 To MAX_CODE_LEN) As Symbols
End Type
Private Type Alphabets
    litLTree As FakeHuffmanTree 'Literal/length
    distTree As FakeHuffmanTree 'Distance
End Type

'*******************************************************************************
'DEFLATE Decompress - https://www.ietf.org/rfc/rfc1951.txt
'The decompressed data is returned ByRef to avoid unnecessary copies
'*******************************************************************************
Public Sub Inflate(ByRef bytes() As Byte _
                 , ByVal startIndex As Long _
                 , ByRef outBytes() As Byte _
                 , Optional ByVal outSizeIfKnown As Long)
    Dim isFinalBlock As Boolean
    Dim blockType As BlockCompressionType
    Dim bytePos As Long: bytePos = startIndex
    Dim bitPos As Long:  bitPos = 0
    Dim outPos As Long:  outPos = 0
    Dim outSize As Long: outSize = outSizeIfKnown
    Dim hAlpha As Alphabets
    '
    On Error GoTo Fail
    If m_2pwr(0) = 0 Then InitSupportArrays
    If outSize <= 0 Then outSize = (UBound(bytes) - LBound(bytes) + 1) * 16&
    ReDim outBytes(0 To outSize - 1)
    Do
        isFinalBlock = CBool(Bits(bytes, bytePos, bitPos, 1))
        blockType = Bits(bytes, bytePos, bitPos, 2)
        If blockType = reserved Then
            Err.Raise 5, , "Invalid block type"
        ElseIf blockType = noCompression Then 'rfc1951 - section 3.2.4.
            bytePos = bytePos + Sgn(bitPos) 'Skip remaining bits if any
            bitPos = 0
            '
            Dim bLen As Long:  bLen = bytes(bytePos) Or bytes(bytePos + 1) * &H100&
            Dim bNLen As Long: bNLen = bytes(bytePos + 2) Or bytes(bytePos + 3) * &H100&
            '
            bytePos = bytePos + 4
            If bLen <> (bNLen Xor &HFFFF&) Then
                Err.Raise 5, , "Invalid block length"
            ElseIf bytePos + bLen - 1 > UBound(bytes) Then
                Err.Raise 9, , "Invalid stream size"
            End If
            If bLen > 0 Then
                If outPos + bLen > outSize Then
                    outSize = EnsureCapacity(outBytes, outPos + bLen)
                End If
                MemCopy VarPtr(outBytes(outPos)), VarPtr(bytes(bytePos)), bLen
                bytePos = bytePos + bLen
                outPos = outPos + bLen
            End If
        Else
            If blockType = dynamicHuffmanCodes Then
                ReadDynamicAlphabets bytes, bytePos, bitPos, hAlpha
            ElseIf blockType = fixedHuffmanCodes Then
                CopyFixedAlphabets hAlpha
            Else
                Err.Raise 5, , "Block type bits read failure"
            End If
            '
            Dim hSymbol As Long
            Do
                hSymbol = DecodeSymbol(bytes, bytePos, bitPos, hAlpha.litLTree)
                If hSymbol = 256& Then Exit Do
                If hSymbol < 256& Then
                    If outPos + 1 > outSize Then
                        outSize = EnsureCapacity(outBytes, outPos + 1)
                    End If
                    outBytes(outPos) = hSymbol 'Literal
                    outPos = outPos + 1
                Else '> 256
                    Dim byteLen As Long: byteLen = m_lengths(hSymbol)
                    Dim lenBits As Long: lenBits = m_lenBits(hSymbol)
                    '
                    If CBool(lenBits) Then
                        byteLen = byteLen + Bits(bytes, bytePos, bitPos, lenBits)
                    End If
                    hSymbol = DecodeSymbol(bytes, bytePos, bitPos, hAlpha.distTree)
                    '
                    Dim backDist As Long: backDist = m_distances(hSymbol)
                    Dim distBits As Long: distBits = m_distBits(hSymbol)
                    '
                    If CBool(distBits) Then
                        backDist = backDist + Bits(bytes, bytePos, bitPos, distBits)
                    End If
                    '
                    If outPos + byteLen > outSize Then
                        outSize = EnsureCapacity(outBytes, outPos + byteLen)
                    End If
                    For outPos = outPos To outPos + byteLen - 1
                         outBytes(outPos) = outBytes(outPos - backDist)
                    Next outPos
                End If
            Loop
        End If
    Loop Until isFinalBlock
    If outPos < outSize Then ReDim Preserve outBytes(0 To outPos - 1)
Exit Sub
Fail:
    Err.Raise Err.Number, "Inflate", Err.Description
End Sub
Private Function EnsureCapacity(ByRef buffer() As Byte _
                              , ByVal neededCapacity As Long) As Long
    Const MAX_CAPACITY As Long = &H7FFFFFFF
    Dim lowBound As Long:    lowBound = LBound(buffer)
    Dim oldCapacity As Long: oldCapacity = UBound(buffer) - lowBound + 1
    '
    If neededCapacity > oldCapacity Then
        Dim newCapacity As Long
        If CDbl(neededCapacity) * 2# > CDbl(MAX_CAPACITY) Then
            newCapacity = MAX_CAPACITY
        Else
            newCapacity = neededCapacity * 2&
        End If
        ReDim Preserve buffer(lowBound To lowBound + newCapacity - 1)
        EnsureCapacity = newCapacity
    Else
        EnsureCapacity = oldCapacity
    End If
End Function

'*******************************************************************************
'Initialize the <length, backward distance> pairs and bit masks
'*******************************************************************************
Private Sub InitSupportArrays()
    'https://www.ietf.org/rfc/rfc1951.txt - section 3.2.5.
    AddLongs m_lenBits, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0
    AddLongs m_lengths, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258
    AddLongs m_distances, 1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577
    AddLongs m_distBits, 0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13
    '
    Dim i As Long: m_2pwr(0) = 1
    For i = 1 To UBound(m_2pwr)
        m_2pwr(i) = m_2pwr(i - 1) * 2 'Single bit mask e.g. 00010000
    Next i
End Sub
Private Sub AddLongs(ByRef arr() As Long, ParamArray longs() As Variant)
    Dim i As Long
    Dim j As Long: j = LBound(arr)
    '
    For i = 0 To UBound(longs)
        arr(j) = longs(i)
        j = j + 1
    Next i
End Sub

'*******************************************************************************
'Returns up to 15 bits from up to 3 bytes
'*******************************************************************************
Private Function Bits(ByRef bytes() As Byte _
                    , ByRef bytePos As Long _
                    , ByRef bitPos As Long _
                    , ByVal bitsCount As Long) As Long
    Dim newBitPos As Long:  newBitPos = bitPos + bitsCount
    Dim extraBytes As Long: extraBytes = newBitPos \ 8&
    Dim useB1 As Long:      useB1 = Sgn(extraBytes)
    Dim useB2 As Long:      useB2 = (extraBytes \ 2&)
    '
    Bits = (bytes(bytePos) Or useB1 * bytes(bytePos + useB1) * &H100& _
                           Or useB2 * bytes(bytePos + useB2 * 2&) * &H10000 _
           ) \ m_2pwr(bitPos) And (m_2pwr(bitsCount) - 1)
    '
    bytePos = bytePos + extraBytes
    bitPos = newBitPos And &H7&
End Function

'*******************************************************************************
'Read representation of Huffman code trees for dynamic compression
'https://www.ietf.org/rfc/rfc1951.txt - section 3.2.7.
'*******************************************************************************
Private Sub ReadDynamicAlphabets(ByRef bytes() As Byte _
                               , ByRef bytePos As Long _
                               , ByRef bitPos As Long _
                               , ByRef outAlphabets As Alphabets)
    Dim hLit As Long:  hLit = Bits(bytes, bytePos, bitPos, 5) + 257
    Dim hDist As Long: hDist = Bits(bytes, bytePos, bitPos, 5) + 1
    Dim hcLen As Long: hcLen = Bits(bytes, bytePos, bitPos, 4) + 4
    '
    Static codeLenOrder(0 To 18) As Long
    Dim codeLenTree As FakeHuffmanTree
    Dim i As Long
    Dim codeLengths() As Long: ReDim codeLengths(0 To 18)
    '
    'Read code lengths in correct order as per specification
    If codeLenOrder(0) = 0 Then
        AddLongs codeLenOrder, 16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15
    End If
    For i = 0 To hcLen - 1
        codeLengths(codeLenOrder(i)) = Bits(bytes, bytePos, bitPos, 3)
    Next i
    CreateFakeTree codeLengths, 0, 18, 7, codeLenTree
    '
    Dim codeLen As Long
    Dim maxBits As Long
    Dim timesToCopy As Long
    Dim sequenceLen As Long: sequenceLen = hLit + hDist
    '
    i = 0
    ReDim codeLengths(0 To sequenceLen - 1)
    Do
        codeLen = DecodeSymbol(bytes, bytePos, bitPos, codeLenTree)
        If codeLen < 16 Then
            codeLengths(i) = codeLen
            i = i + 1
            If codeLen > maxBits Then maxBits = codeLen
        ElseIf codeLen = 16 Then
            timesToCopy = Bits(bytes, bytePos, bitPos, 2) + 3
            For i = i To i + timesToCopy - 1
                codeLengths(i) = codeLengths(i - 1)
            Next i
        Else 'Cover both 17 and 18 without any If statements
            Dim Extra As Long: Extra = 4 * (codeLen - 17)
            i = i + Bits(bytes, bytePos, bitPos, 3 + Extra) + 3 + Extra * 2
        End If
    Loop Until i > sequenceLen - 1
    '
    CreateFakeTree codeLengths, 0, hLit - 1, maxBits, outAlphabets.litLTree
    CreateFakeTree codeLengths, hLit, hDist - 1, maxBits, outAlphabets.distTree
End Sub

'*******************************************************************************
'Generate Huffman codes from the bit lengths of the codes (alphabetical order)
'https://www.ietf.org/rfc/rfc1951.txt - section 3.2.2.
'*******************************************************************************
Private Sub CreateFakeTree(ByRef inCodeLengths() As Long _
                         , ByVal inStartIndex As Long _
                         , ByVal maxCode As Long _
                         , ByVal maxBits As Long _
                         , ByRef outTree As FakeHuffmanTree)
    Dim blCount(0 To MAX_CODE_LEN) As Long
    Dim nextCode(0 To MAX_CODE_LEN) As Long
    Dim code As Long
    Dim i As Long
    Dim cLen As Long
    '
    If outTree.minLenUsed > 0 Then Erase outTree.lengths
    outTree.minLenUsed = MAX_CODE_LEN + 1
    outTree.maxLenUsed = 0
    '
    '1) Count the number of codes for each code length
    For i = inStartIndex To maxCode + inStartIndex
        cLen = inCodeLengths(i)
        blCount(cLen) = blCount(cLen) + 1
        With outTree
            If .minLenUsed * Sgn(cLen) > cLen Then .minLenUsed = cLen
            If .maxLenUsed < cLen Then .maxLenUsed = cLen
        End With
    Next i
    '
    'Extra step for fake tree
    For i = outTree.minLenUsed To outTree.maxLenUsed
        ReDim outTree.lengths(i).codes(0 To m_2pwr(i) - 1)
    Next i
    '
    '2) Find the numerical value of the smallest code for each code length
    code = 0
    blCount(0) = 0
    For i = 1 To maxBits
        code = (code + blCount(i - 1)) * 2 '<< Shift Left (*)
        nextCode(i) = code
    Next i
    '
    '3) Assign numerical values to all codes
    For i = 0 To maxCode
        cLen = inCodeLengths(i + inStartIndex)
        If CBool(cLen) Then 'Assign value to used codes only
            code = nextCode(cLen)
            outTree.lengths(cLen).codes(code) = i + 1 'Non-zero - used as Bool
            nextCode(cLen) = code + 1
        End If
    Next i
End Sub

'*******************************************************************************
'Decode a single symbol using the fake Huffman tree (array of arrays)
'*******************************************************************************
Private Function DecodeSymbol(ByRef bytes() As Byte _
                            , ByRef bytePos As Long _
                            , ByRef bitPos As Long _
                            , ByRef hTree As FakeHuffmanTree) As Long
    Dim code As Long
    Dim hSymbol As Long
    Dim i As Long
    '
    For i = 1 To hTree.minLenUsed - 1
        code = code * 2& + Sgn(bytes(bytePos) And m_2pwr(bitPos))
        bitPos = (bitPos + 1) And &H7&
        bytePos = bytePos + 1 - Sgn(bitPos)
    Next i
    For i = hTree.minLenUsed To hTree.maxLenUsed
        code = code * 2& + Sgn(bytes(bytePos) And m_2pwr(bitPos))
        hSymbol = hTree.lengths(i).codes(code)
        '
        bitPos = (bitPos + 1) And &H7&
        bytePos = bytePos + 1 - Sgn(bitPos)
        '
        If CBool(hSymbol) Then
            DecodeSymbol = hSymbol - 1
            Exit Function
        End If
    Next i
    Err.Raise 5, "DecodeSymbol", "Invalid Huffman code"
End Function

'*******************************************************************************
'Generate Fixed Huffman codes
'https://www.ietf.org/rfc/rfc1951.txt - section 3.2.6.
'*******************************************************************************
Private Sub CopyFixedAlphabets(ByRef outAlphabets As Alphabets)
    Static fixedAlphabets As Alphabets
    '
    If fixedAlphabets.distTree.minLenUsed = 0 Then
        Dim i As Long
        With fixedAlphabets.litLTree.lengths(8)
            ReDim .codes(0 To 199)
            For i = 0 To 143
                .codes(i + 48) = i + 1 'Non-zero - used as Bool when decoding
            Next i
            For i = 280 To 287
                .codes(i + 192 - 280) = i + 1
            Next i
        End With
        With fixedAlphabets.litLTree.lengths(9)
            ReDim .codes(0 To 511)
            For i = 144 To 255
                .codes(i + 256) = i + 1
            Next i
        End With
        With fixedAlphabets.litLTree.lengths(7)
            ReDim .codes(0 To 23)
            For i = 256 To 279
                .codes(i - 256) = i + 1
            Next i
        End With
        With fixedAlphabets.distTree.lengths(5)
            ReDim .codes(0 To 29)
            For i = 0 To 29
                .codes(i) = i + 1
            Next i
        End With
        fixedAlphabets.distTree.minLenUsed = 5
        fixedAlphabets.distTree.maxLenUsed = 5
        fixedAlphabets.litLTree.minLenUsed = 7
        fixedAlphabets.litLTree.maxLenUsed = 9
    End If
    outAlphabets = fixedAlphabets
End Sub
