'Attribute VB_Name = "TypeID"
' TypeID Module - Generate TypeID identifiers for Excel
' TypeID = prefix + "_" + UUIDv7 encoded in base32 (Crockford)
' Complete version with anti-regeneration tools
' IMPORTANT: TypeID encodes 130 bits (2 zero bits + 128 bits UUID)

Option Explicit

' Crockford Base32 alphabet (without I, L, O, U to avoid confusion)
Private Const BASE32_ALPHABET As String = "0123456789abcdefghjkmnpqrstvwxyz"

' ============================================================================
' MAIN FUNCTIONS
' ============================================================================

' Basic function to generate a TypeID (recalculates every time)
Public Function GenerateTypeID(prefix As String) As String
    If Not IsValidPrefix(prefix) Then
        GenerateTypeID = "#ERROR: Invalid prefix"
        Exit Function
    End If
    
    Dim uuid() As Byte
    uuid = GenerateUUIDv7()
    
    Dim suffix As String
    suffix = EncodeBase32(uuid)
    
    If Len(prefix) > 0 Then
        GenerateTypeID = prefix & "_" & suffix
    Else
        GenerateTypeID = suffix
    End If
End Function

' ============================================================================
' UTILITY FUNCTIONS - PREVENT REGENERATION
' ============================================================================

' Generate a TypeID and write it directly to the cell (won't recalculate)
' User-facing macro without parameters (appears in macro list)
Public Sub InsertTypeID()
    Dim prefix As String
    
    ' Ask for prefix
    prefix = InputBox("Enter the prefix for TypeID (or leave empty):", _
                      "TypeID Prefix", "user")
    
    ' If user cancels
    If StrPtr(prefix) = 0 Then Exit Sub
    
    ' Generate TypeIDs in selected cells
    If TypeName(Selection) = "Range" Then
        Application.ScreenUpdating = False
        Dim cell As Range
        For Each cell In Selection
            cell.Value = GenerateTypeID(prefix)
        Next cell
        Application.ScreenUpdating = True
        
        MsgBox "Generated " & Selection.Count & " TypeID(s) with prefix '" & prefix & "'", vbInformation
    Else
        MsgBox "Please select one or more cells.", vbExclamation
    End If
End Sub

' Internal function to insert TypeID with specific prefix (for programmatic use)
Private Sub InsertTypeIDWithPrefix(prefix As String)
    If TypeName(Selection) = "Range" Then
        Dim cell As Range
        For Each cell In Selection
            cell.Value = GenerateTypeID(prefix)
        Next cell
    End If
End Sub

' User interface to generate TypeIDs in batch
Public Sub GenerateTypeIDsBatch()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim prefix As String
    
    Set ws = ActiveSheet
    
    ' Ask for prefix
    prefix = InputBox("Enter the prefix for TypeIDs (or leave empty):", _
                      "TypeID Prefix", "user")
    
    If StrPtr(prefix) = 0 Then Exit Sub
    
    ' Ask for cell range
    On Error Resume Next
    Set rng = Application.InputBox("Select cells where to generate TypeIDs:", _
                                   "Generation Range", _
                                   Type:=8)
    On Error GoTo 0
    
    If rng Is Nothing Then Exit Sub
    
    ' Disable recalculation for performance
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    ' Generate TypeIDs
    For Each cell In rng
        cell.Value = GenerateTypeID(prefix)
    Next cell
    
    ' Re-enable
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Generated " & rng.Count & " TypeID(s) with prefix '" & prefix & "'", vbInformation
End Sub

' Convert TypeID formulas to values to prevent regeneration
Public Sub ConvertFormulasToValues()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim count As Long
    
    Set ws = ActiveSheet
    
    On Error Resume Next
    Set rng = Application.InputBox("Select cells to freeze:", _
                                   "Convert to values", _
                                   Type:=8)
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        count = 0
        For Each cell In rng
            If cell.HasFormula Then
                cell.Value = cell.Value
                count = count + 1
            End If
        Next cell
        
        MsgBox count & " cell(s) converted to values.", vbInformation
    End If
End Sub

' Generate a TypeID only if shouldGenerate is TRUE (for conditional use)
Public Function TypeIDIf(shouldGenerate As Boolean, prefix As String, Optional seed As String = "") As Variant
    ' Use seed to create a stable value per row
    Static dict As Object
    Dim key As String
    
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
    End If
    
    ' Create a unique key based on cell address
    key = Application.Caller.Address & "_" & prefix
    
    If shouldGenerate Then
        ' If we haven't generated for this cell yet
        If Not dict.Exists(key) Then
            dict(key) = GenerateTypeID(prefix)
        End If
        TypeIDIf = dict(key)
    Else
        ' Reset if condition becomes false
        If dict.Exists(key) Then
            dict.Remove key
        End If
        TypeIDIf = ""
    End If
End Function

' ============================================================================
' SELF-REPLACING FUNCTION - Generates TypeID then converts itself to value
' ============================================================================

' Generate a TypeID that automatically converts to a static value
' Usage: =NewTypeID("user")
' After calculation, the formula is replaced by the generated value
Public Function NewTypeID(prefix As String) As String
    Dim typeID As String
    
    ' Validate prefix
    If Not IsValidPrefix(prefix) Then
        NewTypeID = "#ERROR: Invalid prefix"
        Exit Function
    End If
    
    ' Generate TypeID
    typeID = GenerateTypeID(prefix)
    NewTypeID = typeID
    
    ' Schedule the formula replacement
    ' This happens after the function returns, avoiding errors
    If Not IsEmpty(Application.Caller) Then
        Application.OnTime Now, "'ReplaceFormulaWithValue """ & Application.Caller.Address & """, """ & typeID & """'"
    End If
End Function

' Internal sub to replace formula with value (called by OnTime)
Public Sub ReplaceFormulaWithValue(cellAddress As String, value As String)
    On Error Resume Next
    Application.EnableEvents = False
    
    ' Replace the formula with the value
    Range(cellAddress).Value = value
    
    Application.EnableEvents = True
End Sub

' ============================================================================
' INTERNAL FUNCTIONS
' ============================================================================

' Validate prefix (max 63 chars, a-z lowercase only)
Private Function IsValidPrefix(prefix As String) As Boolean
    If Len(prefix) = 0 Then
        IsValidPrefix = True
        Exit Function
    End If
    
    If Len(prefix) > 63 Then
        IsValidPrefix = False
        Exit Function
    End If
    
    ' Check that prefix contains only a-z and underscore
    Dim i As Long
    Dim char As String
    For i = 1 To Len(prefix)
        char = Mid(prefix, i, 1)
        If Not ((char >= "a" And char <= "z") Or char = "_") Then
            IsValidPrefix = False
            Exit Function
        End If
    Next i
    
    IsValidPrefix = True
End Function

' Generate a UUID v7 (16 bytes)
Private Function GenerateUUIDv7() As Byte()
    Dim uuid(15) As Byte
    Dim i As Long
    
    ' Get Unix timestamp in milliseconds (48 bits)
    Dim timestamp As Double
    timestamp = (Now - #1/1/1970#) * 24 * 60 * 60 * 1000
    
    ' Decompose timestamp into two parts to avoid overflow
    Dim tsHigh As Long
    Dim tsLow As Long
    
    ' Divide timestamp in milliseconds
    tsHigh = Int(timestamp / 65536)  ' bits 16-47
    tsLow = CLng(timestamp - (CDbl(tsHigh) * 65536))  ' bits 0-15
    
    ' Extract timestamp bytes (big-endian, 48 bits = 6 bytes)
    uuid(0) = Int(tsHigh / 16777216) And 255        ' bits 40-47
    uuid(1) = Int(tsHigh / 65536) And 255           ' bits 32-39
    uuid(2) = Int(tsHigh / 256) And 255             ' bits 24-31
    uuid(3) = tsHigh And 255                         ' bits 16-23
    uuid(4) = Int(tsLow / 256) And 255              ' bits 8-15
    uuid(5) = tsLow And 255                          ' bits 0-7
    
    ' Version and variant (bytes 6-7)
    ' Byte 6: 4 version bits (0111 = v7) + 4 random bits
    uuid(6) = &H70 Or (Int(Rnd() * 16))
    
    ' Byte 7: 2 variant bits (10) + 6 random bits
    uuid(7) = &H80 Or (Int(Rnd() * 64))
    
    ' Fill remaining bytes with random values (8-15)
    For i = 8 To 15
        uuid(i) = Int(Rnd() * 256)
    Next i
    
    GenerateUUIDv7 = uuid
End Function

' Encode byte array to base32 (Crockford) - 130 bits (2 zeros + 128 bits UUID)
Private Function EncodeBase32(bytes() As Byte) As String
    Dim result As String
    Dim i As Integer
    Dim bitPos As Long
    Dim value As Integer
    
    result = ""
    
    ' TypeID encodes 130 bits: 2 zero bits + 128 bits of UUID
    ' 130 bits / 5 bits per character = 26 characters
    For i = 0 To 25
        ' Bit position in the 130 bits (0-129)
        bitPos = CLng(i) * 5
        
        ' Extract 5 bits from this position
        If bitPos = 0 Then
            ' First character (bits 0-4 of 130 bits):
            ' Bits 0-1 = 00 (padding)
            ' Bits 2-4 = first 3 bits of UUID (bits 0-2)
            ' So we take the 3 most significant bits of first byte
            value = (bytes(0) \ 32) And 7  ' Shift right 5 bits and keep 3 bits
        ElseIf bitPos = 5 Then
            ' Second character (bits 5-9 of 130 bits):
            ' = bits 3-7 of UUID
            ' = bits 3-7 of first byte
            value = (bytes(0) And 31)  ' The 5 least significant bits
        Else
            ' For other characters, adjust for 2-bit padding
            Dim uuidBitPos As Long
            uuidBitPos = bitPos - 2  ' Position in the 128 bits of UUID
            
            value = ExtractBitsFrom128(bytes, uuidBitPos)
        End If
        
        result = result & Mid$(BASE32_ALPHABET, value + 1, 1)
    Next i
    
    EncodeBase32 = result
End Function

' Extract 5 bits from a 16-byte array (128 bits) from a given position
Private Function ExtractBitsFrom128(bytes() As Byte, startBit As Long) As Integer
    Dim byteIndex As Long
    Dim bitOffset As Integer
    Dim currentByte As Integer
    Dim nextByte As Integer
    Dim value As Integer
    
    ' Determine starting byte and offset
    byteIndex = startBit \ 8
    bitOffset = startBit Mod 8
    
    ' Read current byte
    If byteIndex <= UBound(bytes) Then
        currentByte = bytes(byteIndex)
    Else
        currentByte = 0
    End If
    
    ' If all 5 bits are within a single byte
    If bitOffset <= 3 Then
        ' Extract 5 bits starting from bitOffset
        value = (currentByte \ (2 ^ (3 - bitOffset))) And 31
    Else
        ' The 5 bits span 2 bytes
        Dim bitsFromFirst As Integer
        bitsFromFirst = 8 - bitOffset
        
        ' Read next byte
        If byteIndex + 1 <= UBound(bytes) Then
            nextByte = bytes(byteIndex + 1)
        Else
            nextByte = 0
        End If
        
        ' Combine bits from both bytes
        Dim mask1 As Integer
        mask1 = (2 ^ bitsFromFirst) - 1
        
        Dim bitsFromSecond As Integer
        bitsFromSecond = 5 - bitsFromFirst
        
        value = ((currentByte And mask1) * (2 ^ bitsFromSecond)) Or (nextByte \ (2 ^ (8 - bitsFromSecond)))
    End If
    
    ExtractBitsFrom128 = value
End Function

' Helper function to generate multiple TypeIDs
Public Function GenerateTypeIDWithSeed(prefix As String, seed As Double) As String
    ' Initialize random generator with a seed
    Randomize seed
    GenerateTypeIDWithSeed = GenerateTypeID(prefix)
End Function

' ============================================================================
' QUICK GENERATION FUNCTIONS - For buttons with predefined prefixes
' ============================================================================

' Quick generation for "user" prefix
Public Sub QuickUserID()
    Call QuickGenerate("user")
End Sub

' Quick generation for "order" prefix
Public Sub QuickOrderID()
    Call QuickGenerate("order")
End Sub

' Quick generation for "product" prefix
Public Sub QuickProductID()
    Call QuickGenerate("product")
End Sub

' Helper function for quick generation
Private Sub QuickGenerate(prefix As String)
    If TypeName(Selection) = "Range" Then
        Application.ScreenUpdating = False
        Dim cell As Range
        For Each cell In Selection
            cell.Value = GenerateTypeID(prefix)
        Next cell
        Application.ScreenUpdating = True
        
        MsgBox "Generated " & Selection.Count & " TypeID(s) with prefix '" & prefix & "'", vbInformation
    Else
        MsgBox "Please select one or more cells first.", vbExclamation
    End If
End Sub
