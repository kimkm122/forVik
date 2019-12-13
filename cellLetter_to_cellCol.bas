'Here's some logic to convert column letters to a number.
'Python might have a better way to get the index number of a match within an array.

Function cellLetter_to_cellColumn(ByVal cellLetters As String) As Integer
    Dim letter As String
    Dim letterPosition As Integer
    Dim letterNumber As Integer
    
    For letterPosition = 1 To Len(cellLetters)
        letter = Mid(cellLetters, Len(cellLetters) - letterPosition + 1, 1)
        letterNumber = letter_to_number(letter)
        cellLetter_to_cellColumn = cellLetter_to_cellColumn + (letterNumber * 26 ^ (letterPosition - 1))
    Next
End Function
Function letter_to_number(ByVal letter As String) As Integer
    Dim alphabet As Variant
    Dim i As Integer
    alphabet = Array("A", "B", "C", "D", "E", _
                    "F", "G", "H", "I", "J", _
                    "K", "L", "M", "N", "O", _
                    "P", "Q", "R", "S", "T", _
                    "U", "V", "W", "X", "Y", "Z")
    For i = 0 To UBound(alphabet)
        If alphabet(i) = letter Then
            letter_to_number = i + 1
            Exit Function
        End If
    Next
End Function
