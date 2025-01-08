Function IsWithinRange(number, min, max)
  'Corrected function to handle negative numbers and reversed ranges
  If min > max Then
    temp = min
    min = max
    max = temp
  End If
  IsWithinRange = (number >= min) And (number <= max)
End Function