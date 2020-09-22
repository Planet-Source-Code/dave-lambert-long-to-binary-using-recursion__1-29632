Attribute VB_Name = "modLongToBinary"
Option Explicit

' ---------------------------------------------------------------------------------------
' Function to convert a value of type Long into a string representation of a binary value
' D R Lambert 2001 - http://www.drldev.co.uk
'
' Requires : Recursive Function LngToBin() (see below)
'
' Args     : lngValue - Long value to be converted to binary (positive, negative, or hex)
'
' Returns  : Binary representation of lngValue as a 32 character wide string
' ---------------------------------------------------------------------------------------
Public Function LongToBinary(ByVal lngValue As Long) As String
  Dim bNegative As Boolean
  Dim sResult As String
  
  ' Since we can't actually do anything with a positive value >= 2147483648 (&H80000000)
  ' without causing an overflow error, I test whether the number is negative then AND
  ' it with &H7FFFFFFF if it is. This means we only ever have to deal with what appears
  ' to be a positive value. We "bolt" the the value of the Negate bit back onto the front
  ' of the string representing the binary value just before it is returned.
  
  bNegative = (lngValue < 0)  ' Note whether lngValue is negative
  
  If bNegative Then           ' Convert lngValue into a positive number
    lngValue = lngValue And &H7FFFFFFF
    sResult = "1"             ' The "Negate" bit to be prepended to the result string
  Else
    sResult = "0"             ' The "Negate" bit to be prepended to the result string
  End If
  
  ' Call the recursive function to build the binary value
  sResult = sResult & Right$(String(31, "0") & LngToBin(lngValue), 31)
  
  ' Return the result
  LongToBinary = sResult
End Function


' ---------------------------------------------------------------------------------------
' Recursive function used for conversion of a positive long value into
' a string representation of a binary value.
' D R Lambert 2001 - http://www.drldev.co.uk
'
' Arguments: Value  Long value to be converted
'            Bit    Bit being tested
'
' Returns  : String representation of the current bit being tested, "0" or "1"
' ---------------------------------------------------------------------------------------
Private Function LngToBin(ByRef Value As Long, Optional ByRef Bit As Long = 1) As String
  ' As an optimisation, we only bother converting up to the highest bit that is
  ' set to 1 (Bit <= Value) and also make sure we don't generate an overflow
  ' error (Bit < &H40000000)
  If Bit <= Value And Bit < &H40000000 Then
    LngToBin = LngToBin(Value, Bit * 2) & CStr((Value And Bit) \ Bit)
  Else ' Either Bit > Value and/or Bit has reached the maximum allowed (&H40000000)
       ' i.e. this is the last recursion
    LngToBin = CStr((Value And Bit) \ Bit)
  End If
End Function

