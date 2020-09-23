Attribute VB_Name = "modConvert"
Public Function FlagToArray(flag As Long) As board
Dim i As Integer
For i = 0 To 15
    If (flag And (2 ^ i)) = (2 ^ i) Then 'match
        FlagToArray.pegval(i) = True
    End If
Next i
End Function

Public Function ArrayToFlag(curarray As board) As Long
Dim i As Integer, curvalue As Long
For i = 1 To 15
    If curarray.pegval(i) = True Then curvalue = curvalue + (2 ^ i)
Next i
ArrayToFlag = curvalue
End Function
