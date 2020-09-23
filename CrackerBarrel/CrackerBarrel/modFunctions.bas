Attribute VB_Name = "modFunctions"
'The game's coordinate system...this is used to create rules for when it is legal to move a piece
'  Row 5--------------- 15
'                       /\
'  Row 4------------- 13-14
'                     /\  /\
'  Row 3------------ 10-11-12
'                    / \/ \/ \
'  Row 2----------- 6--7--8--9
'                  / \/ \/ \/ \
'  Row 1--------- 1--2--3--4--5
'                / \/ \/ \/ \/ \
'               /  /\ /\ /\ /\  \
'              /  /  /  /  /  \  \
'             /  /  / \/ \/ \  \  \
'          Pcol /  /  /\ /\  \  \  \
'           1  2  3  4  5  \  \  \  \
'    (positive slope)    \  \  \  \  \
'                         \  \  \  \  \
'                        Ncol \  \  \  \
'                          1  2  3  4  5
'                          (negative slope)

Public Function FindRow(pegnum As Integer) As Integer   'Returns the row number of a given pegnum index
    If pegnum < 6 Then FindRow = 1: Exit Function
    If pegnum < 10 Then FindRow = 2: Exit Function
    If pegnum < 13 Then FindRow = 3: Exit Function
    If pegnum < 15 Then FindRow = 4: Exit Function
    FindRow = 5 'the only option remaining
End Function

Public Function FindPcol(pegnum) As Integer 'Returns the pcol number of a given pegnum index
    If pegnum = 5 Then FindPcol = 5: Exit Function
    If pegnum = 4 Or pegnum = 9 Then FindPcol = 4: Exit Function
    If pegnum = 3 Or pegnum = 8 Or pegnum = 12 Then FindPcol = 3: Exit Function
    If pegnum = 2 Or pegnum = 7 Or pegnum = 11 Or pegnum = 14 Then FindPcol = 2: Exit Function
    FindPcol = 1 'the only option remaining
End Function

Public Function FindNcol(pegnum) As Integer 'Returns the ncol number of a given pegnum index
    If pegnum = 1 Then FindNcol = 1: Exit Function
    If pegnum = 2 Or pegnum = 6 Then FindNcol = 2: Exit Function
    If pegnum = 3 Or pegnum = 7 Or pegnum = 10 Then FindNcol = 3: Exit Function
    If pegnum = 4 Or pegnum = 8 Or pegnum = 11 Or pegnum = 13 Then FindNcol = 4: Exit Function
    FindNcol = 5 'the only option remaining
End Function

Public Function CoordtoPegnum(row As Integer, Optional pcol As Integer, Optional ncol As Integer) As Integer   'Returns a pegnum based on two of three coordinates
    If row = 5 Then CoordtoPegnum = 15 'There's only one pegnum in row 5
    If row = 4 Then
        If pcol <> 0 Then
            CoordtoPegnum = CInt(Mid("1314", (pcol - 1) * 2 + 1, 2))
        Else
            CoordtoPegnum = CInt(Mid("1314", (ncol - 4) * 2 + 1, 2))
        End If
    End If
    If row = 3 Then
        If pcol <> 0 Then
            CoordtoPegnum = CInt(Mid("101112", (pcol - 1) * 2 + 1, 2))
        Else
            CoordtoPegnum = CInt(Mid("101112", (ncol - 2) * 2 - 1, 2))
        End If
    End If
    If row = 2 Then
        If pcol <> 0 Then
            CoordtoPegnum = CInt(Mid("6789", pcol, 1))
        Else
            CoordtoPegnum = CInt(Mid("6789", (ncol - 1), 1))
        End If
    End If
    If row = 1 Then 'Row 1 is really easy -- the pcol = ncol = pegnum
        If pcol <> 0 Then
            CoordtoPegnum = pcol
        Else
            CoordtoPegnum = ncol
        End If
    End If
End Function
