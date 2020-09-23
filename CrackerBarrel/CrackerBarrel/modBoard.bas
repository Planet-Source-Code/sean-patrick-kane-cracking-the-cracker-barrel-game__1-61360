Attribute VB_Name = "modBoard"
Public Type board
    pegval(15) As Boolean
End Type

Public Function AddValidMoves(curarray As board, curpath As String, pegnum As Integer) As Boolean
'Pegs might be able to move NE/SW, E/W, or NW/SE -- we'll have to check for all six directions
Dim tmparray As board, tmppath As String, addedmove As Boolean
tmparray = curarray: tmppath = curpath  'Let's make some temporary copies to prepare for the first direction

'Move east (possible pegs: 1-3)
If (pegnum < 4) Then
If (tmparray.pegval(pegnum + 1) = True) And (tmparray.pegval(pegnum + 2) = False) Then
    tmppath = tmppath & "*" & pegnum & "E" & (pegnum + 2)
    tmparray.pegval(pegnum) = False 'the peg moved from this position
    tmparray.pegval(pegnum + 2) = True 'to this position
    tmparray.pegval(pegnum + 1) = False 'And jumped over this peg
    addedmove = True
    AddNewBoard tmparray, tmppath
    'Reset the tmp variables
    tmparray = curarray: tmppath = curpath
End If
End If
'Move west (possible pegs: 3-5)
If (pegnum > 2) And (pegnum < 6) Then
If (tmparray.pegval(pegnum - 1) = True) And (tmparray.pegval(pegnum - 2) = False) Then
    tmppath = tmppath & "*" & pegnum & "W" & (pegnum - 2)
    tmparray.pegval(pegnum) = False 'the peg moved from this position
    tmparray.pegval(pegnum - 2) = True 'to this position
    tmparray.pegval(pegnum - 1) = False 'And jumped over this peg
    addedmove = True
    AddNewBoard tmparray, tmppath
    'Reset the tmp variables
    tmparray = curarray: tmppath = curpath
End If
End If

'Move northeast
If (FindNcol(pegnum) < 4) And (FindRow(pegnum) < 4) Then
If (tmparray.pegval(CoordtoPegnum(FindRow(pegnum) + 1, FindPcol(pegnum))) = True) And (tmparray.pegval(CoordtoPegnum(FindRow(pegnum) + 2, FindPcol(pegnum))) = False) Then
    tmppath = tmppath & "*" & pegnum & "NE" & CoordtoPegnum(FindRow(pegnum) + 2, FindPcol(pegnum))
    tmparray.pegval(pegnum) = False 'the peg moved from this position
    tmparray.pegval(CoordtoPegnum(FindRow(pegnum) + 2, FindPcol(pegnum))) = True 'to this position
    tmparray.pegval(CoordtoPegnum(FindRow(pegnum) + 1, FindPcol(pegnum))) = False 'And jumped over this peg
    addedmove = True
    AddNewBoard tmparray, tmppath
    'Reset the tmp variables
    tmparray = curarray: tmppath = curpath
End If
End If
'Move southwest
If (FindNcol(pegnum) > 2) And (FindRow(pegnum) > 2) Then
If (tmparray.pegval(CoordtoPegnum(FindRow(pegnum) - 1, FindPcol(pegnum))) = True) And (tmparray.pegval(CoordtoPegnum(FindRow(pegnum) - 2, FindPcol(pegnum))) = False) Then
    tmppath = tmppath & "*" & pegnum & "SW" & CoordtoPegnum(FindRow(pegnum) - 2, FindPcol(pegnum))
    tmparray.pegval(pegnum) = False 'the peg moved from this position
    tmparray.pegval(CoordtoPegnum(FindRow(pegnum) - 2, FindPcol(pegnum))) = True 'to this position
    tmparray.pegval(CoordtoPegnum(FindRow(pegnum) - 1, FindPcol(pegnum))) = False 'And jumped over this peg
    addedmove = True
    AddNewBoard tmparray, tmppath
    'Reset the tmp variables
    tmparray = curarray: tmppath = curpath
End If
End If

'Move northwest
If (FindPcol(pegnum) > 2) And (FindRow(pegnum) < 4) Then
If (tmparray.pegval(CoordtoPegnum(FindRow(pegnum) + 1, , FindNcol(pegnum))) = True) And (tmparray.pegval(CoordtoPegnum(FindRow(pegnum) + 2, , FindNcol(pegnum))) = False) Then
    tmppath = tmppath & "*" & pegnum & "NW" & CoordtoPegnum(FindRow(pegnum) + 2, , FindNcol(pegnum))
    tmparray.pegval(pegnum) = False 'the peg moved from this position
    tmparray.pegval(CoordtoPegnum(FindRow(pegnum) + 2, , FindNcol(pegnum))) = True 'to this position
    tmparray.pegval(CoordtoPegnum(FindRow(pegnum) + 1, , FindNcol(pegnum))) = False 'And jumped over this peg
    addedmove = True
    AddNewBoard tmparray, tmppath
    'Reset the tmp variables
    tmparray = curarray: tmppath = curpath
End If
End If
'Move southeast
If (FindPcol(pegnum) < 4) And (FindRow(pegnum) > 2) Then
If (tmparray.pegval(CoordtoPegnum(FindRow(pegnum) - 1, , FindNcol(pegnum))) = True) And (tmparray.pegval(CoordtoPegnum(FindRow(pegnum) - 2, , FindNcol(pegnum))) = False) Then
    tmppath = tmppath & "*" & pegnum & "SE" & CoordtoPegnum(FindRow(pegnum) - 2, , FindNcol(pegnum))
    tmparray.pegval(pegnum) = False 'the peg moved from this position
    tmparray.pegval(CoordtoPegnum(FindRow(pegnum) - 2, , FindNcol(pegnum))) = True 'to this position
    tmparray.pegval(CoordtoPegnum(FindRow(pegnum) - 1, , FindNcol(pegnum))) = False 'And jumped over this peg
    addedmove = True
    AddNewBoard tmparray, tmppath
    'Reset the tmp variables
    tmparray = curarray: tmppath = curpath
End If
End If

AddValidMoves = addedmove 'If we return false, we can't move anywhere and this board is done
End Function

Public Function AddNewBoard(board_array As board, past_path As String)
Dim i As Integer
ReDim Preserve lstBoards(UBound(lstBoards) + 1)
lstBoards(UBound(lstBoards)).flag = ArrayToFlag(board_array)
lstBoards(UBound(lstBoards)).past = past_path
frmMain.cntBoards = UBound(lstBoards)
End Function

Public Function AddFinalBoard(board_array As board, past_path As String)
Dim i As Integer, strBoard As String, numleft As Integer
strBoard = CStr(ArrayToFlag(board_array))
For i = 1 To 15
    If board_array.pegval(i) = True Then
        numleft = numleft + 1
    End If
Next i
frmMain.lstFinal.AddItem "[" & numleft & "]" & past_path
frmMain.cntFinal = frmMain.cntFinal + 1
End Function

Public Function EvalBoard(curarray As board, curpath As String) As Integer
'This is the "parent" function that calls the function AddValidMoves
Dim NoMoves As Boolean, i As Integer
For i = 1 To 15
    If curarray.pegval(i) = True Then NoMoves = NoMoves Or AddValidMoves(curarray, curpath, i)
Next i

If NoMoves = False Then 'There were no moves for any of the pegs -- this board is completed
    AddFinalBoard curarray, curpath
    EvalBoard = -1
End If
End Function
