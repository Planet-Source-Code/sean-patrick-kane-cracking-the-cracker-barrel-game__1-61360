Attribute VB_Name = "modPublic"
Public lstBoards() As liveboards
Public tmpBoards() As liveboards

Public Type liveboards
    flag As Long
    past As String
    delete As Boolean
End Type
