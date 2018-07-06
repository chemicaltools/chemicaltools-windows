Attribute VB_Name = "modExam"
Public ExamScore As Double
Public ExamNo As Integer
Public ExamElementNumber As Integer
Public ExamIf As Boolean
Public ExamTime As Integer

Public examCorrectNumber As Integer
Public examIncorrectNumber As Integer

Function ExamAbbr(n As Integer, str As String) As Boolean
    '¶Ô·µ»Øtrue£¬´í·µ»Øfalse
    If UCase(str) = UCase(ElementAbbr(n)) Then
        ExamAbbr = True
    Else
        ExamAbbr = False
    End If
End Function

Function ExamRnd(n As Integer) As Integer
    Randomize
    ExamRnd = Int(Rnd() * n) + 1
End Function
