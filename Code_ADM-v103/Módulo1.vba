Option Compare Database

Public Function WorkingDays(StartDate As Date, EndDate As Date) As Integer
'....................................................................
' Name:     WorkingDays
' Inputs:     StartDate As Date
'     EndDate As Date
' Returns: Integer
' Author: Arvin Meyer
' Date:     February 19, 1997
' Comment: Accepts two dates and returns the number of weekdays between them
' Note that this function does not account for holidays.
'....................................................................
On Error GoTo Err_WorkingDays

Dim intCount As Integer

StartDate = StartDate + 1
'If you want to count the day of StartDate as the 1st day
'Comment out the line above

intCount = 0
Do While StartDate <= EndDate
    'Make the above < and not <= to not count the EndDate
    
    Select Case Weekday(StartDate)
    
    Case Is = 1, 7
        intCount = intCount
    Case Is = 2, 3, 4, 5, 6
        intCount = intCount + 1
    
    End Select
    
    StartDate = StartDate + 1
Loop

WorkingDays = intCount

Exit_WorkingDays:
Exit Function

Err_WorkingDays:
Select Case Err

Case Else
MsgBox Err.Description
Resume Exit_WorkingDays
End Select

End Function
