Attribute VB_Name = "zDate"
Private Declare Sub GetLocalTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Function GetDateStr() As String
    Dim bTime As SYSTEMTIME
    GetLocalTime bTime
    GetDateStr = GetDay(bTime.wDayOfWeek) & " " & bTime.wDay & " " & GetMonth(bTime.wMonth) & " " & bTime.wYear
    
End Function

Private Function GetDay(DayOfWeek As Integer)
    Select Case DayOfWeek
        Case 0:  GetDay = "Sunday"
        Case 1:  GetDay = "Monday"
        Case 2:  GetDay = "Tuesday"
        Case 3:  GetDay = "Wednesday"
        Case 4:  GetDay = "Thursday"
        Case 5:  GetDay = "Friday"
        Case 6:  GetDay = "Saturday"
     Case Else:  GetDay = "Unknown"
    End Select
End Function

Private Function GetMonth(bMonth As Integer)
    Select Case bMonth
        Case 1:   GetMonth = "January"
        Case 2:   GetMonth = "Feburary"
        Case 3:   GetMonth = "March"
        Case 4:   GetMonth = "April"
        Case 5:   GetMonth = "May"
        Case 6:   GetMonth = "June"
        Case 7:   GetMonth = "July"
        Case 8:   GetMonth = "August"
        Case 9:   GetMonth = "September"
        Case 10:  GetMonth = "October"
        Case 11:  GetMonth = "November"
        Case 12:  GetMonth = "December"
      Case Else:  GetMonth = "Unknown"
    End Select
End Function

