VBA_Digital_Rain
================

A pure VBA implementation of the Matrix's Digital Rain effect

* Workbook Code
 
```vbnet

Option Explicit
Option Compare Text

' When the workbook first opens
Private Sub Workbook_Open()

  Dim wb As Workbook

  For Each wb In Workbooks
    
    ' Find any workbooks which aren't personal or this one
    If wb.Name <> Me.Name And Right(wb.Name, 4) <> "xlsb" Then
    
      ' We don't want to risk someone losing their work if this crashes.
      MsgBox "Please do not use this with other workbooks open"
      Exit Sub
      
    End If
    
  Next wb

  ' Give the user the option to trigger it or simply open the workbook without starting anything
  If MsgBox("Start Digital Rain?", vbYesNo + vbQuestion, "Matrix") = vbYes Then Matrix

End Sub

```

* Module Code

```vbnet

Option Explicit

' Main sub
Sub Matrix()

  Dim row_count As Long, col_count As Long, i As Long
  
  ' Set the StatusBar so the user knows how to quit
  Application.staturbar = "Press ESC or Ctrl+Break to stop the macro."
  
  Application.ScreenUpdating = False
  
  ' square everything up and make it black
  Format_Cells
  
  ' Work out the visible dimensions so we fit the window
  With ActiveWindow.VisibleRange
    row_count = .Rows.Count
    col_count = .Columns.Count
  End With
  
  ' Set up the top row numbers
  Preset_Data col_count
  
  ' Set the black, greens, and white colours dependant on the top row numbers
  Configure_Conditional_Formats
  
  ' Either loop a set amount or infinitely
  'Do While True
  For i = 1 To 100
    
    ' Hide the work going on behind the scenes
    Application.ScreenUpdating = False
    
    ' Decrement the row that drives the formatting
    Update_Data col_count
    
    ' Write a random character matrix into the cells
    With ThisWorkbook.Sheets("Matrix")
      .Range(.Cells(2, 1), .Cells(row_count, col_count)).Value = Character_Matrix(row_count - 1, col_count)
    End With
  
    ' Show the finished results for this iteration
    Application.ScreenUpdating = True
    DoEvents
    
  'Loop
  Next i

End Sub

' Make it all square & black before we start
Private Sub Format_Cells()
  
  With ThisWorkbook.Sheets("Matrix")
    With .Cells
      ' Make the cells squareish
      .ColumnWidth = 2.71
      .RowHeight = 18
      ' Make the characters central
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      ' Prevent characters like "=" making Excel sulk
      .NumberFormat = "@"
      ' Black is the default colour
      .Interior.Color = vbBlack
      .Font.Color = vbBlack
    End With
    ' Get the selection cursor out of the way
    .[A1].Select
    ' Conceal the top row as that only has numbers in
    .Rows(1).EntireRow.Hidden = True
  End With
  
End Sub

' Set up the top row numbers that drive the conditional formatting
Private Sub Preset_Data(col_count)

  Dim i As Long
  For i = 1 To col_count
    ' The numbers don't have to be this far apart, but it's more stylish
    ThisWorkbook.Sheets("Matrix").Cells(1, i).Value = NumBetween(1000, 10000)
  Next i

End Sub

' Decrement the top row
Private Sub Update_Data(col_count)
  
  Dim i As Long
  For i = 1 To col_count
    With ThisWorkbook.Sheets("Matrix").Cells(1, i)
      ' Keep the viewer guessing with some random movement patterns
      If .Value Mod 50 < NumBetween(0, 20) Then
        ' go faster!
        .Value = .Value - 2
      ElseIf NumBetween(0, 30) = 0 Then
        'stop once in a while
      Else
        ' step down 1
        .Value = .Value - 1
      End If
    End With
  Next i

End Sub

' Set up the conditional formatting that creates the rain illusion
Private Sub Configure_Conditional_Formats()
  
  Dim step As Long
  Dim c As Range
  Dim f As FormatCondition
  
  ' Since there are vb constants for the rest, we may as well standardise that syntax
  Dim vbDarkGreen As Long, vbDarkerGreen As Long
  vbDarkGreen = 5287936
  vbDarkerGreen = 32768
  
  ' Give each column slightly different settings
  For Each c In ActiveWindow.VisibleRange.Columns
  
    ' Clean up anything already there
    c.FormatConditions.Delete
  
    ' Randomise it a bit
    step = NumBetween(9, 19)
    
    ' Randomise the font sizes to create a bit of illusory 3D
    c.Font.Size = NumBetween(4, 14)
    
    ' Each colour is a different position relative to the others, forming the trail
    
    Set f = c.FormatConditions.Add(Type:=xlExpression, Formula1:="=0=MOD(ROW()+A$1," & step & ")")
    f.Font.Color = vbWhite
    
    Set f = c.FormatConditions.Add(Type:=xlExpression, Formula1:="=0=MOD(ROW()+A$1+1," & step & ")")
    f.Font.Color = vbGreen
    
    Set f = c.FormatConditions.Add(Type:=xlExpression, Formula1:="=OR(0=MOD(ROW()+A$1+2," & step & "),0=MOD(ROW()+A$1+3," & step & "))")
    f.Font.Color = vbDarkGreen
    
    Set f = c.FormatConditions.Add(Type:=xlExpression, Formula1:="=OR(0=MOD(ROW()+A$1+4," & step & "),0=MOD(ROW()+A$1+5," & step & "))")
    f.Font.Color = vbDarkerGreen
    
    ' If it's a long column, let's give it a longer tail
    If step > 15 Then
      Set f = c.FormatConditions.Add(Type:=xlExpression, Formula1:="=OR(0=MOD(ROW()+A$1+6," & step & "),0=MOD(ROW()+A$1+7," & step & "))")
      f.Font.Color = vbDarkerGreen
    End If
    
  Next c
  
End Sub

' We fill a 2D Array with random characters, then return it
Private Function Character_Matrix(row_count As Long, col_count As Long) As String()

  Dim i As Long, j As Long
  Dim TempArray() As String
  ReDim TempArray(1 To row_count, 1 To col_count) As String
  
  ' Nested loops are fun
  For i = 1 To row_count
    For j = 1 To col_count
      TempArray(i, j) = Chr(NumBetween(32, 255))
    Next j
  Next i

  ' Return the Array
  Character_Matrix = TempArray

End Function

' Slightly more succinct way of writing this function
Private Function NumBetween(a As Long, b As Long) As Long
  NumBetween = WorksheetFunction.RandBetween(a, b)
End Function

```
