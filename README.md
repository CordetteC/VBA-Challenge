This VBA script was created with the help of:
Instructor office hours - help with code structure (the order of the for loops and if statements)
Xpert – help with making the code run on multiple sheets and setting the conditional colors
“For Each ws In ThisWorkbook.Worksheets 
' Your code here
Next ws”
“Case Is > 0
        ws.Range("L" & 2 + l).Interior.Color = RGB(0, 255, 0) ' Green for positive change
    Case Is < 0
        ws.Range("L" & 2 + l).Interior.Color = RGB(255, 100, 100) ' Lighter red for negative change”
