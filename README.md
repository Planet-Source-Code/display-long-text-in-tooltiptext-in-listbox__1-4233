<div align="center">

## Display long text in tooltiptext in listbox


</div>

### Description

This code displays the text on the line the mouse is over in the tooltiptext box. This is useful for when your text string is longer than the textbox can display.
 
### More Info
 
Just place a list box on a form (less than 25 lines to see the

scrolling ability. More than 25 to see empty space handling)

On slow PC's with alot of stuff running the tipbox will appear to flash :(


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Unknown
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/display-long-text-in-tooltiptext-in-listbox__1-4233/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
  'load a bunch of long messages in the listbox
  For i = 0 To 25
    List1.AddItem (i & ". This is a long string that you can't _
            see all of in the list box, it's #: " & i)
  Next i
End Sub
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, _
              X As Single, Y As Single)
  'the height of the default text (you will have to change this
  'if you change the font size)
  WordHeight = 195
  'go through the loop until you get to the file
  For i = 0 To List1.ListCount - 1
    'check to what line the text is over (you need to go
    'through the whole list in case you've scrolled down
    If Y > WordHeight * (i - List1.TopIndex) _
      And Y < (WordHeight * i + WordHeight) Then
        'set the tooltiptext to the list box value
        List1.ToolTipText = List1.List(i)
    'see if your in "empty space"
    ElseIf Y > (WordHeight * i + WordHeight) Then
      List1.ToolTipText = "Empty space"
    End If
  Next i
End Sub
```

