<div align="center">

## DetectPreviousInstance


</div>

### Description

How do I prevent multiple instances of my program? In VB 3 and above, the property App.PrevInstance is set to True if an older instance of the program already exist.
 
### More Info
 
As Robert Knienider(rknienid@email.tuwien.ac.at) informed me, this piece of code will not work for non-English versions of Mirosoft Windows where the word for "Restore" does not have "R" as the underlined word. Replace the "R" in the SendKeys line above with "{ENTER}" or "~".

Note that you shouldn't prevent multiple instances of your application unless you have a good reason to do so, since this is a very useful feature in MS Windows. Windows will only load the code and dynamic link code once, so it (normally) uses much less memory for the later instances than the first.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB FAQ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-faq.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-faq-detectpreviousinstance__1-74/archive/master.zip)





### Source Code

```
Sub Form_Load ()
  If App.PrevInstance Then
    SaveTitle$ = App.Title
    App.Title = "... duplicate instance." 'Pretty, eh?
    Form1.Caption = "... duplicate instance."
    AppActivate SaveTitle$
    SendKeys "% R", True
    End
  End If
End Sub
```

