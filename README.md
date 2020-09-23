<div align="center">

## Great stuff for people creating text editors\!


</div>

### Description

Read Below
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[RyanConard](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ryanconard.md)
**Level**          |Beginner
**User Rating**    |3.6 (36 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ryanconard-great-stuff-for-people-creating-text-editors__1-38209/archive/master.zip)





### Source Code

<i>In this example, you will learn a very simple way to add a Color and Font dialog control to your word/text application. You will also learn how to add a live "Character Counter" to your application as well</i>
<p>
<h2><u>Part 1: Adding Color</u></h2>
<br>
Ok, first lets cover the required objects. You will need a Common Dialog Control and a Rich Text Box control, which can both be found under the Components section of VB5.0 or VB6.0...
<p>
So, once you have those added, lets set the names. First, click on the Common Dialog Control you just added. Its name should be auto set to "CommonDialog1" but, were gonna rename it to "CD1" to shorten it and make it easier. Next, click on the Rich Text Box Control you just added. Its name is usually set to "RichTextBox1" but, were gonna rename it to "RTB1." Please note, that you can change these to what ever you want later, but for now just use what I gave you.
<p>
Ok, now we have our required controls and we named them to what we want. So, lets add the code! You now need to create a button. Name the button "cmdColor" and set it's caption to "Edit Back Color." Once you have that done, double click the new button so you can edit its code.
<p>
Inside the code, you should see:
<p>
<b>Private Sub cmdColor_Click()
<p>
End Sub</b>
<p>
Now, time for the good stuff! In between the code you just read above, put this:
<p>
<b>Dim SelectedColor As Long
<br>
CD1.Flags = cdlCCRGBInit
<br>
CD1.ShowColor
<br>
SelectedColor = CD1.Color
<br>
RTB1.BackColor = SelectedColor</b>
<p>
Thats it! The code is very basic. Basically what it does, is when you click cmdColor, the Common Dialog will open the Color Dialog Table. Once you choose a color from that table and click Ok, it will save the color you chose as our variable "SelectedColor." After it saves the color, it then calls the color into the RichTextBox and sets it as its background! Sounds complicated, but its not! And dont worry about editing the Font Color, we'll get to that in a minute!
<p>
<h2><u>Part 2: Editing Font</u></h2>
<p>
This part gets a little more tricky and has a few more lines of code, but its still very simple to understand.
<p>
Now, create a new button right next to your old one and name it "cmdFont." After, set its caption to "Edit Font." After you got that, double click the button so you can edit the code. You should now see something like this:
<p>
<b>Private Sub cmdFont_Click()
<p>
End Sub</b>
<p>
Now, within that code you need to put this:
<p>
<b>Dim TextColor As Long
<br>
Dim Bold As Boolean
<br>
Dim Italic As Boolean
<br>
Dim Underline As Boolean
<br>
Dim StrikeThru As Boolean
<br>
Dim Font As String
<br>
Dim Size As Integer
<br>
CD1.Flags = cdlCFEffects Or cdlCFBoth
<br>
CD1.ShowFont
<br>
TextColor = CD1.Color
<br>
Bold = CD1.FontBold
<br>
Italic = CD1.FontItalic
<br>
Underline = CD1.FontUnderline
<br>
StrikeThru = CD1.FontStrikeThru
<br>
Font = CD1.FontName
<br>
Size = CD1.FontSize
<br>
RTB1.SelFontName = Font
<br>
RTB1.SelFontSize = Size
<br>
RTB1.SelColor = TextColor
<br>
If Bold = True Then
<br>
RTB1.SelBold = Bold
<br>
If Italic = True Then
<br>
RTB1.SelItalic = Italic
<br>
If Underline = True Then
<br>
RTB1.SelUnderline = Underline
<br>
If StrikeThru = True Then
<br>
RTB1.SelStrikeThru = StrikeThru
<br>
End If
<br>
End If
<br>
End If
<br>
End If</b>
<p>
Wow, that was alot of typing for me! But, thats it! Once you have that code inside the button, you can edit every aspect of the font in a RichTextBox.
<p>
<h2><u>Part 3: Real-Time Character Counter</u></h2>
<p>
This is very cool. This will count every character you type and tell you how much you've typed.
<p>
All you need to do is create a Label. You can put it anywhere you, but name it "lblCount" and clear its caption. Once you have done that you need to double click the Rich Text Box so you can edit the code. Once you do that, you should see:
<p>
<b>Private Sub RTB1_Change()
<p>
End Sub</b>
<p>
Now, within that code you need to put this:
<p>
<b>Dim Text As String
<br>
Dim Count As Integer
<br>
Text = RTB1.Text
<br>
Count = Len(Text)
<br>
lblCount.Caption = Count</b>
<p>
Thats it! You now have a Character Counter!
<p>
Well, I gotta go! Hope you learned something from this, enjoy! Please Vote!

