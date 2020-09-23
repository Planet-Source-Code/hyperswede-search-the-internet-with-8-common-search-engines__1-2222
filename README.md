<div align="center">

## Search the Internet with 8 common search engines\!\!


</div>

### Description

Search the internet for web pages or files from your VB App.

8 common search engines:

AltaVista, Excite, HotBot, Infoseek, Lycos, Yahoo, SoftSeek, AudioFind

An easy to use function.

You can add as many Search Engines to the Search-engine list as you want, this list contains 8 search-engines. One of my lists contain 24 common search-engines.

Use like this:

SearchTheWeb Inputbox("Enter search-words:","",""), wAltaVista
 
### More Info
 
Just specify the search engine and words to search for.

Opens your web browser to a page showing the results of your search.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hyperswede](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hyperswede.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hyperswede-search-the-internet-with-8-common-search-engines__1-2222/archive/master.zip)

### API Declarations

It's all in the source code below.


### Source Code

```
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'The Search-Engine list:
Const SearchEngineList As String = "http://www.altavista.digital.com/cgi-bin/query?pg=q&what=web&fmt=.&q=||http://www.excite.com/search.gw?c=web&search=||http://www.hotbot.com/?SW=web&SM=MC&MT=||http://guide-p.infoseek.com/Titles?qt=|&col=WW&sv=IS&lk=noframes|http://www.lycos.com/cgi-bin/pursuit?query=||http://search.yahoo.com/bin/search?p=||http://search02.softseek.com/cgi-bin/search.cgi?keywords=|&seekindex=index&maxresults=025&cb=|http://www.audiofind.com:70/?audiofindsearch=|&audiofindtype="
Const wAltaVista As Long = 1
Const wExcite As Long = 3
Const wHotBot As Long = 5
Const wInfoseek As Long = 7
Const wLycos As Long = 9
Const wYahoo As Long = 11
Const wSoftSeek As Long = 13
Const wAudioFind As Long = 15
Function PartOfString(Str As String, Seperator As String, Number As Long)
Dim Current, Temp, Full
Current = 1
For q = 1 To Len(Str)
Temp = Mid(Str, q, 1)
If Temp = Seperator Then Current = Current + 1
If Current = Number And Not Temp = Seperator Then Full = Full & Temp
Next q
PartOfString = Full
End Function
Sub SearchTheWeb(ForWhat As String, WithWhat As Long)
ret& = ShellExecute(Me.hwnd, "Open", PartOfString(SearchEngineList, "|", WithWhat) & ForWhat & PartOfString(SearchEngineList, "|", WithWhat + 1), "", App.Path, 1)
End Sub
```

