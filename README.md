<div align="center">

## Adding website's shortcut to desktop through website


</div>

### Description

To create a desktop website's shortcut through the website
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kapil Bhagia](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kapil-bhagia.md)
**Level**          |Intermediate
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kapil-bhagia-adding-website-s-shortcut-to-desktop-through-website__4-8123/archive/master.zip)





### Source Code

<BR><BR><small> <b>Adding link to the web browser's favourites is not a big deal. But to create shortcut of your website at the user's desktop is cool. And its pretty easy.
<BR><BR>Its simple, you create a VBS file (lets say abcd.vbs) with the following code:<BR><BR>
<table><tr><td nowrap>
dim WshShell, strDesktop, oUrlLink<BR>
set WshShell = CreateObject("WScript.Shell")<BR>
strDesktop = WshShell.SpecialFolders("Desktop")<BR>
set oUrlLink = WshShell.CreateShortcut (strDesktop + "\Your website's shortcut.URL")<BR>
oUrlLink.TargetPath = "http://www.yahoo.com" <BR>
oUrlLink.Save<BR><BR>
</td></tr></table>
You change the URL link in the second last line to your website URL and you can rename the shortcut name which will appear on the user's desktop, in the third last line.<Br><BR>
Now add the following code in the HTML where you want this functionality:<BR>
&lt;a href="abcd.vbs"&gt;Add shortcut to Desktop&lt;/a&gt; <BR><BR>
Try it out, if you find it good enough then vote for this article.</b></small>

