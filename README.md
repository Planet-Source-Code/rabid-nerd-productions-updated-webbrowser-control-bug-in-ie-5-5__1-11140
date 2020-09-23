<div align="center">

## UPDATED\! WebBrowser Control Bug in IE 5\.5


</div>

### Description

After developing a full-blown application (just went out to beta) for my employer, I uncovered a bug in IE 5.5 that you may want to know about.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rabid Nerd Productions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rabid-nerd-productions.md)
**Level**          |Advanced
**User Rating**    |4.7 (61 globes from 13 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rabid-nerd-productions-updated-webbrowser-control-bug-in-ie-5-5__1-11140/archive/master.zip)





### Source Code

<font size="2" color=red>UPDATE! - 12/26/2000</font><br>Microsoft LISTS THIS BUG OFFICIALLY at:
<a href="http://support.microsoft.com/support/kb/articles/Q279/6/68.ASP">http://support.microsoft.com/support/kb/articles/Q279/6/68.ASP</a><br><font color=red>End UPDATE </font><p>
<font size="2" color=red>UPDATE! - Sent from Microsoft</font><br><font face="arial">Herb,<br>
<br>
 <br>
<br>
Thanks for sending in the codes. I am able to <br>reproduce the same problem with your code. I verified that in IE5 navigate2 accept a string as <br>URL, in IE5.5 it only accept a variant. Or you could pass in the url inline without using a <br>variable. Then I looked at the source of both versions and set up a debugger to verify. Here is the situation:<br>
<br>
When you specific the URL as a string, VB will passed in VT_BSTR|VT_BYREF<br>
<br>
In both IE5.0 and IE5.5, the header info are the same in the IDL file which accept a variant<br>
<br>
[id(0x000001f4), helpstring("Navigates to a URL or file or pidl.")]<br>
<br>
HRESULT Navigate2(<br>
<br>
    [in] VARIANT* URL, <br>
<br>
    [in, optional] VARIANT* Flags, <br>
<br>
    [in, optional] VARIANT* <br>TargetFrameName, <br>
<br>
    [in, optional] VARIANT* PostData, <br>
<br>
    [in, optional] VARIANT* Headers);<br>
<br>
However, the implementations are different.<br>
<br>
In IE5.0, the URL param could be VT_BSTR, <br>VT_BSTR|VT_BYREF or VT_ARRAY|VT_UI1<br>
<br>
In IE5.5, it uses a function to determine the values of the URL and it only accepts VT_BYREF | <br>VT_VARIANT or VT_BSTR<br>
<br>
So, it returns an error E_INVALIDARG when VB passed in VT_BSTR|VT_BYREF.<br>
<br>
 <br>
And I filed a bug on this.<br>
<br>
So at the mean time, you would have to use variant as I believe passing the url inline<br> probably won&#8217;t go too far in most applications.
<br>
 <br>
<br>
Please let me know if you have any more concerns.
<br><br>
Joshua Lee (MCP + Site Building)<br>
<br>
Content Lead<br>
<br>
Internet Client Team<br>
<br>
Microsoft Developer Support<br>
</font><p><font color=red>End UPDATE. Original article follows:</font><br>
<font size="2">In the Microsoft Knowledge Base article at:<p><a href="http://support.microsoft.com/support/kb/articles/Q269/6/14.ASP?LN=EN-US&SD=SO&FR=1">http://support.microsoft.com/support/kb/articles/Q269/6/14.ASP?LN=EN-US&SD=SO&FR=1</a><p>A bug in IE 5.5 users' WebBrowser controls is exposed.<br>The WebBrowser1_NavigateComplete event is NOT FIRED when the control is set to visible = FALSE.(Invisible)<p>I have also found another bug (and I am working with Microsoft on it) with the 5.5 WebBrowser Control:<p>When you use a string variable in the Navigate2 method, the control fails to navigate (and may cause an error!)<p>
Here is a code that you can place in a form with a webbrowser control on it:<p><pre>
Sub form_load()
Dim urly As String
urly = "http://directleads.com/ad.html?o=993&a=cd15860"
WebBrowser1.Navigate2 urly
End Sub
Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
MsgBox "Hello!"
End Sub</pre><p>When you execute this code on a computer with IE other than 5.5, You will get the appropriate MsgBox with<br>Hello! in it (and our website). With 5.5, however, you will get and Error 5.<p>The workaround is to declare the URL variable (urly) as a Variant <b>-OR-</b> Use the Navigate (no 2) Method. (as described in the<br> <a href="http://support.microsoft.com/support/kb/articles/Q167/4/35.asp" target="_blank">MS Knowledge Base</a> (See #28). The variant may not be as efficient as the string type, but it works on ALL versions of IE. <p>If you appreciate this 'bulletin' of sorts, I humbly ask for a vote. I understand this is not an application, but hey<br> it's important to anyone with a webbrowser control on their app!<p>Have Fun,<br>Herbert L. Riede<br>Programmer, <a href="http://directleads.com/ad.html?o=993&a=cd15860" target="_blank">WinDough.com, Inc.</a></font>

