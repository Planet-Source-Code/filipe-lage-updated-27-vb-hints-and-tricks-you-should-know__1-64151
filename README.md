<div align="center">

## Updated \- 27 VB Hints and Tricks you should know


</div>

### Description

Just a bunch of VB6 hints and tricks I thought I could share with you.

Many new hints and tricks you should know in VB6.

How to add events to Windows Application log

How to add controls in run time

VB6 and the 2GB File limit - Be aware

How to hide your application from task manager

ASM Subclassing - Moving back is the safest way

How to check for non-Modal permitions

How to implement DIR$ correctly in your application.

Convert ByteArrays to String and vice versa

And more...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Filipe Lage](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/filipe-lage.md)
**Level**          |Intermediate
**User Rating**    |4.6 (64 globes from 14 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/filipe-lage-updated-27-vb-hints-and-tricks-you-should-know__1-64151/archive/master.zip)





### Source Code

<p align="center"><font face="Verdana" size="4"><b>VB Hints:</b></font></p>
<p><font face="Verdana"><b>Give description to your print jobs:</b></font></p>
<blockquote>
	<p><font face="Verdana">Before printing set APP.TITLE to the description of
	the print job. <br>
	This way, the description of the document being printer will appear in the
	printer job window instead of your application <br>
	name.<br>
	Ex:</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">app.title = "Invoice #1234"</font></p>
	</blockquote>
	<p><font face="Verdana">[do the printing code and enddoc]</font></p>
</blockquote>
<p><font face="Verdana"><b>Quickly get data from a separated string</b></font></p>
<blockquote>
	<p><font face="Verdana">Let's consider </font></p>
	<blockquote>
		<p><font face="MS Sans Serif">dim a as string<br>
		dim atmp() as string<br>
		a="Text1;Text2;Text3;Text4"<br>
		atmp=split(a,";")<br>
		Test = a(3)</font></p>
	</blockquote>
	<p><font face="Verdana">Now let's look at...</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">dim a as string<br>
		a="Text1;Text2;Text3;Text4"<br>
		Test = split(a,";")(3)</font></p>
	</blockquote>
	<p><font face="Verdana">This way you can get the "Text4" string directly
	from split instead of mapping a temporary string (previous example). <br>
	It's actually faster too ;)<br>
	<br>
	Naturally, this also applies to tab delimited files. Example, Create a file
	in excel and export as TXT (Tab delimited)<br>
	You can get the cell from the respective row and column after reading
	contents to memory</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">function CellData(data as string, row as
		integer, column as integer) as variant<br>
		CellData = split(split(data,vbcrlf)(Row),vbtab)(Column)<br>
		end function</font></p>
	</blockquote>
	<p> </p>
</blockquote>
<p><font face="Verdana"><b>Quickly get rounding of a number the right way:</b></font></p>
<blockquote>
	<p><font face="Verdana">Since VB rounds fail in mathematical functions (ex:
	Round(2.5) results in 2 instead of 3)<br>
	we can avoid that by creating a new Round Function in VB</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">function MathRound(value as Double,
		optional lngDecimals as Long = 0) as double<br>
		MathRound = CDbl(Format$(value * 10^lngDecimals, "0")) / 10^lngDecimals<br>
		end function</font></p>
	</blockquote>
	<p><font face="Verdana">This function also supports negative decimal places.<br>
	Ex: <i>MathRound(1100,-3)=1000</i><br>
	<br>
	There's a faster function (also created by me) available on the net if you
	prefer speed over simplicity.<br>
	Check it at <br>
	http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=61414&lngWId=1<br>
 </font></p>
</blockquote>
<p><b><font face="Verdana">Immediate window</font></b></p>
<blockquote>
	<p><font face="Verdana">In case you don't know, you can use "?" to replace "debug.print"
	in the immediate window<br>
	Ex:<br>
	Go to immediate window and type:</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">? 1+1</font></p>
	</blockquote>
	<p><font face="Verdana">It will appear: 2 (naturally)<br>
	It's useful to test functions or parts of code parsing the output value to
	the immediate window, or simply to check variables.<br>
	Ex: </font></p>
	<blockquote>
		<p><font face="MS Sans Serif">? a=3<br>
		False</font></p>
	</blockquote>
	<p><font face="Verdana">BTW, if type it in your VB code itself, the "?" will
	be replaced to "Print" just after you press enter.<br>
	Ex:</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">me.? "test"</font></p>
	</blockquote>
	<p><font face="Verdana">will be automatically replaced to:</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">Me.Print test</font></p>
	</blockquote>
</blockquote>
<p><b><font face="Verdana">Use mouseweelfix</font></b></p>
<blockquote>
	<p><font face="Verdana">VB6 doesn't support mousewheel natively, so you
	can't scroll up and down with your mouse.<br>
	But there's a fix... actually, it's an addon for VB that scrolls the text
	and implements that so-much-needed function.<br>
	You can find it in microsoft here: http://support.microsoft.com/?id=837910<br>
 </font></p>
</blockquote>
<p><b><font face="Verdana">How to register an ActiveX in code.</font></b></p>
<blockquote>
	<p><font face="Verdana">Just make a declaration such as:</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">Private Declare Function REGISTER_MYDLL
		Lib "MYDLL.DLL" _<br>
  Alias "DllRegisterServer" () As Long<br>
		Private Declare Function UNREGISTER_MYDLL Lib "MYDLL.DLL" _<br>
  Alias "DllUnregisterServer" () As
		Long<br>
		Private Const ERROR_SUCCESS = &H0</font></p>
	</blockquote>
	<p><font face="Verdana">Then, simply call</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">REGISTER_DLL</font></p>
	</blockquote>
	<p><font face="Verdana">or</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">UNREGISTER_DLL</font></p>
	</blockquote>
	<p><font face="Verdana">in your code to register or unregister the dll.<br>
	<br>
	There are more advanced functions that enable you to specify the DLL to get
	registered on code instead of <br>
	mapping the actual DLL file in the declaration. But if you know your DLL's,
	then simply include the declarations in the EXE <br>
	and add <br>
	an option to Fix components by calling the respectivefunctions<br>
	This also works for OCX files.</font></p>
</blockquote>
<p><b><font face="Verdana">XCOPY Install?</font></b></p>
<blockquote>
	<p><font face="Verdana">YES! That is, if you only use OCX files...<br>
	In case your EXE requires external OCX files, and you put them in the same
	path as the EXE, they are loaded without any <br>
	problems! :)<br>
	Doesn't work for dll's though. You still have to register them using
	regsvr32 or using the previous hint.<br>
	If you want to make sure your app works even without VB6 Runtimes, simply
	include MSVBVM60.DLL to your EXE path, but don't <br>
	register it.<br>
	But if you want to run your application even though you're not sure if the
	target system has VB6 runtimes, just include<br>
	the MSVBVM60.DLL in the same path as the EXE... it works :). <br>
	Note for "Virgin" Win95/98: <br>
	Must have DCOM previous installed (it is installed with IE 4.5 anyway and
	MDAC's, and other updates)<br>
 </font></p>
</blockquote>
<p><b><font face="Verdana">How to enable/disable all controls in a form/container
</font></b></p>
<blockquote>
	<p><font face="Verdana">to disable:</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">on error resume next<br>
		for each o in me.controls: o.enabled = False: next</font></p>
	</blockquote>
	<p><font face="Verdana">to enable:</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">on error resume next<br>
		for each o in me.controls: o.enabled = True: next<br>
 </font></p>
	</blockquote>
</blockquote>
<p><b><font face="Verdana">How to change container of an object.</font></b></p>
<blockquote>
	<p><font face="Verdana">Example, place a commandbutton in a form, and add a
	frame next to it. <br>
	Note that both controls will have Form1 as a parent.<br>
	If you want commandbutton to be included inside the frame, but to make it
	work in runtime simply add:</font></p>
	<blockquote>
		<p><font face="MS Sans Serif">SET Command1.container = Frame1<br>
 </font></p>
	</blockquote>
</blockquote>
<p><b><font face="Verdana">Avoid using VB strings greater than 32k</font></b></p>
<blockquote>
	<p><font face="Verdana">VB Strings is the "Achilles's heel" of VB in terms of
	speed. (ok, strings and threading/subclassing)<br>
	I recomend you use a stringhelper object (check AllocString page at http://www.xbeat.net/vbspeed/)
	if you want <br>
	big strings (1MB or more) to store data. Beware of this AllocString since
	the data inside the string will not be blank!</font></p>
	<p>&nbsp;</p>
</blockquote>
<p><font face="Verdana"><i><font size="1">[Added 2006-02-02]<br>
</font></i><b>Read file from disk into memory (Fastest way possible without API
with low cpu usage)</b></font></p>
<blockquote>
	<p><font face="Verdana">I've done this function to obtain all data from an
	existing file to memory. Note that if you have very large files (like 1GB)
	it will take 1GB of RAM as well... It's great to read data and handle it in
	memory. I get about 7MBytes/second in my P4-3000.</font></p>
	<p><font face="MS Sans Serif">Public Function ReadFullFile(file As String)
	As Byte()<br>
	Dim a As Long<br>
	a = FreeFile<br>
	Open file For Binary As #a<br>
	ReDim ReadFullFile(LOF(a)-1)<br>
	Get #a, , ReadFullFile<br>
	Close #a<br>
	End Function</font><font face="Verdana"><br>
 </font></p>
	<p><font face="Verdana">It stores all data in a bytearray... It's better
	than storing in a VB String, since all VB strings are stored in UNICODE,
	meaning that for each byte in the file, it will take 2 bytes of RAM. So, if
	I used a string, I would need 200MB of RAM to read a file of 100MB.</font></p>
	<p><font face="Verdana">Naturally, you can convert it to string using the
	code:</font></p>
	<p><font face="MS Sans Serif">Dim FileData as string<br>
	FileData = StrConv(Readfullfile(file),vbUnicode)</font></p>
	<p><font face="Verdana">Be careful, since you can run out of out of memory
	with very large files!<br>
	You should also take in consideration, that if the file has 0 bytes or
	simply doesn't exist, it will result in an error, so you should make sure
	that the file being read exists and it's not 0 bytes long.</font></p>
	<p><font face="Verdana">You can also change the function to read the file
	directly to a string, by using</font></p>
	<p><font face="MS Sans Serif">Public Function ReadFullFile(file As String)
	As String<br>
	Dim a As Long<br>
	a = FreeFile<br>
	Open file For Binary As #a<br>
	ReadFullFile = Space(LOF(a)) ' You can use the VBSpeed's StringHelper to
	make this faster for large files<br>
	Get #a, , ReadFullFile<br>
	Close #a<br>
	End Function</font><font face="Verdana"><br>
 </font></p>
</blockquote>
<p><b><font face="Verdana">Avoid VB IDE bugs</font></b></p>
<blockquote>
	<p><font face="Verdana">I've used VB for several years now, and I've
	discovered several bugs that many times corrupt projects and you should be
	aware of that.</font></p>
	<p><font face="Verdana"><b><i>Lost bags</i></b><br>
	First of all, if you have an UserControl present in your EXE project,
	remember that if you change the project name, all properties previously set
	in your forms will be lost. To be exact, all <i>propbag</i>'s in your user
	controls become "blanks" and defaults are used.<br>
	Example:<br>
	You've added an UserControl to your project and you're using it in Form1.
	One of the properties of that user control is "Caption" and you've set that
	to "Hello world"... Nice, that is saved on the usercontrol's propbag... If
	you change the EXE project name, and check your form again, the "Hello world"
	is now gone.</font></p>
	<p><font face="Verdana"><b><i>Corrupt VBP's<br>
	</i></b>One of the most annoying things in VB6 is that it sometimes corrupts
	the VBP's by mistaking some ActiveX objects with ActiveX controls.<br>
	Example: If you have ActiveX DLL's in your project references, and you also
	use external ActiveX Objects (usercontrols) in your forms, sometimes VB6
	will list the object as a reference. In conclusion, when you open the VBP
	once again, it will give a load error and all forms that use the "mixed up"
	control will have their objects replaced with a picture box.</font></p>
	<p><font face="Verdana">This happens when you open several projects in VB (ex:
	EXE + DLL) and use an external OCX UserControl, compile the DLL with the other projects loaded, quit VB
	and save changes. After that, just load the EXE VBP.<br>
	When this problem happens, the solution I've found is to open the VBP with
	notepad and delete the "Reference" line that includes the OCX/VBP. Open the
	VBP, include the OCX once again in the add controls, and re-save the project
	(just the VBP). Reopen the VBP and all is well again.</font></p>
</blockquote>
<p><font face="Verdana"><br>
<i><font size="1">[Added 2006-03-07]<br>
</font></i><b>Avoid using Subclassing... At least with ASM code on it</b><br>
&nbsp;&nbsp;&nbsp; Until recently, I've been using the ASM subclassing from (the
great)
VBAccelerator.com. The file ssubtmr6.dll to be exact.<br>
&nbsp;&nbsp;&nbsp; Unfortunatly, I had to return to the previous non-ASM code since every call crashed my application in a computer I had...
<br>
&nbsp;&nbsp;&nbsp; I investigated, and I found
the reason... DEP - Data Execution Prevention... <br>
&nbsp;&nbsp;&nbsp; Naturally, when DEP is used Windows XP and 2003.NET with a
DEP compliant CPU (ex: AMD64 or the latests Intel CPU's), Windows will deny the ASM part of the subclassing
to run (since the code is stored in a variable area and not in a code execution
area)... <br>
&nbsp;&nbsp;&nbsp; Windows automatically shows a GPF when the subclassing is initialized in
this mode and the application is closed.<br>
<br>
&nbsp;&nbsp;&nbsp; There are two ways to avoid the problem:<br>
&nbsp;&nbsp;&nbsp; 1) Not recomended<br>
&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp; Change the boot.ini of
the operating system (not recomended) or add your application to the DEP 'exclusion'
lists.<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Either way, it doesn't garantee a crash free operation,
and the user has to add your application to the exclusion list manually.<br>
&nbsp;&nbsp;&nbsp; 2) Recomended<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Return to the previous subclassing that
doesn't use the ASM code.<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; A little slower, but it's crash free with DEP compliant cpu's.</font></p>
<p><font face="Verdana">&nbsp;&nbsp;&nbsp; This is my recomendation if you want to make a stable,
subclassing application to be used in Windows XP and/or 2003.NET.</font></p>
<p>&nbsp;</p>
<p><b><font face="Verdana">On error resume next... Beware!</font></b><font face="Verdana"><br>
&nbsp;&nbsp;&nbsp; If you want a 99% crash free app, you can always add the on
error resume next on the first line of every sub and function... I don't
recomend it, but at least it won't show any VB Runtime errors... However,
remember to set &quot;on error goto 0&quot; before the end of the function/sub, if not,
your code may not work at all (exits the first function calls another one
without resuming the next line if an error was raised in the second function).&nbsp;&nbsp;&nbsp; </font></p>
<p><font face="Verdana"><br>
&nbsp;</font></p>
<p><b><font face="Verdana">DIR$ - A great thing if you implement it right.</font></b><font face="Verdana"><br>
&nbsp;&nbsp;&nbsp; VB has the DIR$ function so you can list folders and files,
however you should be aware that this function is shared across your entire
application. So, if you have one function that does something like:<br>
<br>
</font><font face="MS Sans Serif">&nbsp;&nbsp;&nbsp; Sub ListFolder()<br>
&nbsp;&nbsp;&nbsp; mainpath = &quot;c:\windows&quot;<br>
&nbsp;&nbsp;&nbsp; a$=Dir$(mainpath)<br>
&nbsp;&nbsp;&nbsp; do until a$=&quot;&quot;<br>
&nbsp;&nbsp;&nbsp; b$=HowManyDirs(mainpath &amp; &quot;\&quot; &amp; a$)<br>
&nbsp;&nbsp;&nbsp; a$=Dir$<br>
&nbsp;&nbsp;&nbsp; Loop<br>
<br>
&nbsp;&nbsp;&nbsp; Function HowManyDirs(f) as long<br>
&nbsp;&nbsp;&nbsp; b$=dir$(f,vbDirectory)<br>
&nbsp;&nbsp;&nbsp; do until b$=&quot;&quot;<br>
&nbsp;&nbsp;&nbsp; HowManyDirs=HowManyDirs+1<br>
&nbsp;&nbsp;&nbsp; b$=dir$<br>
&nbsp;&nbsp;&nbsp; loop<br>
&nbsp;&nbsp;&nbsp; end function<br>
</font><font face="Verdana"><br>
&nbsp;&nbsp;&nbsp; This code won't work at all, since the Dir$ function is
common in the entire application. If the first sub is using the Dir$, no
function should use it before the first sub finishes. You won't get the right
results if you do.</font></p>
<p>&nbsp;</p>
<p><font face="Verdana"><b>Beyond 2GB files with VB<br>
&nbsp;&nbsp;&nbsp;&nbsp; </b>All VB functions to get file size (LOF(x) or
FileLen(f)) are limited to longs... <br>
&nbsp;&nbsp;&nbsp; That means that you only have 31 bits (+1 bit for sign) to
store the size of the file... That gives VB a limit of 2147483648 bytes (2GB).<br>
&nbsp;&nbsp;&nbsp; Using internal functions and calls, like OPEN, SEEK, etc, you
can't get data beyond this point, so you'll need to use the API for that.<br>
&nbsp;&nbsp;&nbsp; Anyway, you should be aware of this VB6 limitation if your
project deals with very large files (ex: VOB's, MPG's, AVI's, etc) so you can
implement the necessary 64-bit functions to avoid this limitation.<br>
&nbsp;&nbsp;&nbsp; In terms of internal results (ex: a function you implement to
get the file size in 64-bits), I recomend you use the CURRENCY to get the file
size... Even though &quot;Currency&quot; data type isn't a full integer type, you'll have
the limit raised to 922.337.203.685.477 bytes (around 920 TeraBytes) per file
that I think is good enough for the next few years ;)<br>
<br>
&nbsp;&nbsp;&nbsp; Hint: If you're using a function to check if a file exists on
the hard disk, and your code is similar to:<br>
&nbsp;&nbsp;&nbsp; </font><font face="MS Sans Serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Function DoesFileExist(f as string) as Boolean<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
On error resume next<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
DoesFileExist = (filelen(f)&gt;0)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
End Function<br>
</font><font face="Verdana">&nbsp;&nbsp;&nbsp; You should be aware that VB's
FileLen function reports negative values when a file is bigger than 2GB, so
avoid using it unless you know what you're doing. In some cases, it can even
report that the file doesn't exist even though the file is there.</font></p>
<p>&nbsp;</p>
<p><font face="Verdana"><b>Use VB's Application LogEvent to track your
application status:<br>
</b>&nbsp;&nbsp;&nbsp; VB provides a good way to log events to a file or to
Windows NT/XP Application Log.<br>
&nbsp;&nbsp;&nbsp; Note that this will only in the compiled file. No event will
be logged in IDE mode.<br>
&nbsp; <br>
&nbsp;&nbsp;&nbsp; How to log events to an external file:<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
<font face="MS Sans Serif">App.StartLogging &quot;c:\test.log&quot;, vbLogToFile ' or
VbLogOverwrite<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
App.LogEvent &quot;Hello world&quot;, vbLogEventTypeError<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
App.LogEvent &quot;Hello world&quot;, vbLogEventTypeWarning<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
App.LogEvent &quot;Hello world&quot;, vbLogEventTypeInformation<br>
<br>
</font><font face="Verdana">&nbsp;&nbsp;&nbsp; How to log events to NT
Application Log:<br>
</font><font face="MS Sans Serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
App.StartLogging &quot;My Application&quot;, vbLogToNT<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
App.LogEvent &quot;Hello world&quot;, vbLogEventTypeError<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
App.LogEvent &quot;Hello world&quot;, vbLogEventTypeWarning<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
App.LogEvent &quot;Hello world&quot;, vbLogEventTypeInformation<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><font face="Verdana">One great thing about
this is that your other calls (ex: DLL's and OCX's) can use the logevent to the
same log as the main EXE file.<br>
&nbsp;&nbsp;&nbsp; This is great to debug a applications or communication
services. <br>
&nbsp;&nbsp;&nbsp; Just remember not to log TOO MUCH or else it will be filled
with irrelevent data.</font></p>
<p>&nbsp;</p>
<p><font face="Verdana"><b>Some functions that most people are unware of<br>
</b>&nbsp;&nbsp;&nbsp; How to convert a byte array to a string: <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><font face="MS Sans Serif">
MyString = strconv(MyByteArray, vbUnicode)</font></p>
<p><font face="Verdana">&nbsp;&nbsp;&nbsp; How to convert a string to a byte
array: <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><font face="MS Sans Serif">
MyByteArray = StrConv(MyString, vbFromUnicode)</font></p>
<p><font face="Verdana">&nbsp;&nbsp;&nbsp; How to add controls to your forms in
<b><u>run mode</u></b>: <br>
&nbsp;&nbsp; </font><font face="MS Sans Serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Private WithEvents Text1 As TextBox ' So you can also have events<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Sub
AddTextBox()<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Set Text1 =
Me.Controls.Add(&quot;VB.TextBox&quot;, &quot;Text1&quot;)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ' Now we have
the control, just as if it was added on design mode.<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Text1.Move 0,
0, 500, 100<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Text1.Visible
= True<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End With<br>
</font><font face="Verdana">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; This also
works with other controls (ex: Winsock) as long as the control is present in
your project's VB toolbox.<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; In this case you also need to remove
the check 'Remove information about unused ActiveX' in <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; the VB compilation options unless if
you have at least one control present in any of your project forms.</font></p>
<p><font face="Verdana">&nbsp;&nbsp;&nbsp; How can you check the number of forms
currently loaded:<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><font face="MS Sans Serif">
NumberOfFormsLoaded = vb.Forms.Count</font><font face="Verdana"> </font> </p>
<p><font face="Verdana">&nbsp;&nbsp;&nbsp; How to unload all forms in a MDI
project safely:<br>
&nbsp;&nbsp; </font><font face="MS Sans Serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
do until vb.Forms.Count &lt;=0 <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; unload
vb.forms(0)<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; loop<br>
</font><font face="Verdana">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Note: if
you have in your form's QueryUnload or Unload events, the possibility of a
cancel operation this code won't work properly.</font></p>
<p><font face="Verdana">&nbsp;&nbsp;&nbsp; Check if you can show a non-Modal
form before you try it to show it.<br>
&nbsp;&nbsp;&nbsp; </font><font face="MS Sans Serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
MyForm.Show (1+App.NonModalAllowed)</font><font face="Verdana"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; This will automatically show your
form in Modal mode is a previous form is already in that mode...<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Note that if you try to show a
non-Modal form when a Modal form is visible, VB will stop the execution with a
run time error, crashing your application entirely, so this is safe to use
ensuring that no &quot;Non-Modal&quot; run time error occurs.</font></p>
<p><font face="Verdana">&nbsp;&nbsp;&nbsp; How to hide your application from the
Task Manager's &quot;Applications&quot; tab (however it will be visible in the &quot;Processes&quot;
tab):<br>
&nbsp;&nbsp;&nbsp; </font><font face="MS Sans Serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
App.TaskVisible = False</font></p>
<p><font face="Verdana">Cheers<br>
<br>
// FCLage</font> </p>
2006-03-07

