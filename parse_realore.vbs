Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")

set xml = CreateObject("Chilkat.Xml")
xml.LoadXmlFile ("c:\tmp\realrore.xml")

xml.FindChild2 ("games")

n = xml.NumChildrenHavingTag("game")

For i = 0 To n - 1

	Set game = xml.GetNthChildWithTag("game",i)

	Set outFile = fso.CreateTextFile("c:\tmp\output"&i&".txt", True)
		
	'get description
	strDesc450 = game.GetChildContent("descr_full")
		
	strLogo = game.GetChildContent("pic_122x110")

	strName = game.GetChildContent("name")

	strExe = game.GetChildContent("download_url")

	ReDim arrScreenshot (6)
	arrScreenshot (1) = game.GetChildContent("preview1")
	arrScreenshot (2) = game.GetChildContent("screen1")
	arrScreenshot (3) = game.GetChildContent("preview2")
	arrScreenshot (4) = game.GetChildContent("screen2")
	arrScreenshot (5) = game.GetChildContent("preview3")
	arrScreenshot (6) = game.GetChildContent("screen3")
		
'======= compose file
		outfile.Write (strName)
		outfile.WriteLine ("<div class="&"layer1"&"><div align="""&"center"&""""&_
		"><img src="&""""&strLogo&""""& " align="""&"middle"&""""&" alt="&_
		""""&strName&""""&"/> </div></div>")
		outfile.WriteLine strDesc450
		outfile.WriteLine "<br><noindex><a rel=""nofollow"" href="&""""&strExe&""""&"><img align="&""""&"right"&_
		""""&" src="&""""&"/images/down.gif"&""""&" alt="&""""&"Скачать мини-игру "&_
		strName&""""&"/></a></noindex>"
		outfile.WriteLine "<br><h3>Картинки из мини-игры "&strName&":"&"</h3><br>"
		outfile.WriteLine "<div align="&""""&"center"&""""&">"
		'for k = 1 to j step 2
		for k = 1 to 6 step 2
			outfile.Write "<a href="&""""&arrScreenshot(k+1)&""""&" target="&""""&_
			"blank"&""""&"> <img src="&""""&arrScreenshot(k)&""""&" alt="&""""&_
			"мини-игра "&strName&""""&"> </a>"
		next
		outfile.WriteLine "</div>"
		outfile.Close

	set game = nothing
next
