Dim fso, outFile
Set fso = CreateObject("Scripting.FileSystemObject")

set xml = CreateObject("Chilkat.Xml")
xml.LoadXmlFile ("c:\tmp\games_agsn_xml.php.xml")

xml.FindChild2 ("Languages")
xml.FindChild2 ("Language")
xml.FindChild2 ("Catalogs")
xml.FindChild2 ("Catalog")
xml.FindChild2 ("Items")

n = xml.NumChildrenHavingTag("Item")

For i = 0 To n - 1

	Set game = xml.GetNthChildWithTag("Item",i)

	game.FindChild2 ("Properties")
	str = game.GetChildContentByIndex(1)'get release date
	gameDate= CDate (str)
	dateUpd = now-10'Date compare with
	
	if gameDate > dateUpd then 
		
		Set outFile = fso.CreateTextFile("c:\tmp\output"&i&".txt", True)
		'get description
		j = game.NumChildrenHavingTag("Property")
		for k = 0 to j-1
			
			set Prop = game.GetNthChildWithTag("Property",k)
			if Prop.HasAttrWithValue ("Code", "Description450") then 
				strDesc450 = game.GetChildContentByIndex(k)
			end if			
			set Prop = Nothing
			
		next
		
		game.getparent2
		game.FindChild2 ("Images")
		j = game.NumChildrenHavingTag("Image")
		for k = 0 to j-1
			
			set Logo = game.GetNthChildWithTag("Image",k)
			if Logo.HasAttrWithValue ("Type", "logo190x140") then 
				strLogo = game.GetChildContentByIndex(k)
			end if			
			set Logo = Nothing
			
		next
				
		game.getparent2
		strName = game.GetChildContent("Name")
		
		game.FindChild2 ("Files")
		strExe = game.GetChildContent("File")
		
		game.getparent2
		game.FindChild2 ("Screenshots")
		j = game.NumChildrenHavingTag("Screenshot")
		ReDim arrScreenshot (j)
		
		for k = 1 to j step 2
			arrScreenshot (k) = game.GetChildContentByIndex(k-1)
			arrScreenshot (k+1) = game.GetChildContentByIndex(k)
		next 
		
		'compose file
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
		for k = 1 to j step 2
			outfile.Write "<a href="&""""&arrScreenshot(k+1)&""""&" target="&""""&_
			"blank"&""""&"> <img src="&""""&arrScreenshot(k)&""""&" alt="&""""&_
			"мини-игра "&strName&""""&"> </a>"
		next
		outfile.WriteLine "</div>"
		outfile.Close
		
	end if

next
