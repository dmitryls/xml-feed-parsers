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
	dateUpd = now-14'Date compare with
	
	if gameDate > dateUpd then 
		bOLexists = false
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
		
		'get logo
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
		
		'get name		
		game.getparent2
		strName = game.GetChildContent("Name")
		
		'get file name
		game.FindChild2 ("Files")
		strExe = game.GetChildContent("File")
		
		'get screenshots
		game.getparent2
		game.FindChild2 ("Screenshots")
		j = game.NumChildrenHavingTag("Screenshot")
		ReDim arrScreenshot (j)
		
		'get screenshots
		for k = 1 to j step 2
			arrScreenshot (k) = game.GetChildContentByIndex(k-1)
			arrScreenshot (k+1) = game.GetChildContentByIndex(k)
		next 
		
		'get OL version if exists
		game.getparent2
		game.FindChild2 ("RelatedItems")
		'check if exists
		if game.HasChildWithTag("RelatedItemCatalog") then
			bOLexists = true
			game.FindChild2 ("RelatedItemCatalog")
			game.FindChild2 ("RelatedItem")
			intOnlineID = game.GetAttributeValue (0)
			'find OL
			set xmlOL = CreateObject("Chilkat.Xml")
			xmlOL.LoadXmlFile ("c:\tmp\games_agsn_xml.php.xml")
			xmlOL.GetRoot2
			xmlOL.FindChild2 ("Languages")
			xmlOL.FindChild2 ("Language")
			xmlOL.FindChild2 ("Catalogs")
			xmlOL.FindChild2 ("Catalog")
			xmlOL.NextSibling2
			xmlOL.FindChild2 ("Items")
			j1 = xmlOL.NumChildrenHavingTag("Item")
			for k = 0 to j1-1
				set OLGame = xmlOL.GetNthChildWithTag("Item",k)
				if OLGame.HasAttrWithValue ("ID", intOnlineID) then 
					OLGame.FindChild2("Properties")
					
					m = OLgame.NumChildrenHavingTag("Property")
					for p = 0 to m-1
			
						set Prop = OLgame.GetNthChildWithTag("Property",p)
						if Prop.HasAttrWithValue ("Code", "Embed") then 
							strOLhtml = OLgame.GetChildContentByIndex(p)
							'wsh.Echo strOLhtml
							
							Set outOLFile = fso.CreateTextFile("c:\tmp\output"&i&"_online.html", True)
							outOLFile.WriteLine ("<html>")
							outOLFile.WriteLine ("<head>")
							'outOLFile.WriteLine ("<title>"&strName)
							outOLFile.WriteLine ("<body>")
							outOLFile.WriteLine ("<div align="&"""center"""&">")
							outOLFile.Write (strOLhtml)
							outOLFile.WriteLine ("</div>")
							outOLFile.WriteLine ("</html>")
							
							strOLGamePage = "/Olgames/output"&i&"_online.html"
							Set outOLFile = Nothing
						end if			
						set Prop = Nothing
			
					next
										 
				end if			
				set Prop = Nothing
			next
			
		end if
		
		
		'compose file
		outfile.Write (strName)
		outfile.WriteLine ("<div class="&"layer1"&"><div align="""&"center"&""""&_
		"><img src="&""""&strLogo&""""& " align="""&"middle"&""""&" alt="&_
		""""&strName&""""&"/> </div></div>")
		outfile.WriteLine strDesc450
		
		outfile.WriteLine "<table><tr><td>"
		
		outfile.WriteLine "<br><noindex><a rel=""nofollow"" href="&""""&strExe&""""&"><img align="&""""&"right"&_
		""""&" src="&""""&"/images/down.gif"&""""&" alt="&""""&"Скачать мини-игру "&_
		strName&""""&"/></a></noindex>"
		
		outfile.WriteLine "</td></tr>"
		
		if bOLexists = true then 
			outfile.WriteLine "<tr><td>"
			outfile.WriteLine "Также <strong>в игру "&strName&" можно играть не скачивая</strong>. Для этого просто нажми Играть онлайн"
			outfile.WriteLine "<noindex><a rel=""nofollow"" target=""_blank"" href="&""""&strOLGamePage&""""&"><img align="&""""&"right"&_
			""""&" src="&""""&"/images/play_online.gif"&""""&" alt="&""""&"Играть онлайн в "&_
			strName&""""&"/></a></noindex>"
			outfile.WriteLine "</td></tr>"
		end if
		outfile.WriteLine "</table>"
		
		outfile.WriteLine "<br><h3>Картинки из мини-игры "&strName&":"&"</h3><br>"
		outfile.WriteLine "<div align="&""""&"center"&""""&">"
		for k1 = 1 to j step 2
			outfile.Write "<a href="&""""&arrScreenshot(k1+1)&""""&" target="&""""&_
			"blank"&""""&"> <img src="&""""&arrScreenshot(k1)&""""&" alt="&""""&_
			"мини-игра "&strName&""""&"> </a>"
		next
		outfile.WriteLine "</div>"
		outfile.Close
		
	end if

next
