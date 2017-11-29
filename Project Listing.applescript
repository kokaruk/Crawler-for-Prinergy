-- =====

-- The Project to crawl through active jobs in prinergy and update infosheets

-- =====



(*

Crowler

ver. 6.5

Fix All InfoSheets in all jobs and keep data when possible

*)



-- Global variable ensures that the ProgBar subroutines can understand the fileList.


global folderList

-- Launch ProgBar.
startProgBar() of me

with timeout of (12 * 30 * 60) seconds
	tell application "Finder"
		set sourceFolder to choose folder with prompt "Please select directory."
		
		-- checks if WebRoot is mounted	
		
		if exists disk "PrinergyWebRoot" then
			
			my walkFolders(sourceFolder)
			with timeout of (12 * 60 * 60) seconds
				display dialog "Mission complete" buttons {"OK"} default button 1
			end timeout
		else
			tell application "Finder"
				try
					mount volume "smb://10.9.21.2\\{username}:password@10.9.21.2/PrinergyWebRoot"
					delay 1
				end try
				repeat until (list disks) contains "PrinergyWebRoot"
				end repeat
			end tell
			my walkFolders(sourceFolder)
			eject "PrinergyWebRoot"
			
			
			with timeout of (12 * 60 * 60) seconds
				display dialog "Mission complete" buttons {"OK"} default button 1
			end timeout
		end if
	end tell
end timeout

on walkFolders(mSourceFolder)
	
	-- presume the infosheet is correct to update the data
	
	set infoSheetNew to true
	
	
	tell application "Finder"
		set folderList to every folder of mSourceFolder
		set itemCount to (get count of items in folderList)
		set theFile0 to "PrinergyWebRoot:RUser:custom:JobInfoSheet:Default:Customer:JobInfoSheet_read_en.xslt"
		set theFile1 to "PrinergyWebRoot:RUser:custom:JobInfoSheet:Default:Staff:JobInfoSheet_read_en.xslt"
		set theFile2 to "PrinergyWebRoot:RUser:custom:JobInfoSheet:Default:Staff:JobInfoSheet_edit_en.xslt"
		set theFile3 to "PrinergyWebRoot:RUser:custom:JobInfoSheet:Default:JobInfoSheet.xml"
	end tell
	
	prepareProgBar(itemCount, 1) of me
	fadeinProgBar(1) of me
	
	repeat with i from 1 to itemCount
		
		incrementProgBar(i, itemCount, 1) of me
		
		set theFolder to item i of the folderList
		
		
		
		set folderExists to existsVersionedJobInformationSheets(theFolder) of me
		
		-- checks if job has infoSheets in it
		
		if folderExists then
			
			
			
			tell application "Finder"
				
				-- count infosheet versions
				
				try
					set infoSheetFolderList to every folder of folder "VersionedJobInformationSheets" of folder "Control" of theFolder
					set infSheetItemCount to (get count of items in infoSheetFolderList)
					my writeText(i, theFolder, folderExists)
				on error
					
					my writeText(i, theFolder, "error: no access to folder")
					
				end try
				
				
			end tell
			
			-- cycle starts here
			
			
			set infoSheetFolder to item infSheetItemCount of the infoSheetFolderList as alias
			set theFile to (((infoSheetFolder) as string) & "JobInfoSheet.xml")
			
			try
				
				tell application "Finder"
					duplicate file theFile0 to folder "Customer" of infoSheetFolder with replacing
					duplicate file theFile1 to folder "Staff" of infoSheetFolder with replacing
					duplicate file theFile2 to folder "Staff" of infoSheetFolder with replacing
				end tell
				
				my writeText(i, theFolder, "ok")
				
			on error
				
				my writeText(i, theFolder, "error duplicating")
				
			end try
			
			
			try
				-- my repairXml(theFile)
			on error
				
				try
					tell application "Finder"
						duplicate file theFile3 to folder infoSheetFolder with replacing
					end tell
				on error
					
					my writeText(i, theFolder, "error fixing xml")
					
				end try
				
			end try
			
			
			
		end if
		
	end repeat
	
	fadeoutProgBar(1) of me
	
end walkFolders
on existsVersionedJobInformationSheets(mTheFolder)
	tell application "Finder"
		set folderExists to false
		if (exists folder "VersionedJobInformationSheets" of folder "Control" of mTheFolder) then set folderExists to true
	end tell
	return folderExists
	
end existsVersionedJobInformationSheets

on writeText(i, theFolder, folderExists)
	
	set logPath to (path to desktop) as alias
	
	
	try
		set theFile to ((logPath as string) & "log.txt") as alias
	on error
		tell application "Finder" to make new file at logPath with properties {name:"log.txt"}
		set theFile to ((logPath as string) & "log.txt") as alias
	end try
	set theData to (theFolder) as alias
	set foderNname to name of (info for theData)
	set numberFile to i as text
	set textFolderExist to folderExists as text
	
	try
		open for access theFile with write permission
		write numberFile & " " & foderNname & " " & textFolderExist & return to theFile starting at eof
		close access theFile
	on error
		try
			close access theFile
		end try
	end try
	
end writeText


on repairXml(theFile)
	
	
	set theDoc to XMLOpen( theFile )
	set Root to XMLRoot( theDoc )
	set desInfo to XMLXPath (Root) with "DesignInfo"
	
	-- Create missing data
	
	set newJobName to XMLXPath Root with "DesignInfo/NewJobName"
	if (XMLExists newJobName) is false then
		XMLNewChild "<NewJobName/>" at desInfo with keep blanks
	end if
	
	set productType to XMLXPath Root with "DesignInfo/ProductType"
	if (XMLExists productType) is false then
		XMLNewChild "<ProductType/>" at desInfo with keep blanks
	end if
	
	set InternalAccountManager to XMLXPath Root with "DesignInfo/InternalAccountManager"
	if (XMLExists InternalAccountManager) is false then
		XMLNewChild "<InternalAccountManager/>" at desInfo with keep blanks
	end if
	
	set AccountManager to XMLXPath Root with "DesignInfo/AccountManager"
	if (XMLExists AccountManager) is false then
		XMLNewChild "<AccountManager/>" at desInfo with keep blanks
	end if
	
	set JobType to XMLXPath Root with "DesignInfo/JobType"
	if (XMLExists JobType) is false then
		XMLNewChild "<JobType Process='New'/>" at desInfo with keep blanks
	end if
	
	set ProofInstructions to XMLXPath Root with "DesignInfo/ProofInstructions"
	if (XMLExists ProofInstructions) is false then
		XMLNewChild "<ProofInstructions LaserColour='no' EpsonSemi='no' EpsonBond='no' Mockup='no'/>" at desInfo with keep blanks
	end if
	
	set Ticks to XMLXPath Root with "DesignInfo/Finishing/Ticks"
	set Finishing to XMLXPath Root with "DesignInfo/Finishing"
	if (XMLExists Ticks) is false then
		if (XMLExists Finishing) then
			XMLRemove Finishing
		end if
		XMLNewChild "<Finishing>
			<Ticks Saddlestitched='no' Padded='no' Staples='no' QtrBound='no' FormCut='no' RoundCornered='no' PerfectBound='no' BurstBound='no' Sewn='no' Fanapart='no' Scored='no' Folded='no' Glued='no' DTape='no' DieCut='no' GlueLinesPatterns='no' Embossed='no' Debossed='no' Hologram='no' Foiling='no' TipOnLabelCard='no' Desensitizing='no' Shrinkwrap='no'/>
			<FoldedType></FoldedType>
			<ShrinkWrapSize></ShrinkWrapSize> 
		</Finishing>" at desInfo with keep blanks
	end if
	
	-- create art charge table
	
	
	set artChargeCheck to XMLXPath Root with "ArtCharge"
	
	if (XMLExists artChargeCheck) is false then
		XMLNewChild "<ArtCharge></ArtCharge>" at Root with keep blanks
		set artCharge to XMLXPath Root with "ArtCharge"
		repeat with i from 1 to 6
			set cEevent to "<CEvent ID='" & i & "'></CEvent>" as string
			XMLNewChild cEevent at artCharge with keep blanks
			set cEventCurrentXPath to "ArtCharge/CEvent[@ID='" & i & "']" as string
			set cEventCurrent to XMLXPath Root with cEventCurrentXPath
			XMLNewChild "<Date/>" at cEventCurrent with keep blanks
			XMLNewChild "<Initial/>" at cEventCurrent with keep blanks
			XMLNewChild "<Cost/>" at cEventCurrent with keep blanks
			XMLNewChild "<Proofs/>" at cEventCurrent with keep blanks
			XMLNewChild "<NC/>" at cEventCurrent with keep blanks
			XMLNewChild "<Descr/>" at cEventCurrent with keep blanks
		end repeat
	end if
	
	set artCharge to XMLXPath Root with "ArtCharge"
	set costs to XMLXPath Root with "ArtCharge/Costs"
	set hiddenStock to XMLXPath Root with "DesignInfo/Sheet1/Paper/StockHidden"
	set designInfo to XMLXPath Root with "DesignInfo"
	
	if (XMLExists costs) is false then
		XMLNewChild "<Costs></Costs>" at artCharge with keep blanks
		set costs to XMLXPath Root with "ArtCharge/Costs"
		XMLNewChild "<PreSumArt/>" at costs with keep blanks
		XMLNewChild "<PreSumNc/>" at costs with keep blanks
		XMLNewChild "<ExpCost/>" at costs with keep blanks
		XMLNewChild "<SaleSig/>" at costs with keep blanks
	end if
	
	--CTP chrges table
	
	set CtpChargeCheck to XMLXPath Root with "CtpCharge"
	
	if (XMLExists CtpChargeCheck) is false then
		XMLNewChild "<CtpCharge></CtpCharge>" at Root with keep blanks
		set CtpCharge to XMLXPath Root with "CtpCharge"
		repeat with ii from 1 to 5
			set plate to "<Plate ID='" & ii & "'></Plate>" as string
			XMLNewChild plate at CtpCharge with keep blanks
			set plateXpath to "CtpCharge/Plate[@ID='" & ii & "']" as string
			set plateCurrent to XMLXPath Root with plateXpath
			XMLNewChild "<Date/>" at plateCurrent with keep blanks
			XMLNewChild "<Initial/>" at plateCurrent with keep blanks
			XMLNewChild "<TypeQty/>" at plateCurrent with keep blanks
			XMLNewChild "<Cost/>" at plateCurrent with keep blanks
		end repeat
	end if
	
	set CtpTotal to XMLXPath Root with "CtpCharge/CtpTotal"
	if (XMLExists CtpTotal) is false then
		set CtpCharge to XMLXPath Root with "CtpCharge"
		XMLNewChild "<CtpTotal/>" at CtpCharge with keep blanks
	end if
	
	
	-- check if we need to fix hidden stock
	
	if (XMLExists hiddenStock) is false then
		
		repeat with i from 1 to 6
			set sheet to XMLXPath Root with "DesignInfo/Sheet" & i & "/Paper"
			XMLNewChild "<StockHidden/>" at sheet
			set stock to XMLXPath sheet with "Stock"
			set stockC to XMLGetText stock
			set stockHidden to XMLXPath sheet with "StockHidden"
			XMLSetText stockHidden to stockC
		end repeat
	end if
	
	
	set oldJobName to XMLXPath Root with "DesignInfo/OldJobName"
	if (XMLExists oldJobName) is false then
		XMLNewChild "<OldJobName/>" at designInfo with keep blanks
	end if
	
	set otherStock to XMLXPath Root with "DesignInfo/Sheet1/Paper/OtherStock"
	
	if (XMLExists otherStock) is false then
		repeat with i from 1 to 6
			set sheet to XMLXPath Root with "DesignInfo/Sheet" & i & "/Paper"
			XMLNewChild "<OtherStock/>" at sheet
		end repeat
	end if
	
	
	-- New attributes for "Finishing / Ticks" 
	set attribPath to XMLXPath Root with "DesignInfo/Finishing/Ticks"
	set attribValue to "no"
	my newAttrib("Numbering", attribPath, Root, attribValue)
	my newAttrib("Spotuv", attribPath, Root, attribValue)
	my newAttrib("Other", attribPath, Root, attribValue)
	
	-- new child node "Comments"  at "Finishing" 
	set nodePathString to "DesignInfo/Finishing"
	my newNode("Comments", nodePathString, Root)
	
	-- <ProofRequired> attributes remove "printer" add "Hiddenproof" 
	set remAttr to XMLXPath Root with "DesignInfo/ProofRequired/@Printer"
	set attribPath to XMLXPath Root with "DesignInfo/ProofRequired"
	set attribValue to ""
	if (XMLExists remAttr) then
		XMLRemoveAttribute attribPath name "Printer"
	end if
	my newAttrib("Hiddenproof", attribPath, Root, attribValue)
	
	-- New Node "ProofingDevice" with 'HiddenDevice' attribute 
	set nodePathString to "DesignInfo"
	my newNode("ProofingDevice", nodePathString, Root)
	set attribPath to XMLXPath Root with "DesignInfo/ProofingDevice"
	set attribValue to ""
	my newAttrib("HiddenDevice", attribPath, Root, attribValue)
	
	-- New Node 'InfoSheetVer @ DesignInfo'
	set nodePathString to "DesignInfo"
	my newNode("InfoSheetVer", nodePathString, Root)
	-- version
	set ver to XMLXPath Root with "DesignInfo/InfoSheetVer"
	XMLSetXML ver to "8.00"
	
	-- New nodes TPI1 TPI2 at Sheets
	repeat with i from 1 to 6
		set nodePathString to "DesignInfo/Sheet" & i & "/PerfSlit"
		my newNode("TPI1", nodePathString, Root)
		my newNode("TPI2", nodePathString, Root)
	end repeat
	
	-- New <PdfExport OnSite="no" Printdirect="no" />
	set nodePathString to "DesignInfo"
	my newNode("PdfExport", nodePathString, Root)
	set attribPath to XMLXPath Root with "DesignInfo/PdfExport"
	set attribValue to "no"
	my newAttrib("OnSite", attribPath, Root, attribValue)
	my newAttrib("Printdirect", attribPath, Root, attribValue)
	
	-- New <Printdirect PrintReady="no" OnlinePrev="no" CustID="" ItemNo="" Nonprintable="no"/>
	set nodePathString to "DesignInfo"
	my newNode("Printdirect", nodePathString, Root)
	set attribPath to XMLXPath Root with "DesignInfo/Printdirect"
	set attribValue to "no"
	my newAttrib("PrintReady", attribPath, Root, attribValue)
	my newAttrib("OnlinePrev", attribPath, Root, attribValue)
	my newAttrib("Nonprintable", attribPath, Root, attribValue)
	set attribValue to ""
	my newAttrib("CustID", attribPath, Root, attribValue)
	my newAttrib("ItemNo", attribPath, Root, attribValue)
	
	-- New <Imposition Impose="no"/>
	set nodePathString to "DesignInfo"
	my newNode("Imposition", nodePathString, Root)
	set attribPath to XMLXPath Root with "DesignInfo/Imposition"
	set attribValue to "no"
	my newAttrib("Impose", attribPath, Root, attribValue)
	
	-- New <SelectivePrintPages SelectivePrint="no" SelectivePrintHidden="" SelectivePrintPagesHidden="" SelectiveImpoSides="Front"/>
	set nodePathString to "DesignInfo"
	my newNode("SelectivePrintPages", nodePathString, Root)
	set attribPath to XMLXPath Root with "DesignInfo/SelectivePrintPages"
	
	set attribValue to "no"
	my newAttrib("SelectivePrint", attribPath, Root, attribValue)
	set attribValue to ""
	my newAttrib("SelectivePrintHidden", attribPath, Root, attribValue)
	set attribValue to ""
	my newAttrib("SelectivePrintPagesHidden", attribPath, Root, attribValue)
	set attribValue to "Front"
	my newAttrib("SelectiveImpoSides", attribPath, Root, attribValue)
	
	
	-- New <ProofType ProofOption="LoosePage"/>
	set nodePathString to "DesignInfo"
	my newNode("ProofType", nodePathString, Root)
	set attribPath to XMLXPath Root with "DesignInfo/ProofType"
	set attribValue to "LoosePage"
	my newAttrib("ProofOption", attribPath, Root, attribValue)
	
	
	-- New Node 'CopyJobName' @ 'DesignInfo'
	set nodePathString to "DesignInfo"
	my newNode("CopyJobName", nodePathString, Root)
	
	-- New node 'PartTitle' at Sheets
	repeat with i from 1 to 6
		set nodePathString to "DesignInfo/Sheet" & i
		my newNode("PartTitle", nodePathString, Root)
	end repeat
	
	XMLSave theDoc in file theFile (*without formatting*)
	
end repairXml

on newNode(newName, nodePathString, Root)
	
	set nodePath to XMLXPath Root with nodePathString
	set newPath to XMLXPath Root with nodePathString & "/" & newName
	
	if (not (XMLExists newPath)) then
		XMLNewChild "<" & newName & "/>" at nodePath
	end if
	
end newNode


on newAttrib(AttribName, attribPath, Root, attribValue)
	
	set numb to XMLXPath Root with "DesignInfo/Finishing/Ticks/@" & AttribName
	
	if (not (XMLExists numb)) then
		
		XMLSetAttribute attribPath name AttribName to attribValue
		
	end if
	
end newAttrib


stopProgBar() of me


-- SUBROUTINES BELOW
-- Copy from here to the end of this document.
-- Then, simply paste into the end of your own script.
-- This will give you all of ProgBar's syntax necessary for manipulation.
-- Alternatively, you could paste the code below into a script file of its own and load it as a library.

-- Prepare progress bar subroutine.
on prepareProgBar(someMaxCount, windowName)
	tell application "ProgBar"
		set background color of window windowName to {65535, 65535, 65535}
		set has shadow of window windowName to true
		set level of window windowName to item 7 of {0, 3, 8, 20, 24, 101, 1001}
		set title of window windowName to ""
		set content of progress indicator 1 of window windowName to 0
		set minimum value of progress indicator 1 of window windowName to 0
		set maximum value of progress indicator 1 of window windowName to someMaxCount
	end tell
end prepareProgBar

-- Increment progress bar subroutine.
on incrementProgBar(itemNumber, someMaxCount, windowName)
	tell application "ProgBar"
		set title of window windowName to "Processing " & itemNumber & " of " & someMaxCount & " - " & (item itemNumber of folderList)
		set content of progress indicator 1 of window windowName to itemNumber
	end tell
end incrementProgBar

-- Fade in a progress bar window.
on fadeinProgBar(windowName)
	tell application "ProgBar"
		center window windowName
		set alpha value of window windowName to 0
		set visible of window windowName to true
		set fadeValue to 0.1
		repeat with i from 0 to 9
			set alpha value of window windowName to fadeValue
			set fadeValue to fadeValue + 0.1
		end repeat
		start progress indicator 1 of window windowName with uses threaded animation
	end tell
end fadeinProgBar

-- Fade out a progress bar window.
on fadeoutProgBar(windowName)
	tell application "ProgBar"
		stop progress indicator 1 of window windowName with uses threaded animation
		set fadeValue to 0.9
		repeat with i from 1 to 9
			set alpha value of window windowName to fadeValue
			set fadeValue to fadeValue - 0.1
		end repeat
		set visible of window windowName to false
	end tell
end fadeoutProgBar

-- Show progress bar window.
on showProgBar(windowName)
	tell application "ProgBar"
		center window windowName
		set visible of window windowName to true
		start progress indicator 1 of window windowName with uses threaded animation
	end tell
end showProgBar

-- Hide progress bar window.
on hideProgBar(windowName)
	tell application "ProgBar"
		stop progress indicator 1 of window windowName with uses threaded animation
		set visible of window windowName to false
	end tell
end hideProgBar

-- Enable 'barber pole' behavior of a progress bar.
on barberPole(windowName)
	tell application "ProgBar"
		set indeterminate of progress indicator 1 of window windowName to true
	end tell
end barberPole

-- Disable 'barber pole' behavior of a progress bar.
on killBarberPole(windowName)
	tell application "ProgBar"
		set indeterminate of progress indicator 1 of window windowName to false
	end tell
end killBarberPole

-- Launch ProgBar.
on startProgBar()
	tell application "ProgBar"
		launch
	end tell
end startProgBar

-- Quit ProgBar.
on stopProgBar()
	tell application "ProgBar"
		quit
	end tell
end stopProgBar


