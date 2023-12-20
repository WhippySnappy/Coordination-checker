#Requires Autohotkey v1.1+
#SingleInstance Force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#MaxMem 172

; ******************************************************************************************************************
; SCRAPES CoordCheck FOR VARIABLES & CHANGE #s and SCRAPES INTERSECTION TIMING CARDS FOR NAMES & CHANGE #s THEN COMPARES
; NOTE... might want to try Sift search to see if there are better results...although, this one is pretty decent
; ******************************************************************************************************************


VarList := "", IntList := "", cardList := "", finalList := "Format of this list is that program variables are on first line & timing cards are on 2nd`n`n"
varVarFile := "Variable from Coord List.txt", varNameFile := "Name from Coord List.txt", cardNameFile := "Name from Card List.txt", finalFile := "Variable to Card tester.txt"
zeroMatch := "***************ZERO MATCHES***************`n", oneMatch := "***************ONE MATCH***************`n", multMatch := "***************TWO PLUS MATCHES***************`n"
cardNames := [], coord := {}, card := {}

if FileExist("..\CoordCheck.html")
	MsgBox, UP ONE FOLDER
; FileSelectFile, coordFile, , ..\CoordCheck.html, --------------Choose File to Scrape--------------, *.html
if FileExist("CoordCheck.html")
	MsgBox, UP ONE FOLDER
; FileSelectFile, coordFile, , ..\CoordCheck.html, --------------Choose File to Scrape--------------, *.html
if (ErrorLevel)
{
	MsgBox, 0, Okie Dokie, CANCELED!
	ExitApp
}
FileDelete, %varNameFile%

Loop, read, %coordFile%
{
	if RegExMatch(A_LoopReadLine, "^\s+ChangeDisp")
	{
		if RegExMatch(A_LoopReadLine, "i)explainer") ;ignore legend/explainer section
			continue
		tempLine := RegExReplace(A_LoopReadLine, "^\s+ChangeDisp\('_?(.+)_change\D+(\d+).+", "$1 - $2")
		IntLine := RegExReplace(tempLine, "([a-z])([A-Z])", "$1 $2")
		IntLine := RegExReplace(IntLine, "([a-z])(\d)", "$1 $2")
		IntLine := RegExReplace(IntLine, "(\d)([A-Z])", "$1 $2")
		IntLine := RegExReplace(IntLine, "_", " ")
		IntList .= IntLine . "`n"
	}
}
FileAppend, %IntList%, %varNameFile% ; this file is for comparing to timing cards & see if change #s match

EnvGet, OutputVar, OneDriveCommercial ;get work folder location
OutputVar := OutputVar . "\Paperless\Timing Cards TSS\Intersection Timing Cards"
FileSelectFolder, TimingCards, *%OutputVar%, 0, Select Timing Card Folder ;should be at correct spot
TimingCards := TimingCards . "\*.pdf"
if (ErrorLevel)
{
	MsgBox, 0, Okie Dokie, CANCELED!
	ExitApp
}
Loop, files, %TimingCards%
{
	RegExMatch(A_LoopFileName, "i)(.+?_)(Ch[_\s]*)(\d+)\s*(\.pdf)", tempName)
	RegExMatch(cardNames[cardNames.length()], "i)(.+?_)(Ch[_\s]*)(\d+)\s*(\.pdf)", lastName)
	if (tempName1 = lastName1 && tempName3 > lastName3) ; same name later change#
	{
		cardNames.pop() ; get rid of earlier change#
		cardNames.push(A_LoopFileName) ; put in later change#
	}
	else if (tempName1 = lastName1 && tempName3 < lastName3) ; same name earlier change#...ignore
		continue
	else ; must not have the same name as the last one
	{
		cardNames.push(A_LoopFileName)
	}

}
for key in cardNames ; put array in variable so file is written once
	cardList .= cardNames[key] . "`n"
FileDelete, %cardNameFile%
FileAppend, %cardList%, %cardNameFile%


Loop, Read, %varNameFile%
{
	RegExMatch(A_LoopReadLine, "[\w\s]+? - (\d+)", coordChange) ; get change# in coordChange1
	coordV := RegExReplace(A_LoopReadLine, " - \d+") ; now v is just the name
	coordV := RegExReplace(coordV, "\s{2,}", " ") ; all spaces single
	coordNameArray := StrSplit(coordV, A_Space)
	for k, v in coordNameArray
	{
		coordMidblockTest := 0
		if (v = "midblock")
		{
			coordMidblockTest := 1
			break
		}
	}
	coord.push({full: A_LoopReadLine, names: coordNameArray, change: coordChange1, midblockTest: coordMidblockTest})
}
Loop, Read, %cardNameFile%
{
	tempName := RegExReplace(A_LoopReadLine, "i)(.+?_)(Ch[_\s]?)(\d+[a-zA-Z]*)(\.pdf)", "$1 - $3") ; get just "name - #" so it's like list scraped from coord program
	tempName := RegExReplace(tempName, "_", " ") ; underscore to space
	tempName := RegExReplace(tempName, "\s{2,}", " ") ; all spaces single
	RegExMatch(tempName, "[\w\s]+? - (\d+[a-zA-Z]*)", cardChange) ; get change# in cardChange1
	cardV := RegExReplace(tempName, " - \d+") ; now v is just the name
	cardNameArray := StrSplit(cardV, A_Space)
	for k, v in cardNameArray
	{
		cardMidblockTest := 0
		if (v = "midblock")
		{
			cardMidblockTest := 1
			break
		}
	}
	card.push({full: tempName, names: cardNameArray, change: cardChange1, midblockTest: cardMidblockTest})
}


for coordK in coord ; each intersection in coordination
{
	p := coordK/coord.length()*100 ; get progress %
	Progress, %p%,,, WORKING.................STATUS ; display progress
	matchList := "", zeroList := "", totalMatch := 0 ; flag for when there is a complete match to coord names
	for cardK in card ; each intersection in timing cards
	{
		counter := 0 ; reset counter for name matches
		; check each word from coord against each word from card
		for coordkey, coordval in coord[coordK].names ; each word in current coordination intersection
		{
			;~ cardCheck := ""
			if (coordval = "Ofarrell") ; special case
				coordval := "o'farrell"
			for cardkey, cardval in card[cardK].names ; each word in current timing card intersection
			{
				if (coordval = cardval) ; words match
				{
					counter++ ; tally of matches
					break ; move on so we don't match coord-val to multiple card-val... why keep going if there's a match
				}
			}
		}

		if (counter >= coord[coordK].names.length() && coord[coordK].midblockTest = 0 && card[cardK].midblockTest = 1) ; if variable doesn't contains "midblock" AND timing card does...put in unmatched list
		{
			zeroList := coord[coordK].full . "`n`n"
		}
		else if (counter >= coord[coordK].names.length()) ;if # of matches = # of entry words then all of the variable names are present in card names...but not necessarily the other way
		{
			if (coord[coordK].names.length() != card[cardK].names.length()) ; are the 2 names arrays different lengths
				matchList .= "---MAYBE NOT A MATCH---`n"
			; check for different change #s ... ones that need updating & which ones are OK
			if (coord[coordK].change = card[cardK].change)
				matchList .= coord[coordK].full . "`n" . card[cardK].full . "`n`n"
			else
				matchList .= "---DIFFERENT CHANGE #---`n" . coord[coordK].full . "`n" . card[cardK].full . "`n`n"
			totalMatch += 1, counter := 0
		}
		else
			zeroList := coord[coordK].full . "`n`n"
	}
	if (totalMatch = 0) ; make an UNKNOWN list
		zeroMatch .= zeroList ;. "`n"
	else if (totalMatch = 1) ; list for single match
		oneMatch .= matchList ;. "`n"
	else if (totalMatch > 1) ; list for 2+ match
		if (SubStr(matchList, -1) = "`n`n") ; if there's a blank line above, go back a line & start - this is for visually grouping multiple matches
		{
			matchList := SubStr(matchList, 1, -1)
			multMatch .= "----------------------------------------------" . "`n" . matchList . "----------------------------------------------" . "`n`n"
		}
		else
			multMatch .= "----------------------------------------------" . "`n" . matchList . "----------------------------------------------" . "`n"
}

Progress, off
FileDelete, %cardNameFile%
FileDelete, %varNameFile%
FileDelete, %finalFile%
finalList .= zeroMatch . "`n`n" . oneMatch . "`n`n" . multMatch
FileAppend, %finalList%, %finalFile%

Run, %finalFile%
WinWaitActive, ahk_exe notepad++.exe, , 2
Sleep, 200
IfWinActive, , This file has been modified by another program.
	Send, y
Send, ^f
Sleep, 100
Control, Check, , Button18, ahk_exe notepad++.exe
; MsgBox, WASSUP
Sleep, 200
Send, {text}---MAYBE.+\R---DIFF.+\R.+\R.+\R|---DIFFERENT.+\R.+\R.+\R|---MAYBE.+\R.+\R.+\R
ExitApp