#Requires Autohotkey v2.0
#SingleInstance Force
; V1toV2: Removed #NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ #Warn  ; Enable warnings to assist with detecting common errors.
SendMode("Input")  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir(A_ScriptDir)  ; Ensures a consistent starting directory.
; V1toV2: Removed #MaxMem 172

; ******************************************************************************************************************
; SCRAPES CoordCheck FOR VARIABLES & CHANGE #s and SCRAPES INTERSECTION TIMING CARDS FOR NAMES & CHANGE #s THEN COMPARES
; NOTE... might want to try Sift search to see if there are better results...although, this one is pretty decent
; ******************************************************************************************************************

VarList := "", IntList := "", cardList := "", finalList := "Format of this list is that program variables are on first line & timing cards are on 2nd`n`n"
varVarFile := "Variable from Coord List.txt", varNameFile := "Name from Coord List.txt", cardNameFile := "Name from Card List.txt", finalFile := "Variable to Card tester.txt"
zeroMatch := "***************ZERO MATCHES***************`n", oneMatch := "***************ONE MATCH***************`n", multMatch := "***************TWO PLUS MATCHES***************`n"
cardNames := [], coord := [], card := [], FileCount := 0, progression := 0, coordFileLength := 0, cardFileLength := 0, coordFile := ""
coord.full := []
coord.names := []
coord.change := []
coord.midblockTest := []
card.full := []
card.names := []
card.change := []
card.midblockTest := []

if FileExist(A_ScriptDir "..\CoordCheck.html") ;if scraper is in subfolder of CoordCheck
{
    FileCount++
    coordFile := FileSelect("", A_ScriptDir "..\CoordCheck.html", "--------------Choose File to Scrape--------------", "*.html")
}
 else if FileExist(A_ScriptDir "\CoordCheck.html")
{
    FileCount++
    coordFile := FileSelect("", A_ScriptDir "\CoordCheck.html", "--------------Choose File to Scrape--------------", "*.html")
}
if coordFile = ""
{
    MsgBox("CANCELED!", "Okie Dokie", "T5")
    ExitApp()
}
if (FileCount = 0)
{
    MsgBox("Can't find the file!`nWill exit now", "File Error", "T5")
    ExitApp()
}

loop read, coordFile ;scrape CoordCheck
{
    if RegExMatch(A_LoopReadLine, "^\s+ChangeDisp")
    {
        if RegExMatch(A_LoopReadLine, "i)explainer") ;ignore legend/explainer section
            continue
        tempLine := RegExReplace(A_LoopReadLine, "^\s+ChangeDisp\('_?(.+)_change\D+(\d+).+", "$1 -- $2")
        IntLine := RegExReplace(tempLine, "([a-z])([A-Z])", "$1 $2") ;separate CamelCase
        IntLine := RegExReplace(IntLine, "([a-z])(\d)", "$1 $2") ;separate numbers from letters
        IntLine := RegExReplace(IntLine, "(\d)([A-Z])", "$1 $2")
        IntLine := RegExReplace(IntLine, "_", " ")
        IntList .= IntLine . "`n"
        coordFileLength++
    }
}

TimeCardFolder := EnvGet("OneDriveCommercial") ;get work folder location
TimeCardFolder := TimeCardFolder . "\Archives\"
TimingCards := DirSelect("*" TimeCardFolder, 0, "Select Timing Card Folder") ;should be at correct spot
if TimingCards = ""
{
    MsgBox("CANCELED!", "Okie Dokie", "T5")
    ExitApp()
}

MyGui := Gui()
MyGui.BackColor := "black"
MyGui.Opt("+AlwaysOnTop -Caption +Disabled -SysMenu +ToolWindow")
MyProgress := MyGui.Add("Progress", "cRed w200 h20")
MyText := MyGui.Add("Text", "cSilver Center wp")
MyGui.Show() ; display progress
MyGui.Move(, 10)

StartTime := A_TickCount

loop files, TimingCards "\*.pdf", "R" ; go thru all the pdf filenames
{
    if !(RegExMatch(A_LoopFileDir, "Time Card")) ; only look at Time Card files
        continue
    tempName := RegExReplace(A_LoopFileName, "[_\s\-]|\.(?!pdf)", " ")
    tempName := RegExReplace(tempName, "i)(ch)(\d+)(\.pdf)", "$1 $2$3") ; CH##.pdf to CH ##.pdf
    tempName := RegExReplace(tempName, "\s*&\s*", " ")
    tempName := RegExReplace(tempName, "\s{2,}", " ")
    if RegExMatch(tempName, "i) ch[\s_]?\d+") = 0 ;if there isn't a change# just move on
        continue
    RegExMatch(tempName, "i)(.+?[_\s])(Ch[_\s]*)(\d+).*(\.?pdf)", &tempName) ; current filename
    if cardNames.Length = 0
    {
        cardNames.push(tempName[]) ; AHK v2 needs [] or [0] as the entire match. could pick individual matches if wanted
        continue
    }
    RegExMatch(cardNames[cardNames.Length], "i)(.+?[_\s])(Ch[_\s]*)(\d+).*(\.pdf)", &lastName) ;last name that's already in array
    ;had to use .* instead of \s* after "ch ##" because of non-standard naming in files
    if (tempName[1] = lastName[1] && tempName[3] > lastName[3]) ; same name later change#
    {
        cardNames.pop() ; get rid of earlier change#
        cardNames.push(tempName[]) ; put in later change#
        cardFileLength++
    }
    else if (tempName[1] = lastName[1] && tempName[3] < lastName[3]) ; same name earlier change#...ignore
        continue
    else ; must not have the same name as the last one
    {
        cardNames.push(tempName[])
        cardFileLength++
    }

}
for key, value in cardNames ; put array in variable
{
    cardList .= cardNames[key] "`n"
}

loop parse, IntList, "`n", "`r" ;parse thru coordCheck results
{
    RegExMatch(A_LoopField, "[\w\s]+? -- (\d+)", &coordChange) ; get change# in coordChange1
    coordV := RegExReplace(A_LoopField, " -- \d+") ; now v is just the name
    coordV := RegExReplace(coordV, "\s{2,}", " ") ; all spaces single
    coordNameArray := StrSplit(coordV, A_Space)
    for k, v in coordNameArray ;I had a reason for doing this???
    {
        coordMidblockTest := 0
        if (v = "midblock")
        {
            coordMidblockTest := 1
            break
        }
    }

    if coordV != ""
    {
        ;AHK v2 changed associative array structure. Cannot push an associative array into an object but indexed arrays can be pushed
        ; coord.push("full", A_LoopField, "names", coordNameArray, "change", coordChange[1], "midblockTest", coordMidblockTest) ;OLD WAY
        coord.full.push(A_LoopField)
        coord.names.push(coordNameArray)
        coord.change.push(coordChange[1])
        coord.midblockTest.push(coordMidblockTest)
        ; FileAppend coord.change[coord.change.Length] "`n", "TEMPSTORAGE.txt" ;DEBUG

    }
}

; for k, v in coord.names
; {
;     for kx, vx in coord.names[k]
;     {
;         FileAppend vx " ", "TEMPSTORAGE.txt"
;     }
;     FileAppend "`n", "TEMPSTORAGE.txt"
; }

loop parse, cardList, "`n", "`r" ;parse thru timing card names
{
    if InStr(A_LoopField, "berry", False)
        continue
    tempName := RegExReplace(A_LoopField, "i)(\b(on|off)\s?-?ramp\b)|exit|fwy|ext\b|\bI\b|\bus\b|\.?pdf|\(|\)") ; get rid of stuff
    tempName := RegExReplace(tempName, "_|-|\.", " ") ; underscore hyphen period to space
    tempName := RegExReplace(tempName, "\s{2,}", " ") ; all spaces single
    tempName := RegExReplace(tempName, "i)(.+?)\s(Ch[_\s]?)(\d+[a-zA-Z]*)", "$1 - $3") ; get just "name - #" so it's like list scraped from coord program
    RegExMatch(tempName, "[\w\s]+? - (\d+[a-zA-Z]*)", &cardChange) ; get change# in cardChange1
    cardV := RegExReplace(tempName, " - \d+") ; now v is just the name
    cardV := RegExReplace(cardV, "(\b\w+?\b)(.+)$1", "$1$2")
    cardNameArray := StrSplit(cardV, A_Space)
    ; FileAppend cardV "`n", "TEMPSTORAGE.txt" ;DEBUG
    for k, v in cardNameArray ;loop thru each word
    {
        cardMidblockTest := 0
        if (v = "midblock")
        {
            cardMidblockTest := 1
            break
        }
    }
    ; check if there are any duplicate words in nameArray & delete
    cardV := ""
    x := cardNameArray.Length - 1
    ; MsgBox(x)
    loop cardNameArray.Length ;
    {
        switch
        {
            case (InStr(cardV, cardNameArray[A_Index])): continue
            case (A_Index = cardNameArray.Length) && !InStr(cardV, cardNameArray[A_Index]): cardv .= cardNameArray[A_Index]
            case !InStr(cardV, cardNameArray[A_Index]): cardv .= cardNameArray[A_Index] " "
        }
    }
    ; MsgBox(cardV)
    if cardV != ""
    {
        ; card.push("full", tempName, "names", cardNameArray, "change", cardChange[1], "midblockTest", cardMidblockTest) ;OLD WAY
        card.full.push(A_LoopField)
        card.names.push(cardNameArray)
        card.change.push(cardChange[1])
        card.midblockTest.push(cardMidblockTest)
        ; FileAppend card.full[card.full.Length] "`n", "TEMPSTORAGE.txt" ;DEBUG
    }
}
; coord.Pop() ;array no longer exists on its own, now has sub-objects

for coordK, value in coord.names ; each intersection in coordCheck
{
    MyProgress.Value := Round(coordK / coord.names.Length * 100) ; get progress % and put in progress bar
    MyText.Value := "Comparing Data" ;let's know what it's working on
    matchList := "", zeroList := "", totalMatch := 0 ; flag for when there is a complete match to coord names
    for cardK, value in card.names ; each intersection in timing cards
    {
        counter := 0 ; reset counter for name matches
        ; check each word from coord against each word from card
        for coordkey, coordval in coord.names[coordK] ; each word in current coordination intersection
        {
            if (coordval = "Ofarrell") ; special case
                coordval := "o'farrell"
            for cardkey, cardval in card.names[cardK] ; each word in current timing card intersection
            {
                if (coordval = cardval) ; words match
                {
                    counter++ ; tally of matches
                    break ; move on so we don't match coord-val to multiple card-val... why keep going if there's a match
                }
            }
        }
        if (counter >= coord.names[coordK].Length && coord.midblockTest[coordK] = 0 && card.midblockTest[cardK] = 1) ; if variable doesn't contains "midblock" AND timing card does...put in unmatched list
        {
            zeroList := coord.full[coordK] . "`n`n"
        }
        else if (counter >= coord.names[coordK].Length) ;if # of matches = # of entry words then all of the variable names are present in card names...but not necessarily the other way
        {
            if (coord.names[coordK].Length != card.names[cardK].Length) ; are the 2 names arrays different lengths
                matchList .= "---MAYBE NOT A MATCH---`n"
            ; check for different change #s ... ones that need updating & which ones are OK
            if (coord.change[coordK] = card.change[cardK])
                matchList .= coord.full[coordK] . "`n" . card.full[cardK] . "`n`n"
            else
                matchList .= "---DIFFERENT CHANGE #---`n" . coord.full[coordK] . "`n" . card.full[cardK] . "`n`n"
            totalMatch += 1, counter := 0
        }
        else
            zeroList := coord.full[coordK] . "`n`n"
    }
    if (totalMatch = 0) ; make an UNKNOWN list
        zeroMatch .= zeroList ;. "`n"
    else if (totalMatch = 1) ; list for single match
        oneMatch .= matchList ;. "`n"
    else if (totalMatch > 1) ; list for 2+ match
        if (SubStr(matchList, -2) = "`n`n") ; if there's a blank line above, go back a line & start - this is for visually grouping multiple matches
        {
            matchList := SubStr(matchList, 1, -1)
            multMatch .= "----------------------------------------------" . "`n" . matchList .
                "----------------------------------------------" . "`n`n"
        }
        else
            multMatch .= "----------------------------------------------" . "`n" . matchList .
                "----------------------------------------------" . "`n"
}

MyGui.Destroy
if (FileExist(finalFile))
    FileDelete(finalFile)
finalList .= zeroMatch . "`n`n" . oneMatch . "`n`n" . multMatch
FileAppend(finalList, finalFile)

ElapsedTime := (A_TickCount - StartTime)

Run(finalFile)
ErrorLevel := !WinWaitActive("ahk_exe notepad++.exe", , 2)
Sleep(200)
if WinActive(, "This file has been modified by another program.")
    Send("y")
Sleep(200)
Send("^{Home}")
Sleep(100)
Send("^f")
Sleep(100)
WinWait("Find", , 1.5)
Send("{text}---MAYBE.+\R---DIFF.+\R.+\R.+\R|---DIFFERENT.+\R.+\R.+\R|---MAYBE.+\R.+\R.+\R")
Sleep(200)
ControlSetChecked(1, "Button19", "Find") ;regular expression radio button
Sleep(200)
ControlClick("Button23", "Find") ;Find Next button
Sleep(200)
ControlClick("Button26", "Find") ;Count button

MsgBox "********************`nComparison took...`n" ElapsedTime " milliseconds.`n********************`n" Round(ElapsedTime / 1000, 2) " seconds.`n********************",,"T5"

ExitApp()

ProgressBar(*)
{
    global
    MyProgress.Value := progression
}
