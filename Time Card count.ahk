#Requires Autohotkey v2.0
#SingleInstance Force
; V1toV2: Removed #NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ #Warn  ; Enable warnings to assist with detecting common errors.
SendMode("Input")  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir(A_ScriptDir)  ; Ensures a consistent starting directory.

if FileExist("Card Count.csv")
    FileDelete("Card Count.csv")
TimeCardFolder := EnvGet("OneDriveCommercial") ;get work folder location
TimeCardFolder := TimeCardFolder . "\Archives\"
TimingCards := DirSelect("*" TimeCardFolder, 0, "Select Timing Card Folder") ;should be at correct spot
if TimingCards = ""
{
    MsgBox("CANCELED!", "Okie Dokie", "T5")
    ExitApp()
}
count := 0
listOfNames := "Name, Count`n"
loop files TimingCards "\*", 'D'
{
    if A_Index > 1
        listOfNames .= "," count "`n"
    count := 0
    ToolTip(A_Index)
    subPath := A_LoopFileFullPath . "\"
    RegExMatch(A_LoopFileFullPath, "Archives\\(.*)", &match)
    listOfNames .= match.1
    loop files, subPath "\*.pdf", "FR" ; go thru all the pdf filenames
    {
        if !(RegExMatch(A_LoopFileDir, "Time Card")) ; only look at Time Card files
            continue
        else
            count++
    }
}
listOfNames .= " - " count
A_Clipboard := listOfNames
fileAppend(listOfNames, "Card Count.csv")
MsgBox("Done! Count copied to clipboard and saved to Card Count.csv", "All Done!", "T5")