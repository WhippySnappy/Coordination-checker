#Requires AutoHotkey v2.0
#SingleInstance Force
Position := 10
MyGui := Gui()
MyGui.BackColor := "black"
MyGui.Opt("+AlwaysOnTop -Caption +Disabled -SysMenu +ToolWindow")
MyProgress := MyGui.Add("Progress", "cRed w200")
MyText := MyGui.Add("Text", "cSilver x33 wp-100") ; wp means "use width of previous".
MyGui.Show()
MyGui.Move(, 10) ;move AFTER show
MoveBar()
ExitApp

MoveBar(*)
{
    loop 100
    {
        x := A_Index
        MyProgress.Value := x 
        MyText.Value := A_Index " %" ;
        Sleep 50
    }
}
