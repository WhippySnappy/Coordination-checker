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
/*
COPILOT - JSON EXAMPLE
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Coordination Checker</title>
</head>
<body oncontextmenu="ToggleFullScreen(); return false">
    <script>
        // JSON object to store variables
        const variables = {
            "Valencia_20thSt": {
                "offset": [49],
                "LPI": 4,
                "Yel": 4,
                "Red": 1.5,
                "Adj": [0],
                "Split": [30]
            },
            "Valencia_21stSt": {
                "offset": [19],
                "LPI": 4,
                "Yel": 4,
                "Red": 1.5,
                "Adj": [0],
                "Split": [30]
            },
            // Add more intersections here...
        };

        function Timing() {
            // Example of accessing variables from JSON object
            const valencia20th = variables.Valencia_20thSt;
            ChangeDisp('Valencia_20thSt_change', 19, planNum, cl);
            TimingCalc('Valencia_20thSt_NS_REF', valencia20th);
            TimingCalc('Valencia_20thSt_EW', valencia20th);

            const valencia21st = variables.Valencia_21stSt;
            ChangeDisp('Valencia_21stSt_change', 17, planNum, cl);
            TimingCalc('Valencia_21stSt_NS_REF', valencia21st);
            TimingCalc('Valencia_21stSt_EW', valencia21st);

            // Add more intersections here...
        }

        function ChangeDisp(inter, change, plan, cycle) {
            // Your existing ChangeDisp function code...
        }

        function TimingCalc(PhaseDisp, vars) {
            // Use vars to access the variables for the specific intersection
            console.log(PhaseDisp, vars);
            // Your existing TimingCalc function code...
        }

        // Call the Timing function to start the process
        Timing();
    </script>
</body>
</html>
*/