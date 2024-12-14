; Version 1.46

; Property of Junior Reitzema
; made to automate cumbersome tasks in Cardiologs
; and some remaps so left hand can stay on keyboard and right hand on mouse
; works in Chrome regardless of the amount of tabs open
; works probably in Edge when Cardiologs is the only tab (not tested)
; at the end you can find shortcuts for coloring in Knipprogramma
; I use that as a way to track each finding per complaint
; works only on windows (AHK not for MacOS)


; abbreviations

:*:bnd::benauwdheid             ; typing bnd prints benauwdheid
:*:hkl::hartkloppingen          ; typing hkl prints hartklopping
:*:ovs::overslagen              ; typing ovs prints overslag
:*:snt::sinustachycardie        ; typing snt prints sinustachycardie
:*:snr::sinusritme              ; typing snr prints sinustachycardie
:*:snb::sinusbradycardie        ; typing snb prints sinusbradycardie
:*:sna::sinusaritmie            ; typing sna prints sinusaritmie
:*:lch::lichthoofdigheid        ; typing lch prints lichthoofdigheid
:*:bkl::borstklacht             ; typing bkl prints borstklacht
:*:isan::is aangetroffen        ; typing isan prints is aangetroffen;
:*:dzl::duizeligheid            ; typing dzl prints duizeligheid
:*:krt::kortademigheid          ; typing krt prints kortademigheid
:*:pob::pijn op de borst        ; typing pob prints pijn op de borst
:*:afib::boezemfibrilleren      ; typing afib prints boezemfibrilleren
:*:hh::huishoudelijke werkzaamheden ; typing hh prints huishoudelijke werkzaamheden
:*:dok::druk op de borst        ; tyiping dok prints druk op de borst
:*:uvb::uniforme ventriculaire bigeminie ; typing uvb prints uniforme ventriculaire bigeminie
:*:uvt::uniforme ventriculaire trigeminie ; typing uvt prints uniforme ventriculaire trigeminie
:*:mvb::multiforme ventriculaire bigeminie ; typing mvb prints multiforme ventriculaire bigeminie
:*:mvt::multiforme ventriculaire trigeminie ; typing mvt prints multiforme ventriculaire trigeminie


#Hotif WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") 

; F1 prints the whole list of combinations

title := "Cheat Sheet - Cardiologs Script"
message := "
(
This is the list of automations using this AHK Cardiologs script.
This script contains abbreviations, remaps, and automates processes within Cardiologs.

Regarding the abbreviations:
- Borstklachten:
  - dok: prints druk op de borst
  - pob: prints pijn op de borst
  - bnd: prints benauwdheid
  - bkl: prints borstklacht
- Andere klachten:
  - lch: prints lichthoofdigheid
  - krt: prints kortademigheid
  - hkl: prints hartkloppingen
  - ovs: prints overslag
  - dzl: prints duizeligheid
- Ritmes:
  - afib: prints boezemfibrilleren
  - snt: prints sinustachycardie
  - snb: prints sinusbradycardie
  - sna: prints sinusaritmie
  - snr: prints sinusritme
- Bi- and trigeminie:
  - uvb: prints uniforme ventriculaire bigeminie
  - uvt: prints uniforme ventriculaire trigeminie
  - mvb: prints multiforme ventriculaire bigeminie
  - mvt: prints multiforme ventriculaire trigeminie
- Overig:
  - isan: prints is aangetroffen
  - hh: prints huishoudelijke werkzaamheden

Automations:
- F2: Comments PR time or P:QRS ratio based on type of the AV-block.
- F4: Prints PT noteert.
- Alt + A: Automatically comments 'Aberrantie' and adds it to the report.
- Alt + N: Prints the right (ns)VT.
- Ctrl + Q: Adds langste (& snelste) PAT/SVT to the report when it has a beginning and end.
- Win + 0 (also Numpad0): Calculates the QTc (using Bazett's formula) and prints all durations.
- Win + 1 (also Numpad1): Prints text when handmatige strook is added for atrial rhythm.
- Win + 2 (also Numpad2): Prints text when handmatige strook is added for ventricular rhythm.
- Win + 4 (also Numpad4): Prints the text for the report when only one complaint is present.
- Win + 5 (also Numpad5): Prints the text for the report when no complaint is present.
- Win + 6 (also Numpad6): Prints the text for the report when no diary is present.

Remaps:
- 'm': Allows merging of multiple strips, remapped to the left '\' (left to 'z' on my keyboard).
- 'Delete': Remapped to '`' key (top left key on many keyboards).
- 'Home': Remapped to 'Win + A' for moving to the beginning of all strips within a section.
- 'End': Remapped to 'Win + D' for moving to the end of all strips within a section.
- 'Alt + Down': Remapped to 'Win + S' for navigating to the next section.
- 'Alt + Up': Remapped to 'Win + W' for navigating to the previous section.

)"

F1::ShowCustomBox("message", title, message, "", 800, 600)

; remaps

; left \ to send "m"
sc056:: Send "m"

; Remap cursor control keys 
`:: Send "{Del}"        ; remaps delete button to `
#a:: Send "{Home}"      ; remaps home button to windows + a       
#d:: Send "{End}"       ; remaps end button to windows + d

; Remap ALT + up & down to (Tab + W) & (Tab + D)
#s:: Send "!{Down}"     ; remaps alt + down to windows + s
#w:: Send "!{Up}"       ; remaps alt + up to windows + d


; automation

F2::        ; F2 comments the PR times for AV-blocks
{
    type := ShowCustomBox("input", "Block Type", "1 for Wenckebach, 2 for Mobitz II, 3 for high-degree block", "", 400, 200)

    if (type == "1") {
        PR_before := ShowCustomBox("input", "PR Time Before", "Enter PR time before non-conducted beat:", "", 400, 200)
        PR_after := ShowCustomBox("input", "PR Time After", "Enter PR time after non-conducted beat:", "", 400, 200)

        if (PR_after > PR_before) {
            ShowCustomBox("message", "Error", "PR after is longer than the PR before", "", 400, 200)
        } else {
            WinActivate("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")

            Send "c"
            Sleep 400
            Send "Wenckebach. PQ voor " PR_before "ms. PQ na " PR_after "ms.{Enter}"
        }
    } else if (type == "2") {
        PR := ShowCustomBox("input", "PR Time", "Enter PR time:", "", 400, 200)

        WinActivate("Philips Cardiologs - Google Chrome") or WinActive("Cardiologs - Persoonlijk - Microsoft Edge")

        Send "c"
        Sleep 400
        Send "Mobitz II. PQ voor & na " PR "ms.{Enter}"
    } else if (type == "3") {
        ratio := ShowCustomBox("input", "P:QRS Ratio", "Enter the P:QRS ratio:", "", 400, 200)

        WinActivate("Philips Cardiologs - Google Chrome") or WinActive("Cardiologs - Persoonlijk - Microsoft Edge")

        Send "c"
        Sleep 400
        Send ratio " AV-blok{Enter}"
    } else {
        ShowCustomBox("message", "Error", "Try Again!", "", 400, 200)
    }
}

F4:: Send "PT noteert '{Space}"

!a::{                   ; alt + a is input aberrantie
    Send "c"
    Sleep 200
    Send "Aberrantie"
    Sleep 200
    Send "{Enter}"
    Sleep 200
} 

^q::        ; (CTRL + Q) prints langste (en snelste) PAT/SVT when it consists of a beginning and end
{
    mode := ShowCustomBox("input", "Select Mode", "1 for langste PAT, 2 for langste & snelste PAT, 3 for langste SVT, etc.", "", 400, 200)

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    if (mode == "1") {
        Send "c"
        Sleep 500
        Send "Begin langste PAT"
        Sleep 300
        Send "{Enter}"
        Sleep 800
        Send "{Right}"
        Sleep 500
        Send "c"
        Sleep 500
        Send "Eind langste PAT"
        Sleep 500
        Send "{Enter}"
    } else if (mode == "2") {
        Send "c"
        Sleep 500
        Send "Begin langste & snelste PAT"
        Sleep 300
        Send "{Enter}"
        Sleep 800
        Send "{Right}"
        Sleep 500
        Send "c"
        Sleep 500
        Send "Eind langste & snelste PAT"
        Sleep 500
        Send "{Enter}"
    } else if (mode == "3") {
        Send "c"
        Sleep 800
        Send "Begin langste SVT"
        Sleep 300
        Send "{Enter}"
        Sleep 800
        Send "{Right}"
        Sleep 500
        Send "c"
        Sleep 500
        Send "Eind langste SVT"
        Sleep 500
        Send "{Enter}"
    } else if (mode == "4") {
        Send "c"
        Sleep 500
        Send "Begin langste & snelste SVT"
        Sleep 300
        Send "{Enter}"
        Sleep 800
        Send "{Right}"
        Sleep 500
        Send "c"
        Sleep 500
        Send "Eind langste & snelste SVT"
        Sleep 500
        Send "{Enter}"
    } else {
        ShowCustomBox("message", "Error", "Try Again!", "", 400, 200)
    }
} 


!n::            ; (alt + n) prints the right (ns)VT
{
    what := ShowCustomBox("input", "Langste of Snelste", "1 for langste, 2 for snelste, 3 for langste & snelste:", "", 400, 200)
    form := ShowCustomBox("input", "Morphologie", "1 for monomorf, 2 for polymorf:", "", 400, 200)

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    if (what == "1" and form == "1") {
        Send "c" 
        Sleep 500
        Send "Langste monomorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 
    
    else if (what == "2" and form == "1") {
        Send "c" 
        Sleep 500
        Send "Snelste monomorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 

    else if (what == "3" and form == "1") {
        Send "c" 
        Sleep 500
        Send "Langste & snelste monomorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 
    else if (what == "1" and form == "2") {
        Send "c" 
        Sleep 500
        Send "Langste polymorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 
    
    else if (what == "2" and form == "2") {
        Send "c" 
        Sleep 500
        Send "Snelste polymorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 

    else if (what == "3" and form == "2") {
        Send "c" 
        Sleep 500
        Send "Langste & snelste polymorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 
}

#0::        ; (Windows Key + 0) calculates QTc and prints it
{
    PR := ShowCustomBox("input", "PR Time", "Please enter the PR time (ms):", "", 400, 200)
    QRS := ShowCustomBox("input", "QRS Time", "Please enter the QRS time (ms):", "", 400, 200)
    QT := ShowCustomBox("input", "QT Time", "Please enter the QT time (ms):", "", 400, 200)
    HR := ShowCustomBox("input", "Heart Rate", "Please enter the heart rate (bpm) or RR interval (ms):", "", 400, 200)

    if (PR == "" || QRS == "" || QT == "" || HR == "") {
        ShowCustomBox("message", "Error", "All inputs are required!", 400, 200)
        return
    }

    if (HR < 200) {
        RR := 60 / HR
        QTc := Round(QT / Sqrt(RR))
    } else {
        QTc := Round(QT / Sqrt((HR / 1000)))
    }

    Send "c"
    Sleep 300
    Send "PQ: " PR " ms, QRS: " QRS " ms, QTc: " QTc "ms{Enter}"
    Sleep 600
}

#Numpad0::      ; (Windows key + Numpad 0) automatically calculates QTc based on HR or RR-interval and prints it
{
    PR := ShowCustomBox("input", "PR Time", "Please enter the PR time (ms):", "", 400, 200)
    QRS := ShowCustomBox("input", "QRS Time", "Please enter the QRS time (ms):", "", 400, 200)
    QT := ShowCustomBox("input", "QT Time", "Please enter the QT time (ms):", "", 400, 200)
    HR := ShowCustomBox("input", "Heart Rate", "Please enter the heart rate (bpm) or RR interval (ms):", "", 400, 200)

    if (PR == "" || QRS == "" || QT == "" || HR == "") {
        ShowCustomBox("message", "Error", "All inputs are required!", 400, 200)
        return
    }

    if (HR < 200) {
        RR := 60 / HR
        QTc := Round(QT / Sqrt(RR))
    } else {
        QTc := Round(QT / Sqrt((HR / 1000)))
    }
    
    Send "c"
    Sleep 300
    Send "PQ: " PR " ms, QRS: " QRS " ms, QTc: " QTc "ms{Enter}"
    Sleep 600
}

#1::        ; (Windows Key + 1) adds atrial rhythm
{
    freq := ShowCustomBox("input", "Frequency", "Please enter the frequency (bpm):", "", 400, 200)
    duration := ShowCustomBox("input", "Duration", "Please enter the duration (seconds):", "", 400, 200)

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    Send "Atriaal ritme, freq " freq " spm, duur " duration " sec."
    Send "{Enter}"
}

#Numpad1::      ; (Windows key + numpad 1) print text when handmatige strook is added for atrial rhythm
{
    freq := ShowCustomBox("input", "Frequency Input", "Please enter the frequency (in bpm):", "", 400, 200)
    duration := ShowCustomBox("input", "Duration Input", "Please enter the duration (in seconds):", "", 400, 200)

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    Send "Atriaal ritme, freq " freq " spm, duur " duration " sec."
    Send "{Enter}"
}

#2::            ; (Windows key + 2) print text when handmatige strook is added for ventricular rhythm
{
    morphology := ShowCustomBox("input", "Morphology Input", "Please enter morphology (1 = uniform, 2 = multiform):", "", 400, 200)
    freq := ShowCustomBox("input", "Frequency Input", "Please enter the frequency (in bpm):", "", 400, 200)
    duration := ShowCustomBox("input", "Duration Input", "Please enter the duration (in seconds):", "", 400, 200)

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")

    if (morphology == "1") {
        Send "Uniform ventriculair ritme, freq " freq " spm, duur " duration " sec."
        Send "{Enter}"
    } else {
        Send "Multiform ventriculair ritme, freq " freq " spm, duur " duration " sec."
        Send "{Enter}"
    }
}

#Numpad2::      ; (Windows key + numpad 2) print text when handmatige strook is added for ventricular rhythm
{
    morphology := ShowCustomBox("input", "Morphology Input", "Please enter morphology (1 = uniform, 2 = multiform):", "", 400, 200)
    freq := ShowCustomBox("input", "Frequency Input", "Please enter the frequency (in bpm):", "", 400, 200)
    duration := ShowCustomBox("input", "Duration Input", "Please enter the duration (in seconds):", "", 400, 200)

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")

    if (morphology == "1") {
        Send "Uniform ventriculair ritme, freq " freq " spm, duur " duration " sec."
        Send "{Enter}"
    } else {
        Send "Multiform ventriculair ritme, freq " freq " spm, duur " duration " sec."
        Send "{Enter}"
    }
}

#4::            ; (Windows key + 4) Prints tekst when only one complaint is present
{
    klacht := ShowCustomBox("input", "Complaint Input", "Please enter the complaint:", "", 400, 200)
    bevinding := ShowCustomBox("input", "Finding Input", "Please enter the finding:", "", 400, 200)

    Send "^a"
    Send "Het dagboekje is wel aanwezig.`nEr is 1 keer een klacht aangegeven. De klacht en bevinding zijn:`n- " klacht " waarbij " bevinding " is aangetroffen. Zie bijgevoegde strook."
}

#Numpad4::      ; (Windows key + numpad 4) Prints tekst when only one complaint is present
{
    klacht := ShowCustomBox("input", "Complaint Input", "Please enter the complaint:", "", 400, 200)
    bevinding := ShowCustomBox("input", "Finding Input", "Please enter the finding:", "", 400, 200)

    Send "^a"
    Send "Het dagboekje is wel aanwezig.`nEr is 1 keer een klacht aangegeven. De klacht en bevinding zijn:`n- " klacht " waarbij " bevinding " is aangetroffen. Zie bijgevoegde strook."
}

#5::            ; (Windows key + 5) Prints tekst when no complaint is present
{
    Send "^a"
    Send "Het dagboekje is wel aanwezig.`nEr is 0 keer een klacht aangegeven."
}

#Numpad5::      ; (Windows key + numpad 5) Prints tekst when no complaint is present
{
    Send "^a"
    Send "Het dagboekje is wel aanwezig.`nEr is 0 keer een klacht aangegeven."
}

#6::            ; (Windows key + 6) Prints tekst when no diary is present
{
    Send "^a"
    Send "Het dagboekje is niet aanwezig."
}

#Numpad6::      ; (Windows key + numpad 6) Prints tekst when no diary is present
{
    Send "^a"
    Send "Het dagboekje is niet aanwezig."
}



#Hotif WinActive("Knipprogramma")

; choose colors for in Knipprogramma / snipping tool

q::         ; press q for sinusritme (oranje)
{
    Send "{Alt}"
    Sleep 250
    Send "b"
    Sleep 250
    Send "{Down}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Enter}"
}

w::         ; press w for sinustachycardie (geel)
{    
    Send "{Alt}"
    Sleep 250
    Send "b"
    Sleep 250
    Send "{Down}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Enter}"
}

e::         ; press e for PVC (rood)
{    
    Send "{Alt}"
    Sleep 250
    Send "b"
    Sleep 250
    Send "{Down}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Enter}"
}

a::         ; press a for PAC (lichtblauw)
{
    Send "{Alt}"
    Sleep 250
    Send "b"
    Sleep 250
    Send "{Down}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Enter}"
}

s::         ; press s for PAT (groen)
{    
    Send "{Alt}"
    Sleep 250
    Send "b"
    Sleep 250
    Send "{Down}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

d::         ; press d for sinusbradycardie (bruin)
{    
    Send "{Alt}"
    Sleep 250
    Send "b"
    Sleep 250
    Send "{Down}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Right}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "{Enter}"
}

Persistent

; Function to get the coordinates of the Philips Cardiologs window
GetActiveCardiologsMonitor() {
    if WinActive("Philips Cardiologs - Google Chrome") || WinActive("Philips Cardiologs - Werklijst - Google Chrome") || WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") {
        ; Get the position and size of the active window
        WinGetPos(&x, &y, &w, &h, "A")
        return { X: x, Y: y, W: w, H: h }
    }
    return False
}

ShowCustomBox(type, title, content, default := "", width := 400, height := 200) {
    monitor := GetActiveCardiologsMonitor()
    if monitor {
        ; Calculate position for centering on the Cardiologs window
        x := monitor.X + monitor.W / 4
        y := monitor.Y + monitor.H / 4

        ; Create a new GUI
        myGui := Gui("+AlwaysOnTop -Caption +Border")
        myGui.Color := "White"  ; Set white background
        myGui.SetFont("s12", "Trebuchet MS")  ; Set font to Trebuchet MS

        ; Add a header
        myGui.Add("Text", "x10 y10 w" (width - 20) " Center", title)

        ; Add content or input box based on type
        if (type = "message") {
            ; Add a scrollable message display
            myGui.Add("Edit", "x10 y+10 w" (width - 20) " h" (height - 100) " ReadOnly +Wrap +VScroll BackgroundWhite", content)
        } else if (type = "input") {
            ; Add prompt and input field
            myGui.Add("Text", "x10 y+10 w" (width - 20), content)
            edit := myGui.Add("Edit", "x10 y+10 w" (width - 20), default)
        }

        ; Add OK and Cancel buttons
        buttonWidth := 80  ; Width of each button
        buttonSpacing := 20  ; Space between buttons
        totalButtonWidth := (2 * buttonWidth) + buttonSpacing
        btnOkX := (width - totalButtonWidth) / 2  ; Center the buttons
        btnCancelX := btnOkX + buttonWidth + buttonSpacing

        btnOk := myGui.Add("Button", "w" buttonWidth " Default", "OK")
        btnCancel := myGui.Add("Button", "w" buttonWidth, "Cancel")

        ; Position buttons at the bottom of the GUI
        btnOk.Move(btnOkX + 10, height - 50)  ; Adjust for padding
        btnCancel.Move(btnCancelX + 10, height - 50)

        ; Initialize flag and return value
        guiValue := ""
        guiActive := True

        ; Event handlers
        btnOk.OnEvent("Click", (*) => (
            guiValue := (type = "input") ? edit.Value : "", guiActive := False, myGui.Destroy()
        ))
        btnCancel.OnEvent("Click", (*) => (guiValue := "", guiActive := False, myGui.Destroy()))

        ; Show the GUI
        myGui.Show("x" x " y" y " w" width " h" height)

        ; Wait for user interaction
        while guiActive {
            Sleep 50
        }

        ; Return value based on type
        return guiValue
    } else {
        MsgBox("Philips Cardiologs is not active!")
        return False
    }
}




/*

ShowCustomMessageBox(title, message, width := 800, height := 600) {
    monitor := GetActiveCardiologsMonitor()
    if monitor {
        ; Calculate position for centering on the Cardiologs window
        x := monitor.X + monitor.W / 4
        y := monitor.Y + monitor.H / 4

        ; Create a new GUI for the message box
        myGui := Gui("+AlwaysOnTop -Caption +Border")  ; Add black border
        myGui.Color := "White"  ; Set the GUI background color to white

        ; Set font globally to Trebuchet MS
        myGui.SetFont("s12", "Trebuchet MS")

        ; Add a header
        myGui.Add("Text", "w" width " Center", title)

        ; Add a scrollable Edit control for the main message
        myGui.Add("Edit", "w" width " h" (height - 100) " ReadOnly +Wrap +VScroll BackgroundWhite", message)

        ; Add the OK button
        btnOk := myGui.Add("Button", "w80 Default", "OK")
        btnOk.Move((width - 80) / 2, height - 40)

        ; Event handler for OK button to close the GUI
        btnOk.OnEvent("Click", (*) => myGui.Destroy())

        ; Set GUI title and display
        myGui.Title := title
        myGui.Show("x" x " y" y " w" width " h" height)
    } else {
        MsgBox("Philips Cardiologs is not active!")
    }
}

ShowCustomInputBox(title, prompt, default := "", width := 400, height := 200) {
    monitor := GetActiveCardiologsMonitor()
    if monitor {
        ; Calculate position for centering on the Cardiologs window
        x := monitor.X + monitor.W / 4
        y := monitor.Y + monitor.H / 4

        ; Create a new GUI for the input box
        myGui := Gui("+AlwaysOnTop -Caption +Border")
        myGui.Color := "White"  ; Set white background
        myGui.SetFont("s12", "Trebuchet MS")  ; Set font to Trebuchet MS

        ; Add a prompt
        myGui.Add("Text", "w" width, prompt)

        ; Add an Edit control for user input
        edit := myGui.Add("Edit", "w" (width - 20), default)

        ; Add OK and Cancel buttons
        buttonWidth := 80  ; Explicitly define button width
        buttonSpacing := 10
        totalButtonWidth := 2 * buttonWidth + buttonSpacing
        btnOkX := (width - totalButtonWidth) / 2
        btnCancelX := btnOkX + buttonWidth + buttonSpacing

        btnOk := myGui.Add("Button", "w" buttonWidth " Default", "OK")
        btnCancel := myGui.Add("Button", "w" buttonWidth, "Cancel")

        ; Position buttons at the bottom
        btnOk.Move(btnOkX, height - 40)
        btnCancel.Move(btnCancelX, height - 40)

        ; Initialize flag and value
        guiValue := ""
        guiActive := True

        ; Event handlers
        btnOk.OnEvent("Click", (*) => (guiValue := edit.Value, guiActive := False, myGui.Destroy()))
        btnCancel.OnEvent("Click", (*) => (guiValue := "", guiActive := False, myGui.Destroy()))

        ; Show the GUI
        myGui.Show("x" x " y" y " w" width " h" height)

        ; Wait for user interaction
        while guiActive {
            Sleep 50
        }

        ; Return the user's input or empty string on cancel
        return guiValue
    } else {
        MsgBox("Philips Cardiologs is not active!")
        return False
    }
}

*/









/* something I am working on, not for production
#k::{
    klacht := InputBox("What is the complaint ","Cardiologs - Google Chrome","w300 h150")
    nklacht := InputBox("What is the total input of the complaint","Cardiologs - Google Chrome","w300 h150").Value
    
    bevindingen := Array()
    nbevindingen := Array()

    i := InputBox("How many findings?", "Cardiologs - Google Chrome", "w300 h150").value
       
    Loop i{
        j := InputBox("What is the finding?", "Cardiologs - Google Chrome", "w300 h150")
        k := InputBox("How many times?", "Cardiologs - Google Chrome", "w300 h150").value

        bevindingen.Push j
        nbevindingen.Push k
    }

    Loop bevindingen.Length{
        MsgBox bevindingen[A_Index]
    } 

    
    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    Send "{Raw}-"
    Send "{Space}"
    Send klacht.Value 
    Send "(" nklacht "x) waarbij {Space}"

    for l in nbevindingen.Length {
        Send bevindingen[l]
        Send "(" 
        Send nbevindingen[l]
        Send "x), "
    } 
   
    Send "is aangetroffen."
}
*/ 
