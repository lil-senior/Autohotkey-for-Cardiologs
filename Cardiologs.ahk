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

help := "
( 
This is the list of automations using this AHK Cardiologs script.
This script contains abbreviations, remaps, and automates processes within Cardiologs.

Regarding the abbreviations:
Borstklachten: dok prints druk op de borst, pob prints pijn op de borst, bnd prints benauwdheid, and bkl prints borstklacht.
Andere klachten: lch prints lichthoofdigheid, krt prints kortademigheid, hkl prints hartklopping, ovs prints overslag, and typing dzl prints duizeligheid.
Ritmes: afib prints boezemfibrilleren, snt prints sinustachycardie, snb prints sinusbradycardie, sna prints sinusaritmie, and snr prints sinusritme.
Bi- and trigeminie: uvb prints uniforme ventriculaire bigeminie, uvt prints uniforme ventriculaire trigeminie, mvb prints multiforme ventriculaire bigeminie, and mvt prints multiforme ventriculaire trigeminie.
Overig: isan prints is aangetroffen and hh prints huishoudelijke werkzaamheden

Regarding the automation:
F2 comments PR time or P:QRS ratio based on type of the AV-block.
F4 prints PT noteert.'
Alt + a automatically comments 'Aberrantie' and adds it to the report.
Alt + n prints the right (ns)VT.
Ctrl + q adds langste (& snelste) PAT/SVT to the report when it has a beginning and end.
Win + 0 (also Win + Numpad0) calculates the QTc (using Bazett's formula) and prints all the durations.
Win + 1 (also Win + Numpad1) prints text when handmatige strook is added for atrial rhythm.
Win + 2 (also Win + Numpad2) prints text when handmatige strook is added for ventricular rhythm.
Win + 4 (also Win + Numpad4) prints the text for the report when only one complaint is present.
Win + 5 (also Win + Numpad5) prints the text for the report when no complaint is present.
Win + 6 (also Win + Numpad6) prints the text for the report when no diary is present.

Regarding the remaps:
'm' (to allow merging of multiple strips) is remapped to the left '\' (left to 'z' on my keyboard).
'Delete' is remapped to '`' key (top left key on many keyboards).
'Home' to quickly move to the beginning of all strips within a section is remapped to 'Win + a'.
'End' to quickly move to the end of all strips within a section is remapped to 'Win + d'.
'Alt + Down' to go the next section is remapped to 'Win + s'.
'Alt + Up' to go to the previous section is remapped to 'Win + w'. 
)"

F1::MsgBox(help)

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
    type := InputBox("1 for Wenckebach, 2 Mobitz II, 3 for high-degree block").value

    if (type==1) {
        PR_before := InputBox("Enter PR time before non-conducted beat").value
        PR_after := InputBox("Enter PR time after non-conducted beat").value
        
        if (PR_after > PR_before) {
            MsgBox("Error. PR after is longer than the PR before")
        }
        else {
            WinActivate("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")

            Send "c"
            Sleep 400
            Send "Wenckebach. PQ voor " PR_before "ms. PQ na " PR_after "ms.{Enter}"
        }
    }

    else if (type==2) {
        PR := InputBox("Enter PR time").value
        
        WinActivate("Philips Cardiologs - Google Chrome") or WinActive("Cardiologs - Persoonlijk - Microsoft Edge")

        Send "c"
        Sleep 400
        Send "Mobitz II. PQ voor & na " PR "ms.{Enter}"
    }

    else if (type==3) {

        ratio := InputBox("Enter the P:QRS ratio").value
        
        WinActivate("Philips Cardiologs - Google Chrome") or WinActive("Cardiologs - Persoonlijk - Microsoft Edge")

        Send "c"
        Sleep 400
        Send ratio " AV-blok{Enter} "
    }

    else {
        MsgBox "Try Again!"
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


^q::            ; (CRTL + q) prints langste (en snelste) PAT/SVT when it consists of a beginning and end. 
{
    mode := InputBox("1 for langste PAT, 2 for langste en snelste PAT, 3 for langste SVT, 4 for langste en snelste SVT", "begin/eind", "w300 h150").value
    
    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    if (mode == 1) {
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
    } 
    
    else if (mode == 2) {
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
    } 

    else if (mode == 3) {
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
    }

    else if (mode == 4) {
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
    } 
    else {
        MsgBox "Try Again!"
    }
} 


!n::            ; (alt + n) prints the right (ns)VT 
{
    what:= InputBox("1 for langste, 2 for snelste, 3 for langste & snelste", "Langste of snelste ", "w300 h150").value
    form:= InputBox("1 for monomorf, 2 for polymorf", "Morphologie", "w300 h150").Value

        
    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    if (what == 1 and form == 1 ) {
        Send "c" 
        Sleep 500
        Send "Langste monomorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 
    
    else if (what == 2 and form == 1 ) {
        Send "c" 
        Sleep 500
        Send "Snelste monomorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 

    else if (what == 3 and form == 1 ) {
        Send "c" 
        Sleep 500
        Send "Langste & snelste monomorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 
    if (what == 1 and form == 2 ) {
        Send "c" 
        Sleep 500
        Send "Langste polymorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 
    
    else if (what == 2 and form == 2 ) {
        Send "c" 
        Sleep 500
        Send "Snelste polymorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 

    else if (what == 3 and form == 2 ) {
        Send "c" 
        Sleep 500
        Send "Langste & snelste polymorfe nsVT"
        Sleep 500
        Send "{Enter}"    
    } 
} 

#0::            ; (Windows key + 0) automatically calculates QTc based on HR or RR-interval and prints it
{
    PR := InputBox("Please enter the PR time (in milliseconds)","QT Time Input","w300 h150").value
    QRS := InputBox("Please enter the QRS time (in milliseconds)","QT Time Input","w300 h150").value
    QT := InputBox("Please enter the QT time (in milliseconds)","QT Time Input","w300 h150").value
    HR := InputBox("Please enter the heart rate (in beats per minute) or RR-interval (in ms)","Heart Rate Input", "w300 h150").value

    if (HR < 200) {
        RR := 60 / HR
        QTc := Round(QT / Sqrt(RR))
        }

    else {  
        QTc := Round(QT / Sqrt((HR/1000)))    
        }

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    Sleep 300
    Send "c" 
    Sleep 300
    Send "PQ: " PR " ms, QRS: " QRS " ms, QTc: " QTc "ms{Enter}"
    Sleep 600
}

#Numpad0::      ; (Windows key + Numpad 0) automatically calculates QTc based on HR or RR-interval and prints it
{
    PR := InputBox("Please enter the PR time (in milliseconds)","QT Time Input","w300 h150").value
    QRS := InputBox("Please enter the QRS time (in milliseconds)","QT Time Input","w300 h150").value
    QT := InputBox("Please enter the QT time (in milliseconds)","QT Time Input","w300 h150").value
    HR := InputBox("Please enter the heart rate (in beats per minute) or RR-interval (in ms)","Heart Rate Input", "w300 h150").value

    if (HR < 200) {
        RR := 60 / HR
        QTc := Round(QT / Sqrt(RR))
        }

    else {  
        QTc := Round(QT / Sqrt((HR/1000)))    
        }

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    Sleep 300
    Send "c" 
    Sleep 300
    Send "PQ: " PR " ms, QRS: " QRS " ms, QTc: " QTc "ms{Enter}"
    Sleep 600
}


#1::            ; (Windows key + 1) print text when handmatige strook is added for atrial rhythm
{

    freq := InputBox("Please enter the frequency (in bpm)","frequency atrial rhythm input","w300 h150").value
    duration := InputBox("Please enter the duration (in seconds)","duration atrial rhythm input", "w300 h150").value

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    Send "Atriaal ritme, freq " freq " spm, duur " duration " sec."
    Send "{Enter}"
}

#Numpad1::      ; (Windows key + numpad 1) print text when handmatige strook is added for atrial rhythm
{

    freq := InputBox("Please enter the frequency (in bpm)","frequency atrial rhythm input","w300 h150").value
    duration := InputBox("Please enter the duration (in seconds)","duration atrial rhythm input", "w300 h150").value

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")
    Send "Atriaal ritme, freq " freq " spm, duur " duration " sec."
    Send "{Enter}"
}

#2::            ; (Windows key + 2) print text when handmatige strook is added for ventricular rhythm
{

    morphology := InputBox("Please enter morphology; 1 is uniform, 2 is multiform", "morphology input", "w300 h150").value
    freq := InputBox("Please enter the frequency (in bpm)","frequency ventriculair rhythm input","w300 h150").value
    duration := InputBox("Please enter the duration (in seconds)","duration atrial rhythm input", "w300 h150").value

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")

    if morphology == 1 {
        Send  "Uniform ventriculair ritme, freq " freq " spm, duur " duration " sec."
        Send "{Enter}"
    } else {
        Send "Multiform ventriculair ritme, freq " freq " spm, duur " duration " sec."
        Send "{Enter}"
    }
}

#Numpad2::      ; (Windows key + numpad 2) print text when handmatige strook is added for ventricular rhythm
{

    morphology := InputBox("Please enter morphology; 1 is uniform, 2 is multiform", "morphology input", "w300 h150").value
    freq := InputBox("Please enter the frequency (in bpm)","frequency ventriculair rhythm input","w300 h150").value
    duration := InputBox("Please enter the duration (in seconds)","duration atrial rhythm input", "w300 h150").value

    WinActive("Philips Cardiologs - Google Chrome") or WinActivate("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge")

    if morphology == 1 {
        Send  "Uniform ventriculair ritme, freq " freq " spm, duur " duration " sec."
        Send "{Enter}"
    } else {
        Send "Multiform ventriculair ritme, freq " freq " spm, duur " duration " sec."
        Send "{Enter}"
    }
}

#4::            ; (Windows key + 4) Prints tekst when only one complaint is present 
{
    klacht := InputBox("Please enter the complaint","Klacht","w300 h150")
    bevinding := InputBox("Please enter the finding","Bevinding","w300 h150")

    Send "^a"
    Send "Het dagboekje is wel aanwezig.`nEr is 1 keer een klacht aangegeven. De klacht en bevinding zijn:`n- " klacht.Value " waarbij " bevinding.Value " is aangetroffen. Zie bijgevoegde strook."
}

#Numpad4::      ; (Windows key + numpad 4) Prints tekst when only one complaint is present 
{
    klacht := InputBox("Please enter the complaint","Klacht","w300 h150")
    bevinding := InputBox("Please enter the finding","Bevinding","w300 h150")

    Send "^a"
    Send "Het dagboekje is wel aanwezig.`nEr is 1 keer een klacht aangegeven. De klacht en bevinding zijn:`n- " klacht.Value " waarbij " bevinding.Value " is aangetroffen. Zie bijgevoegde strook."
}

#5::            ; (Windows key + 5) Prints tekst when only no complaint is present 
{ 
    Send "^a" 
    Send "Het dagboekje is wel aanwezig.`nEr is 0 keer een klacht aangegeven."
}

#Numpad5::      ; (Windows key + numpad 5) Prints tekst when only no complaint is present 
{ 
    Send "^a" 
    Send "Het dagboekje is wel aanwezig.`nEr is 0 keer een klacht aangegeven."
}

#6::            ; (Windows key + 6) Prints tekst when no dairy is present
{ 
    Send "^a" 
    Send "Het dagboekje is niet aanwezig."
}

#Numpad6::      ; (Windows key + numpad 6) Prints tekst when no dairy is present
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
