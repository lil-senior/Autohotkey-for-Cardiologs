; Version 1.46

; Property of Junior Reitzema
; made to automate cumbersome tasks in Cardiologs
; and some remaps so left hand can stay on keyboard and right hand on mouse
; works in Chrome regardless of the amount of tabs open
; works probably in Edge when Cardiologs is the only tab (not tested)
; at the end you can find shortcuts for coloring in Knipprogramma
; I use that as a way to track each finding per complaint
; works only on windows (AHK not for MacOS)



#Hotif WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")

; F1 prints the whole list of combinations
F1::{
    MsgBox "This is the list of automations using this AHK Cardiologs script.`nThis scripts contains abbreviations, remaps, and automates processes within Cardiologs.`n`nRegarding the abbreviations:`ntyping dok prints druk op de borst `ntyping snb prints sinusbradycardie `ntyping afib prints boezemfibrilleren `ntyping pob prints pijn op de borst `ntyping krt prints kortademigheid `ntyping dzl prints duizeligheid `ntyping isan prints is aangetroffen; `ntyping bnd prints benauwdheid `ntyping hkl prints hartklopping `ntyping ovs prints overslag `ntyping snt prints sinustachycardie `ntyping snr prints sinusritme `ntyping bkl prints borstklacht `ntyping lch prints lichthoofdigheid `n`nRegarding the automation:`nF2 comments PR time or P:QRS ratio based on type of AV-block. `nF4 prints PT noteert ' `nalt + a automatically comments 'Aberrantie' and adds it to the report.`ncrtl + q adds langste (& snelste) PAT/SVT to the report when it has a beginning and end.`nwin + 0 (also win + numpad0) automatically calculates the QTc (using Bazett's formula) and prints it.`nwin + 1 (also win + numpad1) automatically prints text when handmatige strook is added for atrial rhythm.`nwin + 2 (also win + numpad2) automatically prints text when handmatige strook is added for ventricular rhythm.`nwin + 4 (also win + numpad4) automatically prints the text for the report when only one complaint is present.`nwin + 5 (also win + numpad5) automatically prints the text for the report when no complaint is present.`nwin + 6 (also win + numpad6) automatically prints the text for the report when no diary is present.`n`nRegarding the remaps:`n'm' (to allow merging of multiple strips) is remapped to the left '\' (left to 'z' on my keyboard).`n'Delete' is remapped to '`'.`n'Home' to quickly move to the beginning of all strips within a section is remapped to 'win + a'.`n'End' to quickly move to the end of all strips within a section is remapped to 'win + d'.`n'Alt + down' to go the next section is remapped to 'win + s'.`n'Alt + up' to go to the previous section is remapped to 'win + w'."
}


; abbreviations

:*:bnd::benauwdheid             ; typing bnd prints benauwdheid
:*:hkl::hartkloppingen          ; typing hkl prints hartklopping
:*:ovs::overslagen              ; typing ovs prints overslag
:*:snt::sinustachycardie        ; typing snt prints sinustachycardie
:*:snr::sinusritme              ; typing snr prints sinustachycardie
:*:snb::sinusbradycardie        ; typing snb prints sinusbradycardie
:*:lch::lichthoofdigheid        ; typing lch prints lichthoofdigheid
:*:bkl::borstklacht             ; typing bkl prints borstklacht
:*:isan::is aangetroffen;       ; typing isan prints is aangetroffen;
:*:dzl::duizeligheid            ; typing dzl prints duizeligheid
:*:krt::kortademigheid          ; typing krt prints kortademigheid
:*:pob::pijn op de borst        ; typing pob prints pijn op de borst
:*:afib::boezemfibrilleren      ; typing afib prints boezemfibrilleren
:*:hh::huishoudelijke werkzaamheden ; typing hh prints huishoudelijke werkzaamheden
:*:dok::druk op de borst        ; tyiping dok prints druk op de borst



; remaps



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

    
    WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")
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


; left \ to send "m"
sc056:: Send "m"

; Remap cursor control keys 
`:: Send "{Del}"        ; remaps delete button to `
#a:: Send "{PgUp}"      ; remaps home button to windows + a       
#d:: Send "{PgDn}"       ; remaps end button to windows + d

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
            WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")

            Send "c"
            Sleep 400
            Send "Wenckebach. PQ voor " PR_before "ms. PQ na " PR_after "ms.{Enter}"
        }


    }

    else if (type==2) {
        PR := InputBox("Enter PR time").value
        
        WinActive("Philips Cardiologs - Google Chrome") or WinActive("Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")

        Send "c"
        Sleep 400
        Send "Mobitz II. PQ voor & na " PR "ms.{Enter}"
    }

    else if (type==3) {

        ratio := InputBox("Enter the P:QRS ratio").value
        
        WinActive("Philips Cardiologs - Google Chrome") or WinActive("Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")

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
    
    WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")
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

    WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")
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

    WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")
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

    WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")
    Send "Atriaal ritme, freq " freq " spm, duur " duration " sec."
    Send "{Enter}"
}

#Numpad1::      ; (Windows key + numpad 1) print text when handmatige strook is added for atrial rhythm
{

    freq := InputBox("Please enter the frequency (in bpm)","frequency atrial rhythm input","w300 h150").value
    duration := InputBox("Please enter the duration (in seconds)","duration atrial rhythm input", "w300 h150").value

    WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")
    Send "Atriaal ritme, freq " freq " spm, duur " duration " sec."
    Send "{Enter}"
}

#2::            ; (Windows key + 2) print text when handmatige strook is added for ventricular rhythm
{

    morphology := InputBox("Please enter morphology; 1 is uniform, 2 is multiform", "morphology input", "w300 h150").value
    freq := InputBox("Please enter the frequency (in bpm)","frequency ventriculair rhythm input","w300 h150").value
    duration := InputBox("Please enter the duration (in seconds)","duration atrial rhythm input", "w300 h150").value

    WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")

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

    WinActive("Philips Cardiologs - Google Chrome") or WinActive("Philips Cardiologs - Werklijst - Google Chrome") or WinActive("Philips Cardiologs - Persoonlijk - Microsoft Edge") or WinActive("Philips Cardiologs - School - Microsoft​ Edge")

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

; choose colors for in Knipprogramma

; sinusritme = oranje

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






