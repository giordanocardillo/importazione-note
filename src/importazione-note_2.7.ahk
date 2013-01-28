;*** Importazione note per Nota Sogei v2.7 - (c)Giordano Cardillo 2010 ***

;Impostazione del programma come persistente e forzamento a singola istanza
#Persistent
#SingleInstance  force

;Aggiunta del menu contestuale alla tray
Menu, Tray, NoStandard
Menu, Tray, Add, &Info..., About
Menu, Tray, Add
Menu, Tray, Add, &Esci, ExitProg

;Versione del programma
vers=2.7
;Verifica se il programma Nota è in esecuzione
Process, Exist, nota.exe
PID = %ErrorLevel%
;Se sta eseguendo
If PID <> 0
{
    ;Conferma di svuotamento dell'archivio del programma
    MsgBox,49,Attenzione,Occorre svuotare l'archivio prima di continuare. Svuotarlo?
    IfMsgBox,OK
    {
        ;Attendo che il programma Nota sia attivo
        WinWait,ahk_pid %PID%
        WinActivate
        ;Imposto su "Nota"       
        Send,{AltDown}S{AltUp}P
        WinWait,Personalizza
        WinActivate
        Send,N{Enter}
        ;Richiamo il comando per svuotare l'archivio
        Send,{AltDown}ALUT{AltUp}
        ;Nel caso in cui appaia l'errore di note non presenti, viene premuto OK
        WinWait,Messaggio,Non sono presenti note da annullare,1
        If ErrorLevel = 0
        {
            WinActivate
            Send,{Enter}
            ;Salto alla selezione della cartella contenente le note
            Goto,Select
        }
        ;Nel caso in cui appare la finestra di selezione delle note da cancellare
        WinWait,Cancellazione totale note
        WinActivate
        ;Selezione del pulsante "Tutte"
        Send, {AltDown}T{AltUp}
        WinWait,Messaggio,Confermi la richiesta
        WinActivate
        ;Selezione del pulsante "Si" di conferma
        Send, {AltDown}S{AltUp}        
        ;Salto alla selezione della cartella contenente le note
        Goto,Select        
    }
    Else
        ;Se il programma non è in esecuzione, lo script termina.
        ExitApp,1
    Select:
    Gui, Add, DropDownList, vFormalita,Iscr. Ipot.||Canc. Ipot.|Canc. Pign.|Trasc. Pign.
    Gui, Add, Button,Default,&OK
    Gui, Add, Button,,&Annulla
    Gui, Add, Text, vLabel, Selezione tipo formalità:
    GuiControl,Move,Formalita,x140 y23 w120 h10
    GuiControl,Move,&OK,x120 y53 w67 h23
    GuiControl,Move,&Annulla,x193 y53 w67 h23
    GuiControl,Move,Label,x16 y26 w120 h20
    Gui, Show, w275 h85, Selezione tipo formalità
    return   
    ButtonOK:
    Gui,Submit
    Gui,Destroy
    If (Formalita = "Iscr. Ipot.")
        pattern = nota1
    Else If (Formalita = "Canc. Ipot.")
        pattern = nota6
    Else If (Formalita = "Canc. Pign.")
        pattern = nota9
    Else If (Formalita = "Trasc. Pign.")
        pattern = nota8  
    ;Selezione della cartella contenente le note
    FileSelectFolder, Folder,,0,Seleziona la cartella contenente le note da importare.
    ;Se non viene selezionata nessuna cartella lo script termina
    If Folder =
        ExitApp,1
    ;Se nel percorso della cartella sono presenti spazi, lo script termina avvisando l'utente
    Else IfInString,Folder,%A_Space%
    {
        MsgBox,48,Attenzione,Non possono esserci spazi nel nome del percorso.
        ;Riselezione della cartella
        Goto,Select
    }
    Else
    {        
        ;Altrimenti si controlla se nella cartella selezionata sono presenti file del pattern specificato
        IfExist, %Folder%\%pattern%*.dat
        {
           ;Conteggio del numero dei file del pattern specificato presenti
           Loop, %Folder%\%pattern%*.dat
            {
                Tot=%A_Index%            
            }
            ;Inizio importazione file
           Loop, %Folder%\%pattern%*.dat
            {                
                ;Attivazione della finestra del programma nota
                WinWait,ahk_pid %PID%
                WinActivate        
                ;Invio comando di importazione
                Send,{AltDown}FI{AltUp}
                ;Scelta del file
                WinWait,Scelta del file
                WinActivate
                ;Scrittura del percorso del file nota da importare
                Send, %Folder%\%A_LoopFileName%
                Send, {AltDown}A{AltUp}            
                WinWait,Importa
                WinActivate
                ;Selezione del pulsante "Tutte"
                Send, {AltDown}T{AltUp}
                WinWait,Messaggio,Confermi la richiesta
                WinActivate
                ;Selezione del pulsante "Si" di conferma
                Send, {AltDown}S{AltUp}                 
                Prog:=A_Index/Tot*100 
                ;Creazione e aggiornamento della barra di avanzamento
                Progress,%Prog%,%A_LoopFileName%,Importazione in corso (%A_Index%/%Tot%),Progresso importazione,MS Sans Serif
                Note = %A_Index%
            } 
            Sleep, 1000
            Progress, OFF
            ;Messaggio all'utente contenente il numero di note importate
            MsgBox,64,Importazione terminata,Note importate: %Note%. 
            ;Uscita dall'applicazione
            ExitApp,0
        }
        Else
        {
            ;Avviso all'utente della non esistenza di file del pattern specificato
            MsgBox,48,Attenzione,Non ci sono file compatibili Sogei del tipo selezionato in questo percorso.
            ;Riselezione della cartella
            Goto,Select
        }
    }
    ExitApp,1
}
Else
{
    ;Messaggio all'utente di avviso che il programma Nota non è aperto
    MsgBox,16,Attenzione,Il programma Nota Sogei non è aperto.
    ExitApp,1
}
About:
TrayTip, Importazione note per Nota Sogei v%vers%, Scritto da Giordano Cardillo,,1
Exit
ButtonAnnulla:
GuiClose:
ExitApp,1
ExitProg:
ExitApp,0

