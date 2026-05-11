=============================================
 BDE Assignment Tool — Guida rapida
=============================================

Cosa fa:
  Assegna automaticamente il BDE (Business Data Element) ai campi
  NDG, Codice Fiscale e Partita IVA, classificando i casi con
  campi multipli per tabella in pattern per facilitare la verifica.

Output:
  - output_bde_v10_smart_tag.xlsx    (file principale con BDE e RGP_PERIMETER)
  - REPORT_COMPLETO_DA_VERIFICARE.xlsx (dettaglio casi da verificare)


---------------------------------------------
 INSTALLAZIONE E LANCIO (tutto automatico)
---------------------------------------------

1. Installare Python 3 (se non presente)
   - Windows: https://www.python.org/downloads/
     IMPORTANTE: durante l'installazione spuntare "Add Python to PATH"
   - Mac: gia' presente di default

2. Lanciare lo script:

   Mac/Linux:
     Aprire Terminal, andare nella cartella script e digitare:
     bash setup_e_lancia.sh

   Windows:
     Doppio click su setup_e_lancia.bat

   Il primo lancio installa automaticamente le dipendenze.
   I lanci successivi partono subito.

3. Si apre una pagina nel browser con 3 campi:
   - Clicca "Sfoglia" e seleziona il file Discovery (output_bde_v10_smart_tag...)
   - Clicca "Sfoglia" e seleziona il file Perimetro (bde_CON_RGP_PERIMETER...)
   - Clicca "Sfoglia" e seleziona il file Mapping (DM_Estrazione_Discovery...)

4. Scegli la cartella dove salvare l'output, poi premi "Esegui".
   Il log in basso mostra l'avanzamento.
   Al termine trovi i 2 file Excel nella cartella scelta.

5. Per chiudere: Ctrl+C nel terminale.


---------------------------------------------
 RISOLUZIONE PROBLEMI
---------------------------------------------

Errore "No module named pandas":
   Ripetere il passaggio 4 dell'installazione.

Errore "python non riconosciuto" (Windows):
   Reinstallare Python spuntando "Add Python to PATH",
   oppure usare il percorso completo: C:\Python3x\python.exe auto_bde_gui.py

Errore "Disallineamento colonna...":
   I file Discovery e Perimetro devono avere le stesse righe
   nello stesso ordine. Verificare che siano la versione corretta.

La GUI non si apre:
   Lo script apre una pagina nel browser all'indirizzo
   http://127.0.0.1:8787 — se non si apre automaticamente,
   copiare questo indirizzo nel browser manualmente.


---------------------------------------------
 FILE NELLA CARTELLA
---------------------------------------------

auto_bde_gui.py        Lo script con interfaccia grafica
auto_bde_assignment.py  Lo script da riga di comando (senza GUI)
requirements.txt        Dipendenze Python
README.txt              Questo file
