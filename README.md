# Autocompila
Autocopilatore in Python con istruzioni in windows 
Questo software implementa la lettura di file .xlsx e .xls e la conversione, post lettura, in JSON o PDF e il nuovo pulsante che AUTOCOMPILA UN FORM.

Installazione delle librerie su Windows

1. Aprire il terminale (Prompt dei comandi o PowerShell) ed eseguire i seguenti comandi per creare un ambiente virtuale nella directory del progetto e installare le librerie necessarie:
  
  python -m venv myenv
  myenv\Scripts\activate
  pip install PyQt5 pandas openpyxl xlrd
  pip install reportlab

Nota:L'ambiente virtuale viene utilizzato per l'installazione delle librerie necessarie al progetto senza dover procedere a installazioni globali, evitando così problemi di compatibilità tra versioni diverse.

2 Attivare l'ambiente virtuale (se non già attivo):

  myenv\Scripts\activate

Creazione dell'eseguibile

   1 Installare la libreria pyinstaller eseguendo il seguente comando:
        pip install pyinstaller
    
  2 Generare l'eseguibile con il seguente comando:
          pyinstaller --windowed --onedir ui.py
    
  Nota:  L'eseguibile verrà generato in una cartella chiamata "dist", all'interno della quale saranno inclusi anche i file di supporto necessari al funzionamento del programma.
