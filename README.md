<!-- # -*- coding: utf-8 -*- -->

# RedOdsReader
## Leggere i file OpenDocument(.ODS)

Lo scopo di questa Classe/Libreria è quello di **leggere** i file **.ODS**. Questo formato è un sottoinsieme delle specifiche OpenDocument ed è utilizzato da LibreOffice Calc, OpenOffice e da molti altri software.

### Le funzionalità

La versione corrente di **RedOdsReader** implementa queste funzioni:

* **Caricamento** di un file .ODS.
* Lettura dell'**elenco fogli/tabelle** presenti.
* Per ogni foglio, lettura dell'**ultima colonna** nella quale è presente un dato valido e riconosciuto.
* Per ogni foglio, lettura dell'**ultima riga** nella quale è presente un dato valido e riconosciuto.
* **Lettura del valore di una singola cella**, identificata dalle sua coordinate (Colonna/Riga) all'interno di un singolo foglio/tabella.

Il caricamento dei file avviene interamente in RAM all'interno di una struttura ad oggetti dove vengono conservati i dati trattati dalla libreria.

### Usare RedOdsReader

Vediamo ora come usare **RedOdsReader**.

Prima di tutto occorre includere nel progetto corrente il file **RedOdsCl.cs**. Come secondo, e ultimo, passaggio è necessario includere nel progetto i riferimenti a **System.IO.Compression** e **System.IO.Compression.FileSystem**.

Fatto ciò, vediamo un po di codice:

```c#
    //qualche variabile
    int lastRow;
    int lastCol;
    string sheetName;
    List<string> sheetsList;
    string value;
    
    //Prima di tutto ci serve una istanza 
    //  dell'oggetto RedOdsReader
    ROdsLib.RedOdsReader OdsObj = new ROdsLib.RedOdsReader();
    
    //Apriamo un file .ODS
    OdsObj.LoadFile("c:\\test\\citta_italia.ods");

    //impostiamo il nome del foglio/tabella da leggere
    sheetName = "Foglio1";

    //l'indice numerico(da 1 a N) dell'ultima 
    //  linea sulla quale è presente un valore
    lastRow = OdsObj.GetLastValidRow(sheetName);
    
    //l'indice numerico(da 1 a N) dell'ultima
    //  colonna sulla quale è presente un valore
    lastCol = OdsObj.GetLastValidColumn(sheetName);

    //otteniamo la lista dei fogli/tabelle presenti nel file .ODS
    sheetsList = OdsObj.GetTablesList();

    //leggiamo la cella A1
    value = OdsObj.GetCellValueText(sheetName, 1, 1);
    
    //leggiamo la cella A2
    value = OdsObj.GetCellValueText(sheetName, 2, 1);

    //leggiamo la cella B1
    value = OdsObj.GetCellValueText(sheetName, 1, 2);
```

Allo stato attuale la libreria è in una fase embrionale quindi, in futuro, arriveranno **cambiamenti**, **miglioramenti**, **nuove implementazioni** ed eventuali **correzioni**.


