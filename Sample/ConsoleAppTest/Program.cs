using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppTest
{
    class Program
    {
        static void Main(string[] args)
        {

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
            OdsObj.LoadFile("citta_italia.ods");

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
            Console.WriteLine("Cella A1= " + value);

            //leggiamo la cella A2
            value = OdsObj.GetCellValueText(sheetName, 2, 1);
            Console.WriteLine("Cella A2= " + value);

            //leggiamo la cella B1
            value = OdsObj.GetCellValueText(sheetName, 1, 2);
            Console.WriteLine("Cella B1= " + value);

            Console.WriteLine("\nTest concluso.\nPremere tasto per terminare.");
            Console.ReadKey();
        }
    }
}
