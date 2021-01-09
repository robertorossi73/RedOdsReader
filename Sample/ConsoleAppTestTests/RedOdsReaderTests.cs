using Microsoft.VisualStudio.TestTools.UnitTesting;
using ROdsLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROdsLib.Tests
{
    [TestClass()]
    public class RedOdsReaderTests
    {
        private RedOdsReader ObjReader;

        /// <summary>
        /// Apre il file ODS per i test
        /// </summary>
        private bool OpenODSFIle()
        {
            ObjReader = new RedOdsReader();

            return ObjReader.LoadFile("..\\..\\..\\citta_italia.ods");
        }

        [TestMethod()]
        public void GetLastValidColumnTest()
        {
            string sheetName;
            int lastCol;
            this.OpenODSFIle();
            sheetName = "Foglio1";

            //l'indice numerico(da 1 a N) dell'ultima
            //  colonna sulla quale è presente un valore
            lastCol = ObjReader.GetLastValidColumn(sheetName);
            if (lastCol != 3)
                Assert.Fail();
        }

        [TestMethod()]
        public void GetLastValidRowTest()
        {
            string sheetName;
            int lastRow;
            this.OpenODSFIle();
            sheetName = "Foglio1";

            //l'indice numerico(da 1 a N) dell'ultima 
            //  linea sulla quale è presente un valore
            lastRow = ObjReader.GetLastValidRow(sheetName);
            if (lastRow != 561)
                Assert.Fail();
        }

        [TestMethod()]
        public void GetTablesListTest()
        {
            List<string> sheetsList;
            this.OpenODSFIle();

            //otteniamo la lista dei fogli/tabelle presenti nel file .ODS
            sheetsList = ObjReader.GetTablesList();
            if (
                sheetsList[0] != "Foglio1" ||
                sheetsList[1] != "FoglioB" ||
                sheetsList.Count != 2
                )
            {
                Assert.Fail();
            }
        }

        [TestMethod()]
        public void GetColumnIndexTest()
        {
            this.OpenODSFIle();
            int i;

            i = ObjReader.GetColumnIndex("A");
            if (i != 1)
                Assert.Fail();
            i = ObjReader.GetColumnIndex("AF");
            if (i != 32)
                Assert.Fail();
        }

        [TestMethod()]
        public void GetColumnNameTest()
        {
            this.OpenODSFIle();
            string i;

            i = ObjReader.GetColumnName(1);
            if (i != "A")
                Assert.Fail();
            i = ObjReader.GetColumnName(32);
            if (i != "AF")
                Assert.Fail();

        }

        [TestMethod()]
        public void ExistTableTest()
        {
            bool a;
            bool b;
            this.OpenODSFIle();

            a = ObjReader.ExistTable("Foglio1");
            b = ObjReader.ExistTable("FoglioB");
            if(!a || !b)
                Assert.Fail();
        }

        [TestMethod()]
        public void GetCellValueTextTest()
        {
            this.OpenODSFIle();
            string val;

            val = ObjReader.GetCellValueText("Foglio1", 5, 1);
            if (val != "Acri")
                Assert.Fail();

            val = ObjReader.GetCellValueText("FoglioB", 11, 3);
            if (val != "602,501")
                Assert.Fail("Error value: " + val);
        }

        [TestMethod()]
        public void GetCellValueTextToFloatTest()
        {
            this.OpenODSFIle();
            decimal val;

            val = ObjReader.GetCellValueTextAsDecimal("FoglioB", 17, 1);
            if (!(val.CompareTo(new decimal(153.605)) == 0))
                Assert.Fail("Error value: " + val.ToString());

            val = ObjReader.GetCellValueTextAsDecimal("FoglioB", 18, 1);
            if (!(val.CompareTo(new decimal(11.4553)) == 0))
                Assert.Fail("Error value: " + val.ToString());

            val = ObjReader.GetCellValueTextAsDecimal("FoglioB", 19, 1);
            if (!(val.CompareTo(new decimal(657)) == 0))
                Assert.Fail("Error value: " + val.ToString());

            val = ObjReader.GetCellValueTextAsDecimal("Foglio1", 5, 1);
            if (!(val.CompareTo(new decimal(0)) == 0))
                Assert.Fail("Error value: " + val.ToString());

            val = ObjReader.GetCellValueTextAsDecimal("FoglioB", 11, 3);
            if (!(val.CompareTo(new decimal(602.501)) == 0))
                Assert.Fail("Error value: " + val.ToString());
        }

        [TestMethod()]
        public void LoadFileTest()
        {
            if (!this.OpenODSFIle())
                Assert.Fail();
        }
    }
}