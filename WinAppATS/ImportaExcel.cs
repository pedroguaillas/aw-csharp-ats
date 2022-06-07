using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WinAppATS
{
    class ImportaExcel : ImportaFile
    {
        static DsAts db;
        string filename = string.Format("{0}//data.dat", System.Windows.Forms.Application.StartupPath);

        protected static DsAts App
        {
            get
            {
                if (db == null)
                {
                    db = new DsAts();
                }
                return db;
            }
        }

        public void Parse()
        {
            selectFile("xlsm");
            if (pathFile != null)
            {
                _Application excel = new _Excel.Application();

                Workbook wb = excel.Workbooks.Open(pathFile);

                Worksheet ws = wb.Worksheets["PC_AWCA"];

                //if (File.Exists(filename))
                //{
                //    App.DtCuentas.ReadXml(filename);

                //}

                DsAts data = new DsAts();
                //DsAts.DtCuentasRow row = new DsAts.DtCuentasRow();

                //DataRow row = data.Tables["DtCuentas"].NewRow();
                //row["code"] = ws.Cells[4][1].Value2;
                //row["description"] = ws.Cells[4][2].Value2;
                //row["tst"] = ws.Cells[4][3].Value2;
                //data.Tables["DtCuentas"].Rows.Add(row);


                for (int i = 4; i < 43; i++)
                {
                    DataRow row = App.Tables["DtCuentas"].NewRow();
                    row["code"] = ws.Cells[i][1].Value2;
                    row["description"] = ws.Cells[i][2].Value2;
                    row["tst"] = ws.Cells[i][3].Value2;
                    App.Tables["DtCuentas"].Rows.Add(row);


                    //DtCuentas dt = data.DtCuentas;
                    //data.Insert(ws.Cells[i][1].Value2, ws.Cells[i][2].Value2, ws.Cells[i][3].Value2);
                    //dt.Rows.Add(ws.Cells[i][1].Value2, ws.Cells[i][2].Value2, ws.Cells[i][3].Value2);
                    //dt.AcceptChanges();
                }
                App.WriteXml(filename);
            }
        }

        public void testCorrectLoad()
        {
            //DsAts data = new DsAts();
            //var dt = data.DtCuentas;
            //MessageBox.Show(dt.Rows[0][1].ToString());
        }

        public void Remove()
        {
            //DsAts data = new DsAts();
            //var dt = data.DtCuentas;
            //dt.Rows.Remove(data.DtCuentas.Rows[1]);
            //dt.AcceptChanges();
        }
    }
}