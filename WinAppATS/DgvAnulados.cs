using Newtonsoft.Json;
using System;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace WinAppATS
{
    class DgvAnulados
    {
        DataGridView dgv;
        String info;

        public DgvAnulados(DataGridView dgv)
        {
            this.dgv = dgv;

            //DataSet data = new DataSet("anulados");
            //DataTable dt = new DataTable("detalleAnulados");

            ////Info comprobante
            //dt.Columns.Add(new DataColumn("TipoComprobante"));
            //dt.Columns.Add(new DataColumn("Establecimiento"));
            //dt.Columns.Add(new DataColumn("PuntotoEmision"));
            //dt.Columns.Add(new DataColumn("SecuencialIncio"));
            //dt.Columns.Add(new DataColumn("SecuencialFin"));
            //dt.Columns.Add(new DataColumn("Autorizacion"));

            //data.Tables.Add(dt);

            //this.dgv.DataSource = data.Tables[0];
        }

        public async void Load(string info)
        {
            this.info = info;
            if (File.Exists(Const.filexml(info)))
            {
                using (XmlReader xml = XmlReader.Create(Const.filexml(info)))
                {
                    DataSet data = new DataSet("anulados");
                    data.ReadXml(xml);
                    DataTable dt = data.Tables[0];

                    foreach (DataRow dr in dt.Rows)
                    {
                        dgv.Rows.Add(dr.ItemArray);
                    }
                }
            }
        }

        public async Task<string> toXml()
        {
            if (dgv.Rows.Count > 0)
            {
                DataTable dt = new DataTable("anulado");
                for (int i = 1; i < dgv.Columns.Count + 1; i++)
                {
                    string column = dgv.Columns[i - 1].Name;
                    dt.Columns.Add(column, typeof(string));
                }

                int ColumnCount = dgv.Columns.Count;
                foreach (DataGridViewRow dr in dgv.Rows)
                {
                    DataRow dataRow = dt.NewRow();
                    for (int i = 0; i < ColumnCount; i++)
                    {
                        dataRow[i] = dr.Cells[i].Value;
                    }
                    dt.Rows.Add(dataRow);
                }
                DataSet ds = new DataSet("anulados");
                ds.Tables.Add(dt);
                ds.WriteXml(Const.filexml(info));

                return await Guardar();
            }
            return "0";
        }

        private async Task<string> Guardar()
        {
            string result;
            byte[] bytes = File.ReadAllBytes(Const.filexml(info));

            using (var client = new HttpClient())
            {
                var data = new { file = Convert.ToBase64String(bytes), info = info.Substring(1), tipo = "fileanulado" };

                string json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var result1 = await client.PostAsync(Const.URL + "archivos", content);
                result = await result1.Content.ReadAsStringAsync();
            }
            return result;
        }
    }
}