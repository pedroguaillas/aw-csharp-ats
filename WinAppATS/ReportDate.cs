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
    class ReportDate
    {
        DateTime datestart;
        DateTime dateend;
        string info;

        public ReportDate(DateTime datestart, DateTime dateend, string info)
        {
            this.datestart = datestart;
            this.dateend = dateend;
            this.info = info;
        }

        public async Task<DataSet> ventas()
        {
            DataSet ds = new DataSet("ventas");
            try
            {
                if (updateDate())
                {
                    using (var client = new HttpClient())
                    {
                        var data = new { datestart = datestart.ToString("dd-MM-yyyy"), dateend = dateend.ToString("dd-MM-yyyy"), ruc = info.Substring(0, 13) };
                        string json = JsonConvert.SerializeObject(data);
                        var content = new StringContent(json, Encoding.UTF8, "application/json");
                        var result1 = await client.PostAsync(Const.URL + "reportventa", content);
                        string result = await result1.Content.ReadAsStringAsync();
                        if (result != null)
                        {
                            byte[] abytes = Convert.FromBase64String(result);

                            using (var stream = new MemoryStream(abytes))
                            {
                                using (XmlReader xml = XmlReader.Create(stream))
                                {
                                    ds.ReadXml(xml);
                                }
                            }
                        }
                    }
                }
                else if (File.Exists(Const.filexml("v" + info)))
                {
                    using (XmlReader xml = XmlReader.Create(Const.filexml("v" + info)))
                    {
                        ds.ReadXml(xml);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Se produjo un error al generar el reporte de ventas", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return ds;
        }

        public async Task<DataSet> compras()
        {
            DataSet ds = new DataSet("compras");
            try
            {
                if (updateDate())
                {
                    using (var client = new HttpClient())
                    {
                        var data = new { datestart = datestart.ToString("dd-MM-yyyy"), dateend = dateend.ToString("dd-MM-yyyy"), ruc = info.Substring(0, 13) };
                        string json = JsonConvert.SerializeObject(data);
                        var content = new StringContent(json, Encoding.UTF8, "application/json");
                        var result1 = await client.PostAsync(Const.URL + "reportcompra", content);
                        string result = await result1.Content.ReadAsStringAsync();
                        if (result != null)
                        {
                            byte[] abytes = Convert.FromBase64String(result);

                            using (var stream = new MemoryStream(abytes))
                            {
                                using (XmlReader xml = XmlReader.Create(stream))
                                {
                                    ds.ReadXml(xml);
                                }
                            }
                        }
                    }
                }
                else if (File.Exists(Const.filexml("c" + info)))
                {
                    using (XmlReader xml = XmlReader.Create(Const.filexml("c" + info)))
                    {
                        ds.ReadXml(xml);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Se produjo un error al generar el reporte de compras", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return ds;
        }

        private bool updateDate()
        {
            bool result = false;

            string month = info.Substring(info.Length - 2);
            string year = info.Substring(info.Length - 6, 4);

            DateTime firstDayOfMonth = new DateTime(int.Parse(year), int.Parse(month), 1);
            var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

            result = !(firstDayOfMonth == datestart && lastDayOfMonth == dateend);

            return result;
        }
    }
}
