using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;

namespace WinAppATS
{
    class Archivo
    {

        public void saveError(string error, string path)
        {
            FileInfo fi = new FileInfo(path);
            // Check if file already exists. If yes, delete it.     
            if (fi.Exists)
            {
                fi.Delete();
            }

            // Create a new file     
            using (StreamWriter sw = fi.CreateText())
            {
                sw.WriteLine(error.Trim());
            }
            MessageBox.Show("El Servicio Web del SRI no permitio la descarga de todos los comprobantes, por favor revisa el archivo " + path + " de los comprobantes que no se pudo descargar");
        }

        public async void generarATS(string info, string establecimiento)
        {
            using (var client = new HttpClient())
            {
                try
                {
                    var response = await client.GetAsync(Const.URL + "archivos?info=" + info + "&establecimiento=" + establecimiento);
                    String result = await response.Content.ReadAsStringAsync();

                    if (result != "")
                    {
                        string path = miCarpeta("xml", "ATS");
                        if (path != "")
                        {
                            byte[] bytes = Convert.FromBase64String(result);
                            File.WriteAllBytes(path, bytes);
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Error al generar el ATS");
                }
            }
        }

        public async void downloadVouchers(List<String> clave_accs, string ruc)
        {
            using (var client = new HttpClient())
            {
                try
                {
                    var data = new { clave_accs, ruc };

                    string json = JsonConvert.SerializeObject(data);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(Const.URL + "downloadzip", content);
                    string result = await response.Content.ReadAsStringAsync();

                    if (result != "" && response.IsSuccessStatusCode)
                    {
                        string path = miCarpeta("zip", "comprobantes");
                        if (path != "")
                        {
                            byte[] bytes = Convert.FromBase64String(result);
                            File.WriteAllBytes(path, bytes);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Se genero un error ar descargar los comprobantes", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Se genero un error ar descargar los comprobantes" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        protected string miCarpeta(string type, string name)
        {
            string carpeta = "";

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "Archivos de datos (*." + type + ")|*." + type;
            dlg.DefaultExt = type;
            dlg.FileName = name;
            dlg.AddExtension = true;
            DialogResult result = dlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                carpeta = dlg.FileName;
            }

            return carpeta;
        }
    }
}