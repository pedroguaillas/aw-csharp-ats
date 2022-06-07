using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinAppATS
{
    class ImportContactos : Importa
    {
        public async void registrarList(List<Object> contactos)
        {
            using (var client = new HttpClient())
            {
                var data = new { contactos = contactos };
                string json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                try
                {
                    var resultP = await client.PostAsync(Const.URL + "contactos", content);
                    var result = await resultP.Content.ReadAsStringAsync();
                }
                catch (Exception)
                {
                    MessageBox.Show("Se genero un error al guardar los contactos");
                }
            }
        }
        public async void registrarUno(Object contacto)
        {
            using (var client = new HttpClient())
            {
                var data = new { contacto = contacto };
                string json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                try
                {
                    var resultP = await client.PostAsync(Const.URL + "contactos/store", content);
                    var result = await resultP.Content.ReadAsStringAsync();
                }
                catch (Exception)
                {
                    MessageBox.Show("Error en la conexión con los servicios web");
                }
            }
        }

        public async Task<List<Contacto>> getContactMasive(List<string> providers)
        {
            List<Contacto> contactos = null;

            using (var client = new HttpClient())
            {
                var data = new { providers };
                string json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                try
                {
                    var response = await client.PostAsync(Const.URL + "contactosgetmasive", content);
                    var result = await response.Content.ReadAsStringAsync();

                    if (result != "")
                    {
                        contactos = new List<Contacto>();
                        contactos = JsonConvert.DeserializeObject<List<Contacto>>(result);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Error en la conexión con los servicios web");
                }
            }

            return contactos;
        }

        public async Task<Contacto> getContact(string id)
        {
            Contacto contacto = null;

            using (var client = new HttpClient())
            {
                try
                {
                    var resultP = await client.GetAsync(Const.URL + "contactos/" + id);
                    String res = await resultP.Content.ReadAsStringAsync();

                    if (res != "")
                    {
                        contacto = JsonConvert.DeserializeObject<Contacto>(res);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Error en el servidor");
                }
            }
            return contacto;
        }

        ~ImportContactos() { }
    }
}