using Newtonsoft.Json;
using System;
using System.Data;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;

namespace WinAppATS
{
    public partial class FormLogin : Form
    {
        public FormLogin()
        {
            InitializeComponent();
        }

        private void BtnIngresar_Click(object sender, EventArgs e)
        {
            login();
        }

        private async void login()
        {
            if (tbUser.Text != "" && tbPassword.Text != "")
            {
                using (var client = new HttpClient())
                {
                    try
                    {
                        var data = new { user = tbUser.Text, password = tbPassword.Text };

                        string json = JsonConvert.SerializeObject(data);
                        var content = new StringContent(json, Encoding.UTF8, "application/json");

                        var result1 = await client.PostAsync(Const.URL + "login", content);
                        string clientesAudit = await result1.Content.ReadAsStringAsync();

                        if (clientesAudit != "")
                        {
                            var anonime = new
                            {
                                success = false,
                                user = new { id = -1, name = "", user = "", rol = "", email = "" },
                                remember_token = "",
                                customers = new[] { new { ruc = "", razonsocial = "", nombrecomercial = "", phone = "", mail = "", direccion = "", diadeclaracion = "", sri = "", representantelegal = "", iess1 = "", iess2 = "", mt = "", mrl = "", super = "", contabilidad = "" } }
                            };
                            var datos = JsonConvert.DeserializeAnonymousType(@clientesAudit, anonime);
                            if (datos.success)
                            {
                                FormClientesAuditWhole formClientes = new FormClientesAuditWhole(datos.user.id, datos.remember_token);
                                formClientes.lbAsesor.Text = "Asesor " + datos.user.name;
                                int i = 0;
                                //Agrega datos al DataGridView y se guarda los clientes en un xml
                                if (datos.customers.Length > 0)
                                {
                                    DataTable dt = new DataTable("cliente");
                                    dt.Columns.Add("ruc", typeof(string));
                                    dt.Columns.Add("razonsocial", typeof(string));
                                    dt.Columns.Add("nombrecomercial", typeof(string));
                                    dt.Columns.Add("phone", typeof(string));
                                    dt.Columns.Add("mail", typeof(string));
                                    dt.Columns.Add("direccion", typeof(string));
                                    dt.Columns.Add("diadeclaracion", typeof(string));
                                    dt.Columns.Add("sri", typeof(string));
                                    dt.Columns.Add("representantelegal", typeof(string));
                                    dt.Columns.Add("iess1", typeof(string));
                                    dt.Columns.Add("mt", typeof(string));
                                    dt.Columns.Add("mrl", typeof(string));
                                    dt.Columns.Add("super", typeof(string));
                                    dt.Columns.Add("contabilidad", typeof(string));
                                    dt.Columns.Add("iess2", typeof(string));

                                    foreach (Object obj in datos.customers)
                                    {
                                        DateTime fecha = DateTime.Now;
                                        string anio = fecha.ToString("yyyy");
                                        string mes = (int.Parse(fecha.ToString("MM")) - 1).ToString("00");
                                        if (mes == "00")
                                        {
                                            mes = "12";
                                            anio = (int.Parse(anio) - 1).ToString();
                                        }
                                        formClientes.dgvCustomers.Rows.Add(datos.customers[i].ruc.PadLeft(13, '0'), datos.customers[i].razonsocial, anio, mes, "");

                                        DataRow dataRow = dt.NewRow();
                                        dataRow[0] = datos.customers[i].ruc.PadLeft(13, '0');
                                        dataRow[1] = datos.customers[i].razonsocial;
                                        dataRow[2] = datos.customers[i].nombrecomercial;
                                        dataRow[3] = datos.customers[i].phone == null ? "-" : datos.customers[i].phone;
                                        dataRow[4] = datos.customers[i].mail == null ? "-" : datos.customers[i].mail;
                                        dataRow[5] = datos.customers[i].direccion == null ? "-" : datos.customers[i].direccion;
                                        dataRow[6] = datos.customers[i].diadeclaracion == null ? "-" : datos.customers[i].diadeclaracion;
                                        dataRow[7] = datos.customers[i].sri == null ? "-" : datos.customers[i].sri;
                                        dataRow[8] = datos.customers[i].representantelegal == null ? "-" : datos.customers[i].representantelegal;
                                        dataRow[9] = datos.customers[i].iess1 == null ? "-" : datos.customers[i].iess1;
                                        dataRow[10] = datos.customers[i].mt == null ? "-" : datos.customers[i].mt;
                                        dataRow[11] = datos.customers[i].mrl == null ? "-" : datos.customers[i].mrl;
                                        dataRow[12] = datos.customers[i].super == null ? "-" : datos.customers[i].super;
                                        dataRow[13] = datos.customers[i].contabilidad == null ? "-" : datos.customers[i].contabilidad;
                                        dataRow[14] = datos.customers[i].iess2 == null ? "-" : datos.customers[i].iess2;
                                        dt.Rows.Add(dataRow);

                                        i++;
                                    }
                                    DataSet ds = new DataSet("clientes");
                                    ds.Tables.Add(dt);
                                    ds.WriteXml(Const.filexml("clientes"));
                                }
                                formClientes.Show();
                                this.Hide();
                            }
                            else
                            {
                                MessageBox.Show("Usuario o contraseña incorrecta");
                            }
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Error en el servidor");
                    }
                }
            }
            else
            {
                MessageBox.Show("Ingrese el usuario y contraseña");
            }
        }

        private void tbPassword_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                login();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
