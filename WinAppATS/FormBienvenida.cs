using Newtonsoft.Json;
using System;
using System.Data;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;

namespace WinAppATS
{
    public partial class FormBienvenida : Form
    {
        public FormBienvenida()
        {
            InitializeComponent();
        }

        private void BtnExportar_Click(object sender, EventArgs e)
        {
            GenererateAtsXml xml = new GenererateAtsXml(this.Name.Substring(this.Name.Length - 19), tbEstablecimiento.Text);
            xml.generate();
            //Archivo archivo = new Archivo();
            //archivo.generarATS(this.Name.Substring(this.Name.Length - 19), tbEstablecimiento.Text);
        }

        private void FormBienvenida_Load(object sender, EventArgs e)
        {
            Cliente cliente = new Cliente();
            tbRuc.Text = this.Name.Substring(this.Name.Length - 19, 13);
            DataRow dataRow = cliente.getClienteRow(this.Name.Substring(this.Name.Length - 19, 13));
            if (dataRow != null)
            {
                tbRazonSocial.Text = dataRow["razonsocial"].ToString();
                tbAnio.Text = this.Name.Substring(this.Name.Length - 6, 4);
                tbMes.Text = this.Name.Substring(this.Name.Length - 2);
                tbEstablecimiento.Text = "001";
                tbTelefono.Text = dataRow["phone"].ToString().Equals("-") ? "" : dataRow["phone"].ToString();
                tbCorreo.Text = dataRow["mail"].ToString().Equals("-") ? "" : dataRow["mail"].ToString();
                tbDireccion.Text = dataRow["direccion"].ToString().Equals("-") ? "" : dataRow["direccion"].ToString();
                tbDiaDeclaracion.Text = dataRow["diadeclaracion"].ToString().Equals("-") ? "" : dataRow["diadeclaracion"].ToString();
                tbRepresentanteLegal.Text = dataRow["representantelegal"].ToString().Equals("-") ? "" : dataRow["representantelegal"].ToString();
                tbSri.Text = dataRow["sri"].ToString().Equals("-") ? "" : dataRow["sri"].ToString();
                tbIess1.Text = dataRow["iess1"].ToString().Equals("-") ? "" : dataRow["iess1"].ToString();
                tbMt.Text = dataRow["mt"].ToString().Equals("-") ? "" : dataRow["mt"].ToString();
                tbMrl.Text = dataRow["mrl"].ToString().Equals("-") ? "" : dataRow["mrl"].ToString();
                tbSuper.Text = dataRow["super"].ToString().Equals("-") ? "" : dataRow["super"].ToString();
                tbContabilidad.Text = dataRow["contabilidad"].ToString().Equals("-") ? "" : dataRow["contabilidad"].ToString();
                tbIess2.Text = dataRow["iess2"].ToString().Equals("-") ? "" : dataRow["iess2"].ToString();
            }
        }

        private async void BtnGuardar_Click(object sender, EventArgs e)
        {
            string result;
            var customer = new
            {
                ruc = tbRuc.Text,
                razonsocial = tbRazonSocial.Text,
                phone = tbTelefono.Text.Equals("") ? null : tbTelefono.Text,
                mail = tbCorreo.Text.Equals("") ? null : tbCorreo.Text,
                direccion = tbDireccion.Text.Equals("") ? null : tbDireccion.Text,
                diadeclaracion = tbDiaDeclaracion.Text.Equals("") ? null : tbDiaDeclaracion.Text,
                sri = tbSri.Text.Equals("") ? null : tbSri.Text,
                representantelegal = tbRepresentanteLegal.Text.Equals("") ? null : tbRepresentanteLegal.Text,
                iess1 = tbIess1.Text.Equals("") ? null : tbIess1.Text,
                iess2 = tbIess2.Text.Equals("") ? null : tbIess2.Text,
                mt = tbMt.Text.Equals("") ? null : tbMt.Text,
                mrl = tbMrl.Text.Equals("") ? null : tbMrl.Text,
                super = tbSuper.Text.Equals("") ? null : tbSuper.Text,
                contabilidad = tbContabilidad.Text.Equals("") ? null : tbContabilidad.Text
            };

            using (var client = new HttpClient())
            {
                string json = JsonConvert.SerializeObject(customer);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                try
                {
                    var response = await client.PostAsync(Const.URL + "customers/update", content);
                    result = await response.Content.ReadAsStringAsync();
                    if (result.Contains("OK"))
                    {
                        MessageBox.Show("Se guardo satisfactoriamente");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar los datos" + ex.Message);
                }
            }
        }

        private void tbRazonSocial_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.V && e.Control)
                e.SuppressKeyPress = false;
        }

        private void TbRazonSocial_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar >= 'a' && e.KeyChar <= 'z') || (e.KeyChar >= 'A' && e.KeyChar <= 'Z') || e.KeyChar == (char)Keys.Back || e.KeyChar == ' ')
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void TbEstablecimiento_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void TbTelefono_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void TbCorreo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar >= 'a' && e.KeyChar <= 'z') || (e.KeyChar >= 'A' && e.KeyChar <= 'Z') || e.KeyChar == '@' || e.KeyChar == '.' || e.KeyChar == '_' || e.KeyChar == '-' || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void TbDireccion_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar >= 'a' && e.KeyChar <= 'z') || (e.KeyChar >= 'A' && e.KeyChar <= 'Z') || e.KeyChar == '.' || e.KeyChar == (char)Keys.Back || e.KeyChar == ' ')
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void TbDiaDeclaracion_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void TbRepresentanteLegal_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= 'a' && e.KeyChar <= 'z') || (e.KeyChar >= 'A' && e.KeyChar <= 'Z') || e.KeyChar == (char)Keys.Back || e.KeyChar == (char)Keys.Escape)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void TbContabilidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 'S' || e.KeyChar == 'I' || e.KeyChar == 'N' || e.KeyChar == 'O' || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void TbSri_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                tbIess1.Focus();
            }
        }

        private void TbIess1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                tbIess2.Focus();
            }
        }

        private void TbIess2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                tbMt.Focus();
            }
        }

        private void TbMt_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                tbMrl.Focus();
            }
        }

        private void TbMrl_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                tbSuper.Focus();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.E))
            {
                Archivo archivo = new Archivo();
                archivo.generarATS(this.Name.Substring(this.Name.Length - 19), tbEstablecimiento.Text);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}