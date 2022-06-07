using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace WinAppATS
{
    public partial class FormFront : Form
    {
        string ruc;
        string razonsocial;
        int anio;
        int mes;
        string info;
        int widthMenu;
        bool colapse;
        Button[] buttons;

        Button btnAntes;
        public FormFront(string ruc, string razonsocial, int anio, int mes)
        {
            InitializeComponent();
            this.ruc = ruc.PadLeft(13, '0');
            this.razonsocial = razonsocial;
            this.anio = anio;
            this.mes = mes;
            lbTitulo.Text = mes + "-" + anio + " " + razonsocial;
            this.info = ruc.PadLeft(13, '0') + anio + mes.ToString("00");
            widthMenu = pnSide.Width;
            colapse = false;
            buttons = new Button[13]{ btnIncio, bntCompras , btnVentas, btnAnulados, bntMayorAnalitico,
            btnMayAnaDet,btnComprasProv,btnComprasxTCV,btnComprasConceptoIRR,btnComprasConceptoIRD,
                btnComprasRFIVA,btnVentasXClientes,btnRCompras};
            AbrirForm<FormBienvenida>();
        }

        private void selectButton(Button button)
        {
            button.BackColor = Color.DodgerBlue;
            if (btnAntes != null)
            {
                btnAntes.BackColor = Color.Transparent;
            }
            btnAntes = button;
        }
        private void PtBoxClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PtBoxMax_Click(object sender, EventArgs e)
        {
            this.Size = Screen.PrimaryScreen.WorkingArea.Size;
            this.Location = Screen.PrimaryScreen.WorkingArea.Location;
        }

        private void PtBoxMin_Click(object sender, EventArgs e)
        {

        }

        private void AbrirForm<MiForm>() where MiForm : Form, new()
        {
            Form formulario;
            formulario = pnFront.Controls.OfType<MiForm>().FirstOrDefault(); //Busca en la coleccion de formularios
            if (formulario == null)
            {
                formulario = new MiForm();
                formulario.TopLevel = false;
                formulario.FormBorderStyle = FormBorderStyle.None;
                formulario.Dock = DockStyle.Fill;
                formulario.Name = formulario.Name + info;
                pnFront.Controls.Add(formulario);
                pnFront.Tag = formulario;
                formulario.Show();
            }
            formulario.Width = this.pnFront.Width;
            formulario.Height = this.pnFront.Height;
            formulario.BringToFront();
        }

        private void BtnVentas_Click(object sender, EventArgs e)
        {
            AbrirForm<FormVentas>();
            selectButton(btnVentas);
        }

        private void PbClose_Click(object sender, EventArgs e)
        {
            if (File.Exists(Const.filexml("a" + info)))
            {
                File.Delete(Const.filexml("a" + info));
            }
            if (File.Exists(Const.filexml("c" + info)))
            {
                File.Delete(Const.filexml("c" + info));
            }
            if (File.Exists(Const.filexml("v" + info)))
            {
                File.Delete(Const.filexml("v" + info));
            }
            this.Close();
        }

        int lx, ly;
        int sw, sh;
        private void PbMin_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void BtnProveedores_Click(object sender, EventArgs e)
        {
            AbrirForm<FormProveedores>();
            //selectButton(btnReportes);
        }

        private void BtnAnulados_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRegistroAnulados>();
            selectButton(btnAnulados);
        }

        private void BntMayorAnalitico_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRMayAnaResum>();
            selectButton(bntMayorAnalitico);
        }

        private void btnMayAnaDet_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRMayAnaDetal>();
            selectButton(btnMayAnaDet);
        }

        private void BtnComprasProv_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRComprasProveedor>();
            selectButton(btnComprasProv);
        }

        private void BtnComprasxTCV_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRResumComprasTCV>();
            selectButton(btnComprasxTCV);
        }

        private void BtnComprasConceptoIRR_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRResumForm103>();
            selectButton(btnComprasConceptoIRR);
        }

        private void BtnComprasConceptoIRD_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRDetallForm103>();
            selectButton(btnComprasConceptoIRD);
        }

        private void BtnComprasRFIVA_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRComprasRFIVADetall>();
            selectButton(btnComprasRFIVA);
        }

        private void BtnVentasXClientes_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRVentasATS>();
            selectButton(btnVentasXClientes);
        }

        private async void FormFront_Load(object sender, EventArgs e)
        {
            this.Text = razonsocial;
            using (var client = new HttpClient())
            {
                try
                {
                    var response = await client.GetAsync(Const.URL + "archivos/show?info=" + info);
                    String result = await response.Content.ReadAsStringAsync();

                    if (result != "")
                    {
                        var obj = new { fileanulado = "", filecompra = "", fileventa = "" };
                        var archivos = JsonConvert.DeserializeAnonymousType(@result, obj);

                        if (archivos.fileanulado != null)
                        {
                            byte[] abytes = Convert.FromBase64String(archivos.fileanulado);
                            File.WriteAllBytes(Const.filexml("a" + info), abytes);
                        }
                        if (archivos.filecompra != null)
                        {
                            byte[] cbytes = Convert.FromBase64String(archivos.filecompra);
                            File.WriteAllBytes(Const.filexml("c" + info), cbytes);
                        }
                        if (archivos.fileventa != null)
                        {
                            byte[] vbytes = Convert.FromBase64String(archivos.fileventa);
                            File.WriteAllBytes(Const.filexml("v" + info), vbytes);
                        }
                    }

                    bntCompras.Enabled = true;
                    btnVentas.Enabled = true;
                    btnAnulados.Enabled = true;
                    btnHideReportes.Enabled = true;
                    //btnProveedores.Enabled = true;
                    bntMayorAnalitico.Enabled = true;
                    btnMayAnaDet.Enabled = true;
                    btnComprasProv.Enabled = true;
                    btnComprasxTCV.Enabled = true;
                    btnComprasConceptoIRR.Enabled = true;
                    btnComprasConceptoIRD.Enabled = true;
                    btnComprasRFIVA.Enabled = true;
                    btnVentasXClientes.Enabled = true;
                }
                catch (Exception)
                {
                    MessageBox.Show("Error en el servidor");
                }
            }
        }

        private void PbMax_Click(object sender, EventArgs e)
        {
            lx = this.Location.X;
            ly = this.Location.Y;
            sw = this.Size.Width;
            sh = this.Size.Height;
            pbMax.Visible = false;
            pbRestor.Visible = true;
            this.Size = Screen.PrimaryScreen.WorkingArea.Size;
            this.Location = Screen.PrimaryScreen.WorkingArea.Location;
        }

        private void BtnIncio_Click(object sender, EventArgs e)
        {
            AbrirForm<FormBienvenida>();
            selectButton(btnIncio);
        }

        private void BntCompras_Click(object sender, EventArgs e)
        {
            AbrirForm<FormCompras>();
            selectButton(bntCompras);
        }

        private void BtnAmburger_Click(object sender, EventArgs e)
        {
            if (colapse)
            {
                pnSide.Width = 200;
                pnFront.Width = this.Width - 201;
                colapse = false;
            }
            else
            {
                pnSide.Width = 45;
                pnFront.Width = this.Width - 46;
                colapse = true;
            }
        }

        private void btnReportes_Click(object sender, EventArgs e)
        {
            showSubMenu(pn_reportes);
        }
        //muestra el submenu

        private void showSubMenu(Panel submenu)
        {
            if (submenu.Visible == false)
            {
                //hideSubMenu();
                submenu.Visible = true;
            }
            else
                submenu.Visible = false;
        }

        private void btnHideReportes_Click(object sender, EventArgs e)
        {
            showSubMenu(pn_reportes);
        }

        private void btnRCompras_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRCompras>();
            selectButton(btnRCompras);
        }

        private void btnRVentas_Click(object sender, EventArgs e)
        {
            AbrirForm<FormRVentas>();
            selectButton(btnRVentas);
        }

        //this dll for slider mouse to move
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]

        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {
            //Cod to move screen
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void PbRestor_Click(object sender, EventArgs e)
        {
            pbMax.Visible = true;
            pbRestor.Visible = false;
            this.Size = new Size(sw, sh);
            this.Location = new Point(lx, ly);
        }
        ~FormFront() { }
    }
}