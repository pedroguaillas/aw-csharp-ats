using System;
using System.Data;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace WinAppATS
{
    public partial class FormClientesAuditWhole : Form
    {
        int idAsesor;
        string remember_token;
        string[,] archivos;

        public FormClientesAuditWhole(int idAsesor, string remember_token)
        {
            InitializeComponent();
            this.idAsesor = idAsesor;
            this.remember_token = remember_token;
            this.archivos = new string[10, 3];
        }

        int arc = 0;

        private void DgvCustomers_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (arc < 10)
            {
                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
                {
                    archivos[arc, 0] = dgvCustomers.Rows[e.RowIndex].Cells[0].Value.ToString();
                    archivos[arc, 1] = dgvCustomers.Rows[e.RowIndex].Cells[2].Value.ToString();
                    archivos[arc, 2] = dgvCustomers.Rows[e.RowIndex].Cells[3].Value.ToString();
                    FormFront front = new FormFront(
                        dgvCustomers.Rows[e.RowIndex].Cells[0].Value.ToString(),
                        dgvCustomers.Rows[e.RowIndex].Cells[1].Value.ToString(),
                        int.Parse(dgvCustomers.Rows[e.RowIndex].Cells[2].Value.ToString()),
                        int.Parse(dgvCustomers.Rows[e.RowIndex].Cells[3].Value.ToString())
                    );
                    front.Show();
                    arc++;
                }
            }
            else
            {
                MessageBox.Show("Cierre el programa y vuelva abrir");
            }
        }

        private void PbClose_Click(object sender, EventArgs e)
        {
            foreach (String cli in archivos)
            {

            }
            Application.Exit();
        }

        private void PbMin_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            (dgvCustomers.DataSource as DataTable).DefaultView.RowFilter = string.Format("ruc = '{0}'", tbBuscar.Text);
            //for (int i = 0; i < dgvCustomers.Rows.Count; i++)
            //{
            //    dgvCustomers.Rows[i].C
            //}
        }

        //this dll for slider mouse to move
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]

        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void pnHead_MouseDown(object sender, MouseEventArgs e)
        {
            //Cod to move screen
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
    }
}