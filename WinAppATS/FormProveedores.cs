using System;
using System.Windows.Forms;

namespace WinAppATS
{
    public partial class FormProveedores : Form
    {
        public FormProveedores()
        {
            InitializeComponent();
            addHeadDgv();
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
        }

        private void addHeadDgv()
        {
            dgvCompras.ColumnCount = 5;
            dgvCompras.Columns[0].Name = "RUC";
            dgvCompras.Columns[1].Width = 400;
            dgvCompras.Columns[1].Name = "Razon social";
            dgvCompras.Columns[2].Name = "Tipo id proveedor";
            dgvCompras.Columns[3].Name = "Tipo proveedor";
            dgvCompras.Columns[4].Name = "Obligado llevar contabilidad";

            //dgvCompras.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private async void BtnSave_Click(object sender, EventArgs e)
        {
        }
    }
}