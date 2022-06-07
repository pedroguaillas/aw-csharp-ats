using System;
using System.Windows.Forms;

namespace WinAppATS
{
    public partial class FormRegistroAnulados : Form
    {
        DgvAnulados dgv;
        public FormRegistroAnulados()
        {
            InitializeComponent();
            dgv = new DgvAnulados(dgvAnulados);
        }

        private void FormRegistroAnulados_Load(object sender, EventArgs e)
        {
            dgv.Load("a" + this.Name.Substring(this.Name.Length - 19));
        }

        private async void BtnGuardar_Click(object sender, EventArgs e)
        {
            await dgv.toXml();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            dgvAnulados.Rows.Insert(dgvAnulados.Rows.Count);
        }

        private void DgvAnulados_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    CopyPaste copyPaste = new CopyPaste(dgvAnulados);
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            copyPaste.CopyToClipboard();
                            break;

                        case Keys.V:
                            copyPaste.PasteClipboard("anulados");
                            break;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Fuera de rango ", "Copia y pega", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

    }
}
