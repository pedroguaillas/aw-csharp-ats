using System;
using System.Windows.Forms;
using System.Globalization;
using System.IO;
using WinAppATS.Class;

namespace WinAppATS
{
    public partial class FormVentas : Form
    {
        DgvVentas dgv;
        NumberFormatInfo nfi;
        Panel pnAnterior;

        public FormVentas()
        {
            InitializeComponent();
            dgv = new DgvVentas(dgvVentas);
            cbTpId.SelectedIndex = 1;
            nfi = NumberFormatInfo.CurrentInfo;
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            dgv.Importar();
            sumcolumns();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            save();
        }

        private async void save()
        {
            try
            {
                if (dgvVentas.Rows.Count > 0)
                {
                    btnSave.Enabled = false;
                    string res = await dgv.toXml();
                    if (res.Contains("OK"))
                    {
                        btnSave.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se produjo un error al guardar los datos" + ex.Message);
                btnSave.Enabled = true;
            }
        }

        private void FormVentas_Load(object sender, EventArgs e)
        {
            dgv.Load("v" + this.Name.Substring(this.Name.Length - 19));
            sumcolumns();
        }

        private void TbRuc_KeyPress(object sender, KeyPressEventArgs e)
        {
            int tpId = cbTpId.SelectedIndex;
            if (tpId < 2)
            {
                e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
            }
            else if (tpId == 2)
            {
                var regex = new System.Text.RegularExpressions.Regex(@"[^a-zA-Z0-9\s]");
                if (e.KeyChar != (char)Keys.Back && regex.IsMatch(e.KeyChar.ToString()))
                {
                    e.Handled = true;
                }
            }
            else if (tpId == 3)
            {
                if (e.KeyChar != (char)Keys.Back && e.KeyChar != '9')
                {
                    e.Handled = true;
                }
            }
        }

        string contabilidad = "";

        private void TbRuc_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Control | Keys.V))
                (sender as TextBox).Paste();

            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                int tpId = cbTpId.SelectedIndex;
                string idContacto = tbRuc.Text;

                if (tpId == 1)
                {
                    Validate validate = new Validate();
                    if (validate.validateMessage(tpId, idContacto))
                    {
                        return;
                    }
                }

                tbDia.Focus();
            }
        }

        private void TbDia_KeyPress(object sender, KeyPressEventArgs e)
        {
            ValidadorFecha fecha = new ValidadorFecha();
            fecha.validarmes(e, this.Name, tbDia.Text);
        }

        private void TbDia_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                tbEstablecimiento.Focus();
            }
        }

        private void TbEstablecimineto_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void TbEstablecimineto_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                tbPtoEmi.Focus();
            }
        }

        private void TbPtoEmi_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void TbPtoEmi_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                tbSecuencial.Focus();
            }
        }

        private void TbSecuencial_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void TbSecuencial_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                cbTipoComprobante.Focus();
            }
        }

        private void cbTipoComprobante_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                tbBase0.Focus();
            }
        }

        private void TbBase0_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericHandle(e);
            //e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back || e.KeyChar == '.' || e.KeyChar == ',');
        }

        private void numericHandle(KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
                e.Handled = false;
                return;
            }
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back || e.KeyChar == ',');
        }

        private void TbBase0_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                tbBase12.Focus();
            }
        }

        private void TbBase12_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericHandle(e);
        }

        private void TbBase12_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                btnAdd.Focus();
            }
        }

        private void BtnAdd_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Right)
            {
                addVenta();
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            addVenta();
        }

        private void addVenta()
        {
            string tcv = cbTipoComprobante.SelectedIndex == 0 ? "F" : (cbTipoComprobante.SelectedIndex == 1 ? "N/C" : "");
            dgv.addOne(tbRuc.Text, tbRazonS.Text, tcv, tbDia.Text, tbEstablecimiento.Text, tbPtoEmi.Text, tbSecuencial.Text, tbBase0.Text.ToString(nfi), tbBase12.Text.ToString(nfi));
            cleanInsert();
            cbTpId.Focus();
            dgvVentas.FirstDisplayedScrollingRowIndex = dgvVentas.Rows.Count - 1;
            sumcolumns();
            return;
        }

        private void cleanInsert()
        {
            cbTpId.SelectedIndex = 1;
            tbRuc.Text = "";
            tbRazonS.Text = "";
            tbDia.Text = "";
            tbEstablecimiento.Text = cbEstablecimiento.Checked ? tbEstablecimiento.Text : "";
            tbPtoEmi.Text = cbPtoEmi.Checked ? tbPtoEmi.Text : "";
            tbSecuencial.Text = rbIncrementa.Checked ? (int.Parse(tbSecuencial.Text) + 1).ToString() : rbDecrementa.Checked && int.Parse(tbSecuencial.Text) > 0 ? (int.Parse(tbSecuencial.Text) - 1).ToString() : "";
            tbBase0.Text = "";
            tbBase12.Text = "";
        }

        private void BtnImportRetenciones_Click(object sender, EventArgs e)
        {
            dgv.importRetenciones();
            sumcolumns();
        }

        private void TbRazonS_KeyPress(object sender, KeyPressEventArgs e)
        {
            var regex = new System.Text.RegularExpressions.Regex(@"[^a-zA-Z\s]");
            if (e.KeyChar != (char)Keys.Back && regex.IsMatch(e.KeyChar.ToString()))
            {
                e.Handled = true;
            }
        }

        private void BtnAddCustomer_Click(object sender, EventArgs e)
        {
            int tpId = cbTpId.SelectedIndex;
            if (tpId == 2)
            {
                var obj = new { id = tbRuc.Text, denominacion = tbRazonS.Text, tpId = (tpId + 1).ToString("00"), contabilidad };
                ImportContactos import = new ImportContactos();
                import.registrarUno(obj);
            }
        }

        int searchRow = 0;

        private void btnSearch_Click(object sender, EventArgs e)
        {
            search();
        }

        private void search()
        {
            string searchValue = tbSearchValue.Text;
            int column = cbColumn.SelectedIndex;

            try
            {
                DataGridViewRow row;
                for (; searchRow < dgvVentas.Rows.Count; searchRow++)
                {
                    row = dgvVentas.Rows[searchRow];
                    if (row.Cells[column].Value.ToString().Equals(searchValue) || (column == 5 && int.Parse(row.Cells[column].Value.ToString().Substring(row.Cells[column].Value.ToString().Length - 9)) == int.Parse(searchValue)))
                    {
                        row.Selected = true;
                        dgvVentas.FirstDisplayedScrollingRowIndex = row.Index;
                        searchRow++;
                        break;
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
            btnRepetSearch.Enabled = true;
        }

        private void btnInsertRowBefore_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvVentas.SelectedRows.Count > 0)
                {
                    DataGridViewRow row = dgvVentas.SelectedRows[0];
                    //.................Clonar la fila
                    DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                    for (Int32 index = 0; index < row.Cells.Count; index++)
                    {
                        clonedRow.Cells[index].Value = (index > 5 && index < 16) ? "0" : "";
                    }
                    dgvVentas.Rows.Insert(row.Index, clonedRow);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No se ha seleccionado una fila");
            }
        }

        private void btnInsertRowAfter_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvVentas.SelectedRows.Count > 0)
                {
                    DataGridViewRow row = dgvVentas.SelectedRows[0];
                    //.................Clonar la fila
                    DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                    for (Int32 index = 0; index < row.Cells.Count; index++)
                    {
                        clonedRow.Cells[index].Value = (index > 5 && index < 16) ? "0" : "";
                    }
                    dgvVentas.Rows.Insert(row.Index + 1, clonedRow);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No se ha seleccionado una fila");
            }
        }

        private void DgvVentas_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    CopyPaste copyPaste = new CopyPaste(dgvVentas);
                    switch (e.KeyCode)
                    {
                        case Keys.C:
                            copyPaste.CopyToClipboard();
                            break;
                        case Keys.V:
                            copyPaste.PasteClipboard("ventas");
                            break;
                        case Keys.G:
                            save();
                            break;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Fuera de rango ", "Copia y pega", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnRepetSearch_Click(object sender, EventArgs e)
        {
            searchRow = 0;
        }

        private void CbColumn_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbColumn.SelectedIndex > -1)
            {
                tbSearchValue.Enabled = true;
                tbSearchValue.Text = "";
                searchRow = 0;
                btnRepetSearch.Enabled = false;
            }
        }

        private void TbSizeFichar_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void BtnClean_Click(object sender, EventArgs e)
        {
            cleanInsert();
        }

        private void DgvVentas_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if (e.ColumnIndex > 5 && e.ColumnIndex < 16)
            {
                if (e != null)
                {
                    if (e.Value != null)
                    {
                        try
                        {
                            string valor = e.Value.ToString();
                            e.Value = valor.Replace('.', ',');
                            e.ParsingApplied = true;
                        }
                        catch (FormatException)
                        {
                            e.ParsingApplied = false;
                        }
                    }
                }
            }
            //Fecha
            //else if (e.ColumnIndex == 3)
            //{
            //    if (e != null)
            //    {
            //        if (e.Value != null)
            //        {
            //            try
            //            {
            //                string valor = e.Value.ToString();
            //                e.Value = valor.Replace('.', ',');
            //                e.ParsingApplied = true;
            //            }
            //            catch (FormatException)
            //            {
            //                e.ParsingApplied = false;
            //            }
            //        }
            //    }
            //}
        }

        private void DgvVentas_CellStateChanged(object sender, DataGridViewCellStateChangedEventArgs e)
        {
            try
            {
                int column = e.Cell.ColumnIndex;

                switch (column)
                {
                    case 0:
                        calTicl(e.Cell.RowIndex);
                        break;
                    case 7:
                        calBaseImp(e.Cell.RowIndex);
                        calTotal(e.Cell.RowIndex);
                        calValRetRenta(e.Cell.RowIndex);
                        calValRetIva(e.Cell.RowIndex);
                        break;
                    case 8:
                        calBaseImp(e.Cell.RowIndex);
                        calIva(e.Cell.RowIndex);
                        calTotal(e.Cell.RowIndex);
                        calValRetRenta(e.Cell.RowIndex);
                        calValRetIva(e.Cell.RowIndex);
                        break;
                    case 9:
                        calTotal(e.Cell.RowIndex);
                        break;
                    case 10:
                        calTotal(e.Cell.RowIndex);
                        break;
                    case 12:
                        calValRetRenta(e.Cell.RowIndex);
                        break;
                    case 13:
                        calValRetIva(e.Cell.RowIndex);
                        break;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Digitar del acuerdo al campo correspondiente");
            }
        }

        private void calTicl(int row)
        {
            if (cbCalculable.Checked)
            {
                int lentId = dgvVentas.Rows[row].Cells[0].Value.ToString().Length;
                switch (lentId)
                {
                    case 10:
                        dgvVentas.Rows[row].Cells[2].Value = "05";
                        break;
                    case 13:
                        dgvVentas.Rows[row].Cells[2].Value = dgvVentas.Rows[row].Cells[0].Value.ToString().Equals("9999999999999") ? "07" : "04";
                        break;
                    default:
                        dgvVentas.Rows[row].Cells[2].Value = "06";
                        break;
                }
            }
        }

        private void calBaseImp(int row)
        {
            if (cbCalculable.Checked)
            {
                dgvVentas.Rows[row].Cells[6].Value = double.Parse(dgvVentas.Rows[row].Cells[7].Value.ToString()) + double.Parse(dgvVentas.Rows[row].Cells[8].Value.ToString());
            }
        }

        private void calIva(int row)
        {
            if (cbCalculable.Checked)
            {
                dgvVentas.Rows[row].Cells[10].Value = Const.round(double.Parse(dgvVentas.Rows[row].Cells[8].Value.ToString()) * Const.IVA);
            }
        }

        private void calTotal(int row)
        {
            if (cbCalculable.Checked)
            {
                dgvVentas.Rows[row].Cells[11].Value = double.Parse(dgvVentas.Rows[row].Cells[6].Value.ToString()) + double.Parse(dgvVentas.Rows[row].Cells[9].Value.ToString()) + double.Parse(dgvVentas.Rows[row].Cells[10].Value.ToString());
            }
        }

        private void calValRetRenta(int row)
        {
            if (double.Parse(dgvVentas.Rows[row].Cells[6].Value.ToString()) > 0 && cbCalculable.Checked)
            {
                dgvVentas.Rows[row].Cells[14].Value = Const.round(double.Parse(dgvVentas.Rows[row].Cells[6].Value.ToString()) * .01 * double.Parse(dgvVentas.Rows[row].Cells[12].Value.Equals("") ? "0" : dgvVentas.Rows[row].Cells[12].Value.ToString()));
            }
        }

        private void calValRetIva(int row)
        {
            if (double.Parse(dgvVentas.Rows[row].Cells[10].Value.ToString()) > 0 && cbCalculable.Checked)
            {
                dgvVentas.Rows[row].Cells[15].Value = Const.round(double.Parse(dgvVentas.Rows[row].Cells[10].Value.ToString()) * .01 * double.Parse(dgvVentas.Rows[row].Cells[13].Value.Equals("") ? "0" : dgvVentas.Rows[row].Cells[13].Value.ToString()));
            }
        }

        private void CbRetExt_CheckedChanged(object sender, EventArgs e)
        {
            dgvVentas.Columns[16].Visible = cbRetExt.Checked;
        }

        private void tbSearchValue_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && cbColumn.SelectedIndex > -1)
            {
                search();
            }
        }

        private void btnRecuperado_Click(object sender, EventArgs e)
        {
            dgv.importReport(pgbImport);
        }

        private void btnIRecuperado_Click(object sender, EventArgs e)
        {
            //dgv.importRecuperado();
            dgv.importRecuperadoMasivo();
            sumcolumns();
        }

        private void cbTpId_SelectedIndexChanged(object sender, EventArgs e)
        {
            int type = (int)cbTpId.SelectedIndex;
            //MaxLenght if select cedula 10 other case 13
            tbRuc.MaxLength = type == 1 ? 10 : 13;

            if (type == 1)
            {
                tbRuc.Text = tbRuc.Text.Length > 0 ? tbRuc.Text.Substring(0, tbRuc.Text.Length < 10 ? tbRuc.Text.Length : 10) : "";
            }
            else if (type == 3)
            {
                //if select CONSUMIDOR FINAL
                tbRuc.Text = "9999999999999";
            }
        }

        private void btnFormulario_Click(object sender, EventArgs e)
        {
            hideSubMenu(pnFormulario);
        }

        private void hideSubMenu(Panel panel)
        {
            if (pnAnterior != null)
            {
                pnAnterior.Visible = false;
            }
            if (panel != pnAnterior)
            {
                panel.Visible = !panel.Visible;
                pnAnterior = panel;
            }
        }

        private void btnEditar_Click(object sender, EventArgs e)
        {
            hideSubMenu(pnEditar);
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            hideSubMenu(pnVista);
        }

        private void bntImportFromMenu_Click(object sender, EventArgs e)
        {
            hideSubMenu(pnImportar);
        }

        private void btnDescargar_Click(object sender, EventArgs e)
        {
            dgv.descarga(pgbImport);
        }

        private void dgvVentas_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                double sumNumbers = 0;

                for (int i = 0; i < dgvVentas.SelectedCells.Count; i++)
                {
                    int column = dgvVentas.CurrentCell.ColumnIndex;
                    double nextNumber = 0;
                    if (!dgvVentas.SelectedCells.Contains(dgvVentas.Rows[i].Cells[i]) &&
                        (column > 5 && column < 16) &&
                        double.TryParse(dgvVentas.SelectedCells[i].FormattedValue.ToString(), out nextNumber))
                    {
                        sumNumbers += nextNumber;
                    }
                }

                lbSum.Text = "Suma " + sumNumbers;
                lbRows.Text = dgvVentas.SelectedCells.Count.ToString() + " de " + dgvVentas.Rows.Count + " registros";
            }
            catch (Exception) { }
        }

        //protected override CreateParams CreateParams
        //{
        //    get
        //    {
        //        CreateParams cp = base.CreateParams;
        //        cp.ExStyle |= 0x02000000;  // Turn on WS_EX_COMPOSITED
        //        return cp;
        //    }
        //}

        private void btnActualizar_Click(object sender, EventArgs e)
        {
            sumcolumns();
        }

        private void sumcolumns()
        {
            double b0 = 0, b12 = 0, iva = 0, vrr = 0, vri = 0;

            try
            {
                for (int i = 0; i < dgvVentas.Rows.Count; i++)
                {
                    b0 += double.Parse(dgvVentas.Rows[i].Cells[7].Value.ToString()) * (dgvVentas.Rows[i].Cells[5].Value.ToString().Equals("N/C") ? -1 : 1);
                    b12 += double.Parse(dgvVentas.Rows[i].Cells[8].Value.ToString()) * (dgvVentas.Rows[i].Cells[5].Value.ToString().Equals("N/C") ? -1 : 1);
                    iva += double.Parse(dgvVentas.Rows[i].Cells[10].Value.ToString()) * (dgvVentas.Rows[i].Cells[5].Value.ToString().Equals("N/C") ? -1 : 1);
                    vrr += double.Parse(dgvVentas.Rows[i].Cells[14].Value.ToString());
                    vri += double.Parse(dgvVentas.Rows[i].Cells[15].Value.ToString());
                }

                lbB0.Text = "b0: " + Math.Round(b0, 2);
                lbB12.Text = "b12: " + Math.Round(b12, 2);
                lbIva.Text = "IVA: " + Math.Round(iva, 2);
                lbVrr.Text = "Vrr: " + Math.Round(vrr, 2);
                lbVri.Text = "Vri: " + Math.Round(vri, 2);
            }
            catch (Exception) { }
        }

        private void cbTpId_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right)
            {
                tbRuc.Focus();
            }
        }

        private void dgvVentas_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if (e.Column.Index > 5 && e.Column.Index < 16)
            {
                float value1 = e.CellValue1 != null ? float.Parse(e.CellValue1.ToString()) : 0;
                float value2 = e.CellValue2 != null ? float.Parse(e.CellValue2.ToString()) : 0;

                e.SortResult = value1.CompareTo(value2);
                e.Handled = true;
            }
            else
            {
                string value1 = e.CellValue1 != null ? e.CellValue1.ToString() : string.Empty;
                string value2 = e.CellValue2 != null ? e.CellValue2.ToString() : string.Empty;

                e.SortResult = String.Compare(value1, value2);
                e.Handled = true;
            }
        }

        private void tbRazonS_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Control | Keys.V))
                (sender as TextBox).Paste();
        }

        private void btnXmlToPdf_Click(object sender, EventArgs e)
        {
            Importa select = new Importa();
            string path = select.selectFolder();

            if (path != null)
            {
                string[] files = Directory.GetFiles(path);
                int i = 0;

                pgbImport.Value = 0;
                pgbImport.Visible = true;

                foreach (var file in files)
                {
                    if (file.EndsWith(".xml") || file.EndsWith(".XML"))
                    {
                        ReadXml read = new ReadXml();
                        read.conversion(file, path);
                        pgbImport.Value = (i * 100) / files.Length;
                    }
                    i++;
                }

                pgbImport.Visible = false;
                pgbImport.Value = 100;
            }
        }
    }
}