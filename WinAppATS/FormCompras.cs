using System;
using System.Data;
using System.Windows.Forms;
using System.Globalization;
using System.Drawing;
using System.IO;
using WinAppATS.Class;
using System.Net.Http;

namespace WinAppATS
{
    public partial class FormCompras : Form
    {
        DgvCompras dgv;
        NumberFormatInfo nfi;
        Panel pnAnterior;

        const double IVA = 0.12;

        public FormCompras()
        {
            InitializeComponent();
            dgv = new DgvCompras(dgvCompras);
            cbTpId.SelectedIndex = 0;
            nfi = NumberFormatInfo.CurrentInfo;
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            dgv.Importar(cbCalculable);
            sumcolumns();
        }

        private void cleanForm()
        {
            tbRuc.Text = "";
            cbTpId.SelectedIndex = -1;
            tbRazonS.Text = "";
            cbContabilidad.Checked = false;
            tbDia.Text = "";
            tbEstablecimiento.Text = "";
            tbPtoEmi.Text = "";
            tbSecuencial.Text = "";
            tbAutorizacion.Text = "";
            tbBaseNoIva.Text = "";
            tbBase0.Text = "";
            tbBase12.Text = "";
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            save();
        }

        private async void save()
        {
            try
            {
                if (dgvCompras.Rows.Count > 0)
                {
                    BtnSave.Enabled = false;
                    string rest = await dgv.toXml();
                    if (rest.Contains("OK"))
                    {
                        BtnSave.Enabled = true;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No hay conexión a internet");
            }
        }

        private void FormCompras_Load(object sender, EventArgs e)
        {
            dgv.Load("c" + this.Name.Substring(this.Name.Length - 19));
            sumcolumns();
            duplicate();
        }

        private void BtnRetenciones_Click(object sender, EventArgs e)
        {
            dgv.ocultarRet(true);
        }

        private void CbRetenciones_CheckedChanged(object sender, EventArgs e)
        {
            dgv.ocultarRet(cbRetenciones.Checked);
        }

        private void CbNotasC_CheckedChanged(object sender, EventArgs e)
        {
            dgv.ocultarNCredit(cbNotasC.Checked);
        }

        private void dgvCompras_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            int column = dgvCompras.CurrentCell.ColumnIndex;
            //if (titleText.Equals("Cod cuenta") || titleText.Equals("Cod ATS"))
            if (column == 3 || column == 28)
            {
                TextBox autoText = e.Control as TextBox;

                if (autoText != null)
                {
                    autoText.AutoCompleteMode = AutoCompleteMode.Suggest;
                    autoText.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    AutoCompleteStringCollection DataCollection = new AutoCompleteStringCollection();
                    addItems(DataCollection);
                    autoText.AutoCompleteCustomSource = DataCollection;
                    autoText.KeyUp += textBox_SelectedAutoComplete;
                }
            }
        }

        private void addItems(AutoCompleteStringCollection col)
        {
            int column = dgvCompras.CurrentCell.ColumnIndex;

            if (column == 3)
            {
                Count count = new Count();

                foreach (String val in count.getColecs())
                {
                    col.Add(val);
                }
            }
            else if (column == 28)
            {
                Retencion count = new Retencion();

                foreach (String val in count.getColecs())
                {
                    col.Add(val);
                }
            }
        }

        private void textBox_SelectedAutoComplete(object sender, KeyEventArgs e)
        {
            int column = dgvCompras.CurrentCell.ColumnIndex;
            if (column == 3 && (sender as TextBox).Text.Length > 8)
            {
                string selected = (sender as TextBox).Text;
                int row = dgvCompras.CurrentCell.RowIndex;
                cuentaLoadSearch(selected, row);
            }
            else if (column == 28 && (sender as TextBox).Text.Length > 2)
            {
                string selected = (sender as TextBox).Text;
                int row = dgvCompras.CurrentCell.RowIndex;
                retLoadSearch(selected, row);
            }
        }

        private void BntGenereCod_Click(object sender, EventArgs e)
        {
            dgv.generarCodigo();
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
        }

        private async void TbRuc_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Control | Keys.V))
                (sender as TextBox).Paste();

            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                int tpId = cbTpId.SelectedIndex;
                string idContacto = tbRuc.Text;
                ImportContactos import = new ImportContactos();
                Contacto contacto = null;
            
                switch (tpId)
                {
                    case 0:
                        ResultSriRuc result = null;
                        HttpClient client = new HttpClient();

                        HttpResponseMessage response = await client.GetAsync($"https://srienlinea.sri.gob.ec/sri-catastro-sujeto-servicio-internet/rest/Persona/obtenerPorTipoIdentificacion?numeroIdentificacion={idContacto}&tipoIdentificacion=R");
                        if (response.IsSuccessStatusCode)
                        {
                            result = await response.Content.ReadAsAsync<ResultSriRuc>();
                        }
                        if (result != null)
                        {
                            tbRazonS.Text = remplazar(result.nombreCompleto.Trim());
                            tbDia.Focus();
                        }
                        else{
                            MessageBox.Show("No es un RUC válido", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        break;

                    case 1:
                        Validate validate = new Validate();
                        if (validate.validateMessage(tpId, idContacto))
                        {
                            return;
                        }
                        contacto = await import.getContact(idContacto);
                        if (contacto != null)
                        {
                            tbRazonS.Text = contacto.denominacion.Trim();
                            tbDia.Focus();
                        }
                        else
                        {
                            MessageBox.Show("Registre el nombre del proveedor, no se registrado", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        break;

                    case 2:
                        contacto = await import.getContact(idContacto);
                        if (contacto != null)
                        {
                            tbRazonS.Text = contacto.denominacion.Trim();
                            tbDia.Focus();
                        }
                        else
                        {
                            MessageBox.Show("Registre el nombre del proveedor, no se registrado", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        break;
                }
            }
        }

        private string remplazar(string razonsocial)
        {
            razonsocial = razonsocial.Replace('Á', 'A');
            razonsocial = razonsocial.Replace('É', 'E');
            razonsocial = razonsocial.Replace('Í', 'I');
            razonsocial = razonsocial.Replace('Ó', 'O');
            razonsocial = razonsocial.Replace('Ú', 'U');
            razonsocial = razonsocial.Replace('Ñ', 'N');
            razonsocial = razonsocial.Replace(".", string.Empty);
            razonsocial = razonsocial.Replace(",", string.Empty);

            return razonsocial;
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

        private void TbEstablecimiento_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void TbEstablecimiento_KeyUp(object sender, KeyEventArgs e)
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
                tbAutorizacion.Focus();
            }
        }

        private void TbAutorizacion_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back);
        }

        private void TbAutorizacion_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                cbTpComprobante.Focus();
            }
        }

        private void cbTpComprobante_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                tbBaseNoIva.Focus();
            }
        }

        private void TbBaseNoIva_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericHandle(e);
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

        private void TbBaseNoIva_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Enter)
            {
                tbBase0.Focus();
            }
        }

        private void TbBase0_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericHandle(e);
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
                if (tbRazonS.Text == "")
                {
                    tbRazonS.Focus();
                }
                else
                {
                    btnAdd.Focus();
                }
            }
        }

        private void BtnAdd_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                addCompra();
            }
        }

        private void addCompra()
        {
            string tcv = "";
            switch (cbTpComprobante.SelectedIndex)
            {
                case 0: tcv = "F"; break;
                case 1: tcv = "N/C"; break;
                case 2: tcv = "N/D"; break;
                case 3: tcv = "L/C"; break;
                case 4: tcv = "N/V"; break;
            }
            try
            {
                dgv.addOne(tbRuc.Text, tbRazonS.Text, tcv, tbDia.Text, tbEstablecimiento.Text, tbPtoEmi.Text, tbSecuencial.Text, tbAutorizacion.Text, tbBaseNoIva.Text.ToString(nfi), tbBase0.Text.ToString(nfi), tbBase12.Text.ToString(nfi), "", "", "");
            }
            catch (Exception e)
            {
                MessageBox.Show("Se genero un error" + e.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            cleanInsert();

            if (cbRepetirProveedor.Checked)
            {
                tbDia.Focus();
            }
            else
            {
                cbTpId.Focus();
            }
            dgvCompras.FirstDisplayedScrollingRowIndex = dgvCompras.Rows.Count - 1;
            sumcolumns();
            return;
        }

        private void BntImportRetenciones_Click(object sender, EventArgs e)
        {
            DgvComprasImportRets dgvForRetenciones = new DgvComprasImportRets(dgvCompras);
            dgvForRetenciones.importRetenciones();
            sumcolumns();
        }

        private void btnInsertRowBefore_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvCompras.SelectedRows.Count > 0)
                {
                    DataGridViewRow row = dgvCompras.SelectedRows[0];
                    //.................Clonar la fila
                    DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                    for (Int32 index = 0; index < row.Cells.Count; index++)
                    {
                        clonedRow.Cells[index].Value = (index > 10 && index < 24) || index == 31 ? "0" : "";
                    }
                    dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No se ha seleccionado una fila");
            }
        }

        private void btnDuplicar_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvCompras.SelectedRows.Count > 0)
                {
                    DataGridViewRow row = dgvCompras.SelectedRows[0];
                    //.................Clonar la fila
                    DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                    for (Int32 index = 0; index < row.Cells.Count; index++)
                    {
                        clonedRow.Cells[index].Value = row.Cells[index].Value;
                    }
                    dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No se ha seleccionado una fila");
            }
        }

        private void cbBaseExeMontoIce_CheckedChanged(object sender, EventArgs e)
        {
            dgv.ocultarBExeMIce(cbBaseExeMontoIce.Checked);
        }

        private void btnClean_Click(object sender, EventArgs e)
        {
            cleanForm();
        }

        private void btnAddProv_Click(object sender, EventArgs e)
        {
            var obj = new { id = tbRuc.Text, denominacion = tbRazonS.Text, tpId = (cbTpId.SelectedIndex + 1).ToString("00"), contabilidad = cbContabilidad.Checked ? "SI" : "" };
            ImportContactos import = new ImportContactos();
            import.registrarUno(obj);
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            addCompra();
        }

        private void cleanInsert()
        {
            if (!cbRepetirProveedor.Checked)
            {
                cbTpId.SelectedIndex = 0;
                tbRuc.Text = "";
                tbRazonS.Text = "";
            }
            tbDia.Text = "";
            tbEstablecimiento.Text = "";
            tbPtoEmi.Text = "";
            tbSecuencial.Text = "";
            tbAutorizacion.Text = "";
            tbBaseNoIva.Text = "";
            tbBase0.Text = "";
            tbBase12.Text = "";
        }

        private void TbRazonS_KeyPress(object sender, KeyPressEventArgs e)
        {
            var regex = new System.Text.RegularExpressions.Regex(@"[^a-zA-Z\s]");
            if (e.KeyChar != (char)Keys.Back && regex.IsMatch(e.KeyChar.ToString()))
            {
                e.Handled = true;
            }
        }

        private void TbRazonS_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Control | Keys.V))
                (sender as TextBox).Paste();

            if (e.KeyCode == Keys.Right)
            {
                btnAddProv.Focus();
            }
        }

        private void BtnAddProv_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right)
            {
                btnAdd.Focus();
            }
        }

        private int searchRow = 0;

        private void btnSearch_Click(object sender, EventArgs e)
        {
            search();
        }

        private void DgvCompras_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Modifiers == Keys.Control)
                {
                    CopyPaste copyPaste = new CopyPaste(dgvCompras);
                    switch (e.KeyCode)
                    {
                        //case Keys.C:
                        //    copyPaste.CopyToClipboard();
                        //    break;
                        case Keys.V:
                            copyPaste.PasteClipboard("compras");
                            break;
                        case Keys.G:
                            save();
                            break;
                        case Keys.Z:
                            if (!dgvCompras.CurrentCell.IsInEditMode)
                            {
                                //dgvCompras.DropDownItems["Undo"].PerformClick();
                            }
                            break;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Fuera de rango ", "Copia y pega", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void search()
        {
            string searchValue = tbSearchValue.Text.Trim();
            int column = cbColumn.SelectedIndex;

            if (column == 2)
            {
                column = 3;
            }
            else if (column == 3)
            {
                column = 5;
            }
            else if (column == 4)
            {
                column = 6;
            }
            else if (column == 5)
            {
                column = 9;
            }

            try
            {
                DataGridViewRow row;
                for (; searchRow < dgvCompras.Rows.Count; searchRow++)
                {
                    row = dgvCompras.Rows[searchRow];
                    if (row.Cells[column].Value.ToString().Trim().Equals(searchValue) || (column == 9 && int.Parse(row.Cells[column].Value.ToString()) == int.Parse(searchValue)))
                    {
                        row.Selected = true;
                        dgvCompras.FirstDisplayedScrollingRowIndex = row.Index;
                        break;
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
            searchRow++;
            btnRepetSearch.Enabled = true;
        }

        private void BtnRepetSearch_Click(object sender, EventArgs e)
        {
            searchRow = 0;
        }

        private void BtnCompleteRetencion_Click(object sender, EventArgs e)
        {
            if (dgvCompras.Rows.Count > 0)
            {
                dgv.completeRete();
            }
        }

        private void cbColumn_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbColumn.SelectedIndex > -1)
            {
                tbSearchValue.Enabled = true;
                tbSearchValue.Text = "";
                searchRow = 0;
                btnRepetSearch.Enabled = false;
            }
        }

        private void DgvCompras_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if ((e.ColumnIndex > 10 && e.ColumnIndex < 24) || e.ColumnIndex == 30 || e.ColumnIndex == 31)
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
        }

        private void DgvCompras_CellStateChanged(object sender, DataGridViewCellStateChangedEventArgs e)
        {
            try
            {
                int column = e.Cell.ColumnIndex;

                if (column == 3)
                {
                    cuentaLoadSearch(e.Cell.Value == null ? "" : e.Cell.Value.ToString(), e.Cell.RowIndex);
                }
                else if (column == 12)
                {
                    calTotal(e.Cell.RowIndex);
                    calValRet(e.Cell.RowIndex);
                }
                else if (column == 13)
                {
                    calIva(e.Cell.RowIndex);
                    calTotal(e.Cell.RowIndex);
                    calValRet(e.Cell.RowIndex);
                }
                else if (column == 15)
                {
                    calIva(e.Cell.RowIndex);
                    calTotal(e.Cell.RowIndex);
                    calValRet(e.Cell.RowIndex);
                }
                else if (column == 16)
                {
                    calTotal(e.Cell.RowIndex);
                    calValRet(e.Cell.RowIndex);
                }
                else if (column == 28)
                {
                    retLoadSearch(e.Cell.Value != null ? e.Cell.Value.ToString() : "", e.Cell.RowIndex);
                }
                else if (column == 30)
                {
                    calValRet(e.Cell.RowIndex);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Digitar del acuerdo al campo correspondiente");
            }
        }

        private void cuentaLoadSearch(string cuenta, int row)
        {
            if (cuenta.Length > 8)
            {
                Count count = new Count();
                dgvCompras.Rows[row].Cells[4].Value = count.completeCounts(cuenta);
            }
            else
            {
                dgvCompras.Rows[row].Cells[4].Value = "";
            }
        }

        private void retLoadSearch(string cod, int row)
        {
            if (cod.Length > 2)
            {
                Retencion retencion = new Retencion();
                DataRow dataRow = retencion.completeRetenciones(cod);
                if (dataRow != null)
                {
                    dgvCompras.Rows[row].Cells[29].Value = dataRow[1];
                    double porcentaje = 0;
                    string por = dataRow[2].ToString();
                    por = por.Replace('.', ',');
                    //por = por.Replace(Const.dec, Const.nodec);
                    if (cbCalculable.Checked && double.TryParse(por, out porcentaje))
                    {
                        dgvCompras.Rows[row].Cells[30].Value = porcentaje;
                        calValRet(row);
                    }
                    //else
                    //{
                    //    dgvCompras.Rows[row].Cells[30].Value = 0;
                    //}
                }
            }
            else if (cod.Equals(""))
            {
                dgvCompras.Rows[row].Cells[29].Value = "";
                dgvCompras.Rows[row].Cells[30].Value = "";
                dgvCompras.Rows[row].Cells[31].Value = 0;
                calValRet(row);
            }
        }

        private double calBaseImp(int row)
        {
            return double.Parse(dgvCompras.Rows[row].Cells[12].Value.ToString()) + double.Parse(dgvCompras.Rows[row].Cells[13].Value.ToString());
        }

        private void calIva(int row)
        {
            if (cbCalculable.Checked)
            {
                dgvCompras.Rows[row].Cells[16].Value = Const.round((double.Parse(dgvCompras.Rows[row].Cells[13].Value.ToString())) * IVA);
                // dgvCompras.Rows[row].Cells[16].Value = Math.Round((double.Parse(dgvCompras.Rows[row].Cells[13].Value.ToString())) * IVA, 2);
            }
        }

        private void calTotal(int row)
        {
            if (cbCalculable.Checked)
            {
                dgvCompras.Rows[row].Cells[17].Value = calBaseImp(row) + double.Parse(dgvCompras.Rows[row].Cells[14].Value.ToString()) + double.Parse(dgvCompras.Rows[row].Cells[16].Value.ToString());
            }
        }

        private void calValRet(int row)
        {
            if (cbCalculable.Checked)
            {
                dgvCompras.Rows[row].Cells[31].Value = dgvCompras.Rows[row].Cells[30].Value == null || dgvCompras.Rows[row].Cells[30].Value.Equals("") ? 0 : Const.round((calBaseImp(row) - double.Parse(dgvCompras.Rows[row].Cells[15].Value.ToString())) * double.Parse(dgvCompras.Rows[row].Cells[30].Value.ToString()) / 100);
                //dgvCompras.Rows[row].Cells[31].Value = dgvCompras.Rows[row].Cells[30].Value == null || dgvCompras.Rows[row].Cells[30].Value.Equals("") ? 0 : Math.Round((calBaseImp(row) - double.Parse(dgvCompras.Rows[row].Cells[15].Value.ToString())) * double.Parse(dgvCompras.Rows[row].Cells[30].Value.ToString()) / 100, 2);
            }
        }

        private void tbSearchValue_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && cbColumn.SelectedIndex > -1)
            {
                search();
            }
        }

        private void btnImpRecuperado_Click(object sender, EventArgs e)
        {
            dgv.importReport(pgbImport);
            sumcolumns();
        }

        private void btnRecuperado_Click(object sender, EventArgs e)
        {
            //dgv.importRecuperado();
            dgv.importRecuperadoMasiva();
            dgv.rellenarProveedores();
            dgv.rellenarProveedoresMasivo();
            sumcolumns();
        }

        private void cbTpId_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if select cedula
            if ((int)cbTpId.SelectedIndex == 1)
            {
                tbRuc.MaxLength = 10;
                tbRuc.Text = tbRuc.Text.Length > 0 ? tbRuc.Text.Substring(0, tbRuc.Text.Length < 10 ? tbRuc.Text.Length : 10) : "";
            }
            else
            {
                tbRuc.MaxLength = 13;
            }
            //En caso de seleccionar RUC o Cedula
            if ((int)cbTpId.SelectedIndex > 0)
            {
                //Tipo de comprobante por defecto "Liquidacion en Compra"
                cbTpComprobante.SelectedIndex = 3;
            }
        }

        private void btnFormulario_Click(object sender, EventArgs e)
        {
            hideSubMenu(pnFormulario);
        }

        private void btnEditar_Click(object sender, EventArgs e)
        {
            hideSubMenu(pnEditar);
        }

        private void bntImportFromMenu_Click(object sender, EventArgs e)
        {
            hideSubMenu(pnImportar);
        }

        private void hideSubMenu(Panel panel)
        {
            if (pnAnterior != null)
            {
                pnAnterior.Visible = false;
            }
            panel.Visible = !panel.Visible;
            pnAnterior = panel;
        }

        private void btnVer_Click(object sender, EventArgs e)
        {
            hideSubMenu(pnVista);
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

        private void btnDescargar_Click(object sender, EventArgs e)
        {
            dgv.descargar(this.Name.Substring(this.Name.Length - 19), pgbImport);
        }

        private void dgvCompras_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                double sumNumbers = 0;

                for (int i = 0; i < dgvCompras.SelectedCells.Count; i++)
                {
                    int column = dgvCompras.CurrentCell.ColumnIndex;
                    double nextNumber = 0;

                    if (!dgvCompras.SelectedCells.Contains(dgvCompras.Rows[i].Cells[i]) &&
                        ((column > 10 && column < 24) || column == 31) &&
                        double.TryParse(dgvCompras.SelectedCells[i].FormattedValue.ToString(), out nextNumber))
                    {
                        sumNumbers += nextNumber;
                    }
                }

                lbSum.Text = "Suma " + sumNumbers;
                lbRows.Text = dgvCompras.SelectedCells.Count.ToString() + " de " + dgvCompras.Rows.Count + " registros";
            }
            catch (Exception) { }
        }

        private void btnActualizar_Click(object sender, EventArgs e)
        {
            sumcolumns();
            duplicate();
        }

        private void sumcolumns()
        {
            double bNoIva = 0, b0 = 0, b12 = 0, iva = 0, ret = 0;

            for (int i = 0; i < dgvCompras.Rows.Count; i++)
            {
                bNoIva += double.Parse(dgvCompras.Rows[i].Cells[11].Value.ToString()) * (dgvCompras.Rows[i].Cells[5].Value.ToString().Trim().Equals("N/C") ? -1 : 1);
                b0 += double.Parse(dgvCompras.Rows[i].Cells[12].Value.ToString()) * (dgvCompras.Rows[i].Cells[5].Value.ToString().Trim().Equals("N/C") ? -1 : 1);
                b12 += double.Parse(dgvCompras.Rows[i].Cells[13].Value.ToString()) * (dgvCompras.Rows[i].Cells[5].Value.ToString().Trim().Equals("N/C") ? -1 : 1);
                iva += double.Parse(dgvCompras.Rows[i].Cells[16].Value.ToString()) * (dgvCompras.Rows[i].Cells[5].Value.ToString().Trim().Equals("N/C") ? -1 : 1);
                ret += double.Parse(dgvCompras.Rows[i].Cells[31].Value.ToString());
                //dgvCompras.Rows[i].Cells[5].Style.BackColor = dgvCompras.Rows[i].Cells[5].Value.ToString().Equals("N/C") ? Color.Red : Color.Yellow;
            }

            lbBNoIva.Text = "bNoIva: " + Math.Round(bNoIva, 2);
            lbB0.Text = "b0: " + Math.Round(b0, 2);
            lbB12.Text = "b12: " + Math.Round(b12, 2);
            lbIva.Text = "IVA: " + Math.Round(iva, 2);
            lbRetencion.Text = "Retención: " + Math.Round(ret, 2);
        }

        private void duplicate()
        {
            for (int i = 0; i < dgvCompras.Rows.Count - 1; i++)
            {
                for (int j = i + 1; j < dgvCompras.Rows.Count; j++)
                {
                    if (dgvCompras.Rows[i].Cells[9].Value == dgvCompras.Rows[j].Cells[9].Value)
                    {
                        dgvCompras.Rows[i].Cells[9].Style.BackColor = Color.Red;
                        dgvCompras.Rows[j].Cells[9].Style.BackColor = Color.Red;
                    }
                }
            }
        }

        private void tbRuc_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void cbTpId_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right)
            {
                tbRuc.Focus();
            }
        }

        private void dgvCompras_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            if ((e.Column.Index > 10 && e.Column.Index < 24) || e.Column.Index == 31)
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