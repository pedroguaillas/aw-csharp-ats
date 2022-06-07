using Microsoft.Reporting.WinForms;
using System;
using System.Data;
using System.Windows.Forms;

namespace WinAppATS
{
    public partial class FormRComprasProveedor : Form
    {
        public FormRComprasProveedor()
        {
            InitializeComponent();
        }

        private void FormRComprasProveedor_Load(object sender, EventArgs e)
        {
            string month = this.Name.Substring(this.Name.Length - 2, 2);
            string year = this.Name.Substring(this.Name.Length - 6, 4);

            DateTime firstDayOfMonth = new DateTime(int.Parse(year), int.Parse(month), 1);
            dtpStart.Value = firstDayOfMonth;

            var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
            dtpEnd.Value = lastDayOfMonth;

            reporte();
        }

        private void BtnActualizar_Click(object sender, EventArgs e)
        {
            if (dtpStart.Value > dtpEnd.Value)
            {
                MessageBox.Show("La fecha inicio debe ser menor a la fecha fin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dtpStart.Value.Year != dtpEnd.Value.Year)
            {
                MessageBox.Show("El reporte acumulado debe ser del mismo año", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            this.pnCotainer.Controls.Remove(this.reportViewer1);
            reporte();
        }

        private async void reporte()
        {
            ReportDate reportDate = new ReportDate(dtpStart.Value, dtpEnd.Value, Name.Substring(Name.Length - 19));
            DataSet data = await reportDate.compras();

            if (data.Tables.Count > 0)
            {
                Cliente cliente = new Cliente();
                ReportParameter[] parameters = new ReportParameter[1];
                parameters[0] = new ReportParameter("rs", cliente.getCliente(this.Name.Substring(this.Name.Length - 19, 13)));

                DataTable dataTable = data.Tables[0];
                if (Const.dec == '.')
                {
                    ConvertionDecimal convertion = new ConvertionDecimal();
                    convertion.convertion(dataTable);
                }

                ReportDataSource report = new ReportDataSource("dsCompras", dataTable);
                this.reportViewer1.ServerReport.BearerToken = null;
                this.reportViewer1.TabIndex = 0;
                this.reportViewer1.LocalReport.ReportPath = Const.filereport("RComprasProveedor");
                this.reportViewer1.LocalReport.DataSources.Clear();
                this.reportViewer1.LocalReport.DataSources.Add(report);
                this.reportViewer1.LocalReport.SetParameters(parameters);
                this.reportViewer1.RefreshReport();
                this.pnCotainer.Controls.Add(this.reportViewer1);
            }
            else
            {
                MessageBox.Show("No se ha registrado compras", "Reporte de compras", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
