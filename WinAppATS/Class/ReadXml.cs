using Microsoft.Reporting.WinForms;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace WinAppATS.Class
{
    class ReadXml
    {

        char dec = ',';
        public void conversion(string file, string path)
        {
            try
            {
                XElement XTemp = XElement.Load(file);

                var queryCDATAXML = from element in XTemp.DescendantNodes()
                                    where element.NodeType == System.Xml.XmlNodeType.CDATA
                                    select element.Parent.Value.Trim();

                string BodyHtml = queryCDATAXML.ToList<string>()[0].ToString();
                XDocument xmlDoc = XDocument.Parse(BodyHtml);

                int codDoc = int.Parse(xmlDoc.Descendants("codDoc").FirstOrDefault().Value.Trim());
                int lentparamenters = 0;

                switch (codDoc)
                {
                    case 1:
                        lentparamenters = 23;
                        break;
                    case 7:
                        lentparamenters = 15;
                        break;
                }

                ReportParameter[] parameters = new ReportParameter[lentparamenters];

                parameters[0] = new ReportParameter("autorization", XTemp.Descendants("numeroAutorizacion").FirstOrDefault().Value);
                parameters[1] = new ReportParameter("data_time_autorization", XTemp.Descendants("fechaAutorizacion").FirstOrDefault().Value);
                parameters[2] = new ReportParameter("environment", int.Parse(xmlDoc.Descendants("ambiente").FirstOrDefault().Value.Trim()) == 1 ? "PRUEBAS" : "PRODUCCION");

                parameters[3] = new ReportParameter("company", xmlDoc.Descendants("razonSocial").FirstOrDefault().Value.Trim());
                parameters[4] = new ReportParameter("name", xmlDoc.Descendants("nombreComercial").Any() ? xmlDoc.Descendants("nombreComercial").FirstOrDefault().Value.Trim() : null);
                parameters[5] = new ReportParameter("parenta_ddress", xmlDoc.Descendants("dirMatriz").FirstOrDefault().Value);
                parameters[6] = new ReportParameter("branch_address", xmlDoc.Descendants("dirEstablecimiento").Any() ? xmlDoc.Descendants("dirEstablecimiento").FirstOrDefault().Value : null);
                parameters[7] = new ReportParameter("special", xmlDoc.Descendants("contribuyenteEspecial").Any() ? xmlDoc.Descendants("contribuyenteEspecial").FirstOrDefault().Value : null);
                parameters[8] = new ReportParameter("accounting", xmlDoc.Descendants("obligadoContabilidad").Any() ? xmlDoc.Descendants("obligadoContabilidad").FirstOrDefault().Value : null);
                parameters[9] = new ReportParameter("retention_agent", xmlDoc.Descendants("agenteRetencion").Any() ? xmlDoc.Descendants("agenteRetencion").FirstOrDefault().Value : null);
                parameters[10] = new ReportParameter("micro_business", xmlDoc.Descendants("regimenMicroempresas").Any() ? xmlDoc.Descendants("regimenMicroempresas").FirstOrDefault().Value : null);

                DataTable tableAditional = new DataTable();

                tableAditional.Columns.Add("nombre", typeof(string));
                tableAditional.Columns.Add("valor", typeof(string));

                foreach (var detalle in xmlDoc.Descendants("campoAdicional"))
                {
                    DataRow row = tableAditional.NewRow();
                    row["nombre"] = detalle.Attribute("nombre").Value.Trim();
                    row["valor"] = detalle.Value.Trim();
                    tableAditional.Rows.Add(row);
                }

                switch (codDoc)
                {
                    case 1:
                        invoice(parameters, xmlDoc, path, tableAditional);
                        break;
                    case 7:
                        if (xmlDoc.Root.Attribute("version").Value.Trim().Equals("2.0.0"))
                        {
                            retentionv2(parameters, xmlDoc, path, tableAditional);
                        }
                        else
                        {
                            retentionv1(parameters, xmlDoc, path, tableAditional);
                        }
                        break;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("La carpeta solo debe contener comprobantes electrónicos", "Error al convertir XML a PDF", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void invoice(ReportParameter[] parameters, XDocument xmlDoc, string path, DataTable tableAditional)
        {
            string fechaEmision = xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value;
            parameters[11] = new ReportParameter("name_c", xmlDoc.Descendants("razonSocialComprador").FirstOrDefault().Value);
            parameters[12] = new ReportParameter("identification_c", xmlDoc.Descendants("identificacionComprador").FirstOrDefault().Value);
            parameters[13] = new ReportParameter("address_c", xmlDoc.Descendants("direccionComprador").Any() ? xmlDoc.Descendants("direccionComprador").FirstOrDefault().Value : null);
            parameters[14] = new ReportParameter("emision_date", fechaEmision);

            double no_iva = 0;
            double base0 = 0;
            double base12 = 0;
            double ice = 0;

            foreach (var totalImpuesto in xmlDoc.Descendants("totalImpuesto"))
            {
                switch (Int32.Parse(totalImpuesto.Descendants("codigoPorcentaje").FirstOrDefault().Value))
                {
                    case 0: base0 = Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    case 2: base12 = Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    case 3: base12 = Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    case 6: no_iva = Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    default:
                        if (Int32.Parse(totalImpuesto.Descendants("codigo").FirstOrDefault().Value) == 3)
                        {
                            ice = Math.Round(double.Parse(totalImpuesto.Descendants("valor").FirstOrDefault().Value.Replace('.', dec)), 2);
                        }
                        break;
                }
            }

            double sin_impuesto = double.Parse(xmlDoc.Descendants("totalSinImpuestos").FirstOrDefault().Value.Replace('.', dec));
            double iva = base12 * Const.IVA;
            double descuento = double.Parse(xmlDoc.Descendants("totalDescuento").FirstOrDefault().Value.Replace('.', dec));
            double total = double.Parse(xmlDoc.Descendants("importeTotal").FirstOrDefault().Value.Replace('.', dec));
            double propina = double.Parse(xmlDoc.Descendants("propina").Any() ? xmlDoc.Descendants("propina").FirstOrDefault().Value.Replace('.', dec) : "0");

            parameters[15] = new ReportParameter("base12", base12.ToString("N2"));
            parameters[16] = new ReportParameter("base0", base0.ToString("N2"));
            parameters[17] = new ReportParameter("no_iva", no_iva.ToString("N2"));
            parameters[18] = new ReportParameter("iva", iva.ToString("N2"));
            parameters[19] = new ReportParameter("sin_impuesto", sin_impuesto.ToString("N2"));
            parameters[20] = new ReportParameter("descuento", descuento.ToString("N2"));
            parameters[21] = new ReportParameter("total", total.ToString("N2"));
            parameters[22] = new ReportParameter("propina", propina.ToString("N2"));

            DataTable table = new DataTable();

            table.Columns.Add("codigoPrincipal", typeof(string));
            table.Columns.Add("descripcion", typeof(string));
            table.Columns.Add("cantidad", typeof(string));
            table.Columns.Add("precioUnitario", typeof(string));
            table.Columns.Add("descuento", typeof(string));
            table.Columns.Add("precioTotalSinImpuesto", typeof(string));

            foreach (var detalle in xmlDoc.Descendants("detalle"))
            {
                DataRow row = table.NewRow();
                row["codigoPrincipal"] = detalle.Descendants("codigoPrincipal").Any() ? detalle.Descendants("codigoPrincipal").FirstOrDefault().Value : "";
                row["descripcion"] = detalle.Descendants("descripcion").Any() ? detalle.Descendants("descripcion").FirstOrDefault().Value : "";
                row["cantidad"] = detalle.Descendants("cantidad").Any() ? detalle.Descendants("cantidad").FirstOrDefault().Value : "";
                row["precioUnitario"] = detalle.Descendants("precioUnitario").Any() ? detalle.Descendants("precioUnitario").FirstOrDefault().Value : "";
                row["descuento"] = detalle.Descendants("descuento").Any() ? detalle.Descendants("descuento").FirstOrDefault().Value : "";
                row["precioTotalSinImpuesto"] = detalle.Descendants("precioTotalSinImpuesto").Any() ? detalle.Descendants("precioTotalSinImpuesto").FirstOrDefault().Value : "";
                table.Rows.Add(row);
            }

            string razonSocial = xmlDoc.Descendants("razonSocial").FirstOrDefault().Value.Trim();
            string ruc = xmlDoc.Descendants("ruc").FirstOrDefault().Value.Trim();
            string secuencial = xmlDoc.Descendants("secuencial").FirstOrDefault().Value.Trim();
            string[] vect = razonSocial.Split(' ');

            fechaEmision = fechaEmision.Replace('/', '_');
            fechaEmision = fechaEmision.Replace("\"", "_");

            path = path + "/" + vect[0] + '_' + fechaEmision + '_' + secuencial + '_' + ruc + ".pdf";

            ReportDataSource report = new ReportDataSource("dsItems", table);

            ReportViewer viewer = new ReportViewer();
            viewer.LocalReport.ReportPath = Const.filereport("Invoice");
            //viewer.LocalReport.ReportPath = "../../Reports/Invoice.rdlc";
            viewer.LocalReport.DataSources.Clear();
            viewer.LocalReport.DataSources.Add(report);
            viewer.LocalReport.DataSources.Add(new ReportDataSource("dsDetallAddional", tableAditional));
            viewer.LocalReport.SetParameters(parameters);
            SavePDF(viewer, path);
        }

        private void retentionv1(ReportParameter[] parameters, XDocument xmlDoc, string path, DataTable tableAditional)
        {
            string fechaEmision = xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value;

            parameters[11] = new ReportParameter("name_c", xmlDoc.Descendants("razonSocialSujetoRetenido").FirstOrDefault().Value);
            parameters[12] = new ReportParameter("identification_c", xmlDoc.Descendants("identificacionSujetoRetenido").FirstOrDefault().Value);
            parameters[13] = new ReportParameter("address_c", "");
            parameters[14] = new ReportParameter("emision_date", fechaEmision);
            DataTable table = new DataTable();

            table.Columns.Add("comprobante", typeof(string));
            table.Columns.Add("numero", typeof(string));
            table.Columns.Add("comprobantefechaemision", typeof(string));
            table.Columns.Add("ejerciciofiscal", typeof(string));
            table.Columns.Add("baseImponible", typeof(string));
            table.Columns.Add("impuesto", typeof(string));
            table.Columns.Add("porcentaje", typeof(string));
            table.Columns.Add("valor", typeof(string));

            foreach (var detalle in xmlDoc.Descendants("impuesto"))
            {
                DataRow row = table.NewRow();
                row["comprobante"] = tipoComprobante(detalle.Descendants("codDocSustento").FirstOrDefault().Value);
                row["numero"] = detalle.Descendants("numDocSustento").Any() ? detalle.Descendants("numDocSustento").FirstOrDefault().Value : "null";
                row["comprobantefechaemision"] = detalle.Descendants("fechaEmisionDocSustento").Any() ? detalle.Descendants("fechaEmisionDocSustento").FirstOrDefault().Value : "null";
                row["ejerciciofiscal"] = xmlDoc.Descendants("periodoFiscal").Any() ? xmlDoc.Descendants("periodoFiscal").FirstOrDefault().Value : "";
                row["baseImponible"] = detalle.Descendants("baseImponible").Any() ? detalle.Descendants("baseImponible").FirstOrDefault().Value : "";
                row["impuesto"] = int.Parse(detalle.Descendants("codigo").FirstOrDefault().Value) == 1 ? "Renta" : "IVA";
                row["porcentaje"] = detalle.Descendants("porcentajeRetener").Any() ? detalle.Descendants("porcentajeRetener").FirstOrDefault().Value : "";
                row["valor"] = detalle.Descendants("valorRetenido").Any() ? detalle.Descendants("valorRetenido").FirstOrDefault().Value : "";
                table.Rows.Add(row);
            }

            string razonSocial = xmlDoc.Descendants("razonSocial").FirstOrDefault().Value.Trim();
            string ruc = xmlDoc.Descendants("ruc").FirstOrDefault().Value.Trim();
            string secuencial = xmlDoc.Descendants("secuencial").FirstOrDefault().Value.Trim();
            string[] vect = razonSocial.Split(' ');

            fechaEmision = fechaEmision.Replace('/', '_');
            fechaEmision = fechaEmision.Replace("\"", "_");

            path = path + "/" + vect[0] + '_' + fechaEmision + '_' + secuencial + '_' + ruc + ".pdf";

            ReportDataSource report = new ReportDataSource("dsRetentionItemsV1", table);

            ReportViewer viewer = new ReportViewer();
            viewer.LocalReport.ReportPath = Const.filereport("RetentionV1");
            //viewer.LocalReport.ReportPath = "../../Reports/RetentionV1.rdlc";
            viewer.LocalReport.DataSources.Clear();
            viewer.LocalReport.DataSources.Add(report);
            viewer.LocalReport.DataSources.Add(new ReportDataSource("dsDetallAddditional", tableAditional));
            viewer.LocalReport.SetParameters(parameters);
            SavePDF(viewer, path);
        }

        private void retentionv2(ReportParameter[] parameters, XDocument xmlDoc, string path, DataTable tableAditional)
        {
            string fechaEmision = xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value;

            parameters[11] = new ReportParameter("name_c", xmlDoc.Descendants("razonSocialSujetoRetenido").FirstOrDefault().Value);
            parameters[12] = new ReportParameter("identification_c", xmlDoc.Descendants("identificacionSujetoRetenido").FirstOrDefault().Value);
            parameters[13] = new ReportParameter("address_c", "");
            parameters[14] = new ReportParameter("emision_date", fechaEmision);

            DataTable table = new DataTable();

            table.Columns.Add("comprobante", typeof(string));
            table.Columns.Add("numero", typeof(string));
            table.Columns.Add("fechaemision", typeof(string));
            table.Columns.Add("ejerciciofiscal", typeof(string));
            table.Columns.Add("baseImponible", typeof(string));
            table.Columns.Add("impuesto", typeof(string));
            table.Columns.Add("porcentaje", typeof(string));
            table.Columns.Add("valor", typeof(string));

            foreach (var detalle in xmlDoc.Descendants("retencion"))
            {
                DataRow row = table.NewRow();
                row["comprobante"] = tipoComprobante(xmlDoc.Descendants("codDocSustento").FirstOrDefault().Value);
                row["numero"] = xmlDoc.Descendants("numDocSustento").Any() ? xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value : "null";
                row["fechaemision"] = xmlDoc.Descendants("fechaEmisionDocSustento").Any() ? xmlDoc.Descendants("fechaEmisionDocSustento").FirstOrDefault().Value : "null";
                row["ejerciciofiscal"] = xmlDoc.Descendants("periodoFiscal").Any() ? xmlDoc.Descendants("periodoFiscal").FirstOrDefault().Value : "";
                row["baseImponible"] = detalle.Descendants("baseImponible").Any() ? detalle.Descendants("baseImponible").FirstOrDefault().Value : "";
                row["impuesto"] = int.Parse(detalle.Descendants("codigo").FirstOrDefault().Value) == 1 ? "Renta" : "IVA";
                row["porcentaje"] = detalle.Descendants("porcentajeRetener").Any() ? detalle.Descendants("porcentajeRetener").FirstOrDefault().Value : "";
                row["valor"] = detalle.Descendants("valorRetenido").Any() ? detalle.Descendants("valorRetenido").FirstOrDefault().Value : "";
                table.Rows.Add(row);
            }

            string razonSocial = xmlDoc.Descendants("razonSocial").FirstOrDefault().Value.Trim();
            string ruc = xmlDoc.Descendants("ruc").FirstOrDefault().Value.Trim();
            string secuencial = xmlDoc.Descendants("secuencial").FirstOrDefault().Value.Trim();
            string[] vect = razonSocial.Split(' ');

            fechaEmision = fechaEmision.Replace('/', '_');
            fechaEmision = fechaEmision.Replace("\"", "_");

            path = path + "/" + vect[0] + '_' + fechaEmision + '_' + secuencial + '_' + ruc + ".pdf";

            ReportDataSource report = new ReportDataSource("dsRetentionItems", table);

            ReportViewer viewer = new ReportViewer();
            viewer.LocalReport.ReportPath = Const.filereport("RetentionV2");
            //viewer.LocalReport.ReportPath = "../../Reports/RetentionV2.rdlc";
            viewer.LocalReport.DataSources.Clear();
            viewer.LocalReport.DataSources.Add(report);
            viewer.LocalReport.DataSources.Add(new ReportDataSource("dsDetallAddional", tableAditional));
            viewer.LocalReport.SetParameters(parameters);
            SavePDF(viewer, path);
        }

        private string tipoComprobante(string type)
        {
            string result = "OTROS";

            switch (int.Parse(type))
            {
                case 1: result = "FACTURA"; break;
                case 7: result = "COMP RETENCION"; break;
                case 12: result = "DOCUMENTOS IFIS"; break;
            }

            return result;
        }

        private void SavePDF(ReportViewer viewer, string savePath)
        {
            //SqlServerTypes.Utilities.LoadNativeAssemblies(AppDomain.CurrentDomain.BaseDirectory);

            byte[] Bytes = viewer.LocalReport.Render(format: "PDF", deviceInfo: "");

            using (FileStream stream = new FileStream(savePath, FileMode.Create))
            {
                stream.Write(Bytes, 0, Bytes.Length);
            }
        }

        ~ReadXml() { }
    }
}
