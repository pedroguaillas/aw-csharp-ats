using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace WinAppATS
{
    class DgvVentas : Importa
    {
        DataGridView dgvV;
        String info;
        char dec;

        public DgvVentas(DataGridView dgv1)
        {
            this.dgvV = dgv1;
            dec = ',';
        }

        public void Load(string info)
        {
            this.info = info;
            if (File.Exists(Const.filexml(info)))
            {
                using (XmlReader xml = XmlReader.Create(Const.filexml(info)))
                {
                    DataSet data = new DataSet("ventas");
                    data.ReadXml(xml);

                    DataTable dt = data.Tables[0];

                    foreach (DataRow dr in dt.Rows)
                    {
                        dgvV.Rows.Add(dr.ItemArray);
                    }
                }
            }
        }

        public async Task<string> toXml()
        {
            if (dgvV.Rows.Count > 0)
            {
                DataTable dt = new DataTable("venta");
                for (int i = 1; i < dgvV.Columns.Count + 1; i++)
                {
                    string column = dgvV.Columns[i - 1].Name;
                    dt.Columns.Add(column, typeof(string));
                }

                int ColumnCount = dgvV.Columns.Count;
                foreach (DataGridViewRow dr in dgvV.Rows)
                {
                    DataRow dataRow = dt.NewRow();
                    for (int i = 0; i < ColumnCount; i++)
                    {
                        dataRow[i] = dr.Cells[i].Value;
                    }
                    dt.Rows.Add(dataRow);
                }

                DataSet ds = new DataSet("ventas");
                ds.Tables.Add(dt);
                ds.WriteXml(Const.filexml(info));

                XmlDocument doc = new XmlDocument();
                doc.Load(Const.filexml(info));

                using (XmlTextWriter wr = new XmlTextWriter(Const.filexml(info), Encoding.UTF8))
                {
                    wr.Formatting = System.Xml.Formatting.None; // here's the trick !
                    doc.Save(wr);
                }

                return await Guardar();

            }
            return "0";
        }

        private async Task<string> Guardar()
        {
            string result;
            byte[] bytes = File.ReadAllBytes(Const.filexml(info));

            using (var client = new HttpClient())
            {
                var data = new { file = Convert.ToBase64String(bytes), info = info.Substring(1), tipo = "fileventa" };

                string json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync(Const.URL + "archivos", content);
                result = await response.Content.ReadAsStringAsync();
            }
            return result;
        }

        public void Importar()
        {
            try
            {
                string path = selectFolder();
                if (path != null)
                {
                    string[] files = Directory.GetFiles(path);

                    List<Object> insertados = new List<Object>();
                    List<String> inserts = new List<String>();

                    int i = 0;

                    foreach (var file in files)
                    {
                        if (file.EndsWith(".xml") || file.EndsWith(".XML"))
                        {
                            XElement XTemp = XElement.Load(file);

                            var queryCDATAXML = from element in XTemp.DescendantNodes()
                                                where element.NodeType == System.Xml.XmlNodeType.CDATA
                                                select element.Parent.Value.Trim();

                            string BodyHtml = queryCDATAXML.ToList<string>()[0].ToString();
                            XDocument xmlDoc = XDocument.Parse(BodyHtml);

                            string codDoc = xmlDoc.Descendants("codDoc").FirstOrDefault().Value;

                            if (codDoc == "01" || codDoc == "04")
                            {
                                double base0 = 0;
                                double base12 = 0;

                                foreach (var totalImpuesto in xmlDoc.Descendants("totalImpuesto"))
                                {
                                    switch (Int32.Parse(totalImpuesto.Descendants("codigoPorcentaje").FirstOrDefault().Value))
                                    {
                                        case 0: base0 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                                        case 2: base12 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                                        case 3: base12 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                                    }
                                }

                                string tpId = xmlDoc.Descendants("tipoIdentificacionComprador").FirstOrDefault().Value;

                                dgvV.Rows.Add(
                                    xmlDoc.Descendants("identificacionComprador").FirstOrDefault().Value,
                                    i == 0 || tpId.Equals("06") ? xmlDoc.Descendants("razonSocialComprador").FirstOrDefault().Value : "",
                                    tpId,
                                    xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value,
                                    codDoc == "01" ? "F" : "N/C",//Tipo comprobante venta
                                    xmlDoc.Descendants("estab").FirstOrDefault().Value + "-" + xmlDoc.Descendants("ptoEmi").FirstOrDefault().Value + "-" + xmlDoc.Descendants("secuencial").FirstOrDefault().Value,
                                    base0 + base12, base0, base12,
                                    0, //Monto ICE
                                    Math.Round(base12 * Const.IVA, 2),//IVA
                                    base12 + base0 + Math.Round(base12 * Const.IVA, 2),//Total
                                    0,
                                    0,
                                    0,
                                    0,
                                    ""
                                );

                                if (tpId == "06" && !inserts.Contains(xmlDoc.Descendants("identificacionComprador").FirstOrDefault().Value))
                                {
                                    var obj = new
                                    {
                                        id = xmlDoc.Descendants("identificacionComprador").FirstOrDefault().Value,
                                        denominacion = xmlDoc.Descendants("razonSocialComprador").FirstOrDefault().Value,
                                        tpId = "03",
                                        contabilidad = ""
                                    };

                                    insertados.Add(obj);
                                    inserts.Add(obj.id);
                                }
                            }
                        }
                        i++;
                    }

                    if (insertados.Count > 0)
                    {
                        ImportContactos import = new ImportContactos();
                        import.registrarList(insertados);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("La carpeta solo debe contener comprobantes electrónicos", "Error al importar comprobantes electrónicos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void descarga(ProgressBar bar)
        {
            Importa importar = new Importa();
            string folder = importar.selectFolder();

            if (folder != null)
            {
                string[] files = Directory.GetFiles(folder);
                string problems = "";

                bar.Visible = true;

                foreach (var file in files)
                {
                    if (file.EndsWith(".txt"))
                    {
                        List<string> clave_accs = myAutorizaciones(file);

                        int i = 0;
                        bar.Value = 0;

                        foreach (var code in clave_accs)
                        {
                            AutorizacionFacturas af = new AutorizacionFacturas();
                            if (!af.Descarga(code, @"" + folder + "/xml"))
                            {
                                if (!af.Descarga(code, @"" + folder + "/xml"))
                                {
                                    if (!af.Descarga(code, @"" + folder + "/xml"))
                                    {
                                        problems += "\n" + code;
                                    }
                                }
                            }
                            bar.Value = (i * 100) / clave_accs.Count;
                            i++;
                        }
                    }
                }

                if (problems.Length > 0)
                {
                    string path = @"" + folder + "/Error.txt";
                    Archivo archivo = new Archivo();
                    archivo.saveError(problems, path);
                }

                bar.Visible = false;
            }
        }

        private List<string> myAutorizaciones(string path)
        {
            List<String> cla_accs = new List<string>();
            try
            {
                string line = "";
                StreamReader file = new StreamReader(path);
                int i = 0;
                while ((line = file.ReadLine()) != null)
                {
                    if (i % 2 == 0 && i > 0)
                    {
                        string[] palabras = line.Split('\t');
                        cla_accs.Add(palabras[3]);
                    }
                    i++;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No se ha seleccionado un archivo correcto");
            }
            return cla_accs;
        }

        public void importReport(ProgressBar bar)
        {
            Importa importa = new Importa();
            string folder = importa.selectFolder();

            if (folder == null) { return; }

            string[] files = Directory.GetFiles(folder);
            string problems = "";

            bar.Visible = true;

            foreach (var file in files)
            {
                if (file.EndsWith(".txt"))
                {
                    List<String> clave_accs = myAutorizaciones(file);
                    int i = 0;
                    int lent = clave_accs.Count;

                    foreach (var c in clave_accs)
                    {
                        AutorizacionFacturas autorizacion = new AutorizacionFacturas();
                        string xml = autorizacion.LlamarSri(c);

                        if (xml != null)
                        {
                            XmlDocument docResponse = new XmlDocument();
                            docResponse.LoadXml(xml);

                            if (int.Parse(docResponse.GetElementsByTagName("numeroComprobantes")[0].InnerXml.Trim()) == 1)
                            {
                                string tagComprobante = docResponse.GetElementsByTagName("comprobante")[0].InnerXml.Trim();
                                tagComprobante = tagComprobante.Replace("&lt;", "<");
                                tagComprobante = tagComprobante.Replace("&gt;", ">");

                                XDocument xmlDoc = XDocument.Parse(tagComprobante);

                                string codDoc = xmlDoc.Descendants("codDoc").FirstOrDefault().Value;

                                if (codDoc == "01" || codDoc == "04" || codDoc == "05")
                                {
                                    double imponible = 0;
                                    double base0 = 0;
                                    double base12 = 0;
                                    double ice = 0;
                                    foreach (var totalImpuesto in xmlDoc.Descendants("totalImpuesto"))
                                    {
                                        switch (Int32.Parse(totalImpuesto.Descendants("codigoPorcentaje").FirstOrDefault().Value))
                                        {
                                            case 0: base0 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                                            case 2: base12 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                                            case 3: base12 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                                            case 6: imponible += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                                            default:
                                                if (Int32.Parse(totalImpuesto.Descendants("codigo").FirstOrDefault().Value) == 3)
                                                {
                                                    ice += Math.Round(double.Parse(totalImpuesto.Descendants("valor").FirstOrDefault().Value.Replace('.', dec)), 2);
                                                }
                                                break;
                                        }
                                    }

                                    double iva = Math.Round(base12 * .12, 2);
                                    string tpId = xmlDoc.Descendants("tipoIdentificacionComprador").FirstOrDefault().Value;

                                    dgvV.Rows.Add(
                                        xmlDoc.Descendants("identificacionComprador").FirstOrDefault().Value,
                                        i == 0 || tpId.Equals("06") ? xmlDoc.Descendants("razonSocialComprador").FirstOrDefault().Value : "",
                                        tpId,
                                        xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value,
                                        codDoc == "01" ? "F" : "N/C",//Tipo comprobante venta
                                        xmlDoc.Descendants("estab").FirstOrDefault().Value + "-" + xmlDoc.Descendants("ptoEmi").FirstOrDefault().Value + "-" + xmlDoc.Descendants("secuencial").FirstOrDefault().Value,
                                        base0 + base12, base0, base12,
                                        0, //Monto ICE
                                        Math.Round(base12 * Const.IVA, 2),//IVA
                                        base12 + base0 + Math.Round(base12 * Const.IVA, 2),//Total
                                        0,
                                        0,
                                        0,
                                        0,
                                        ""
                                    );
                                }
                            }
                            else
                            {
                                problems += "\n" + c;
                            }
                        }
                        else
                        {
                            problems += "\n" + c;
                        }
                        i++;
                        bar.Value = (i * 100) / lent;
                    }

                    bar.Value = 100;
                }
            }
            bar.Visible = false;

            if (problems.Length > 0)
            {
                string file1 = @"" + folder + "/Error.txt";
                Archivo archivo = new Archivo();
                archivo.saveError(problems, file1);
            }

            dgvV.FirstDisplayedScrollingRowIndex = dgvV.Rows.Count - 1;
        }

        private List<string> extraerAutorizaciones()
        {
            List<String> cla_accs = new List<string>();
            try
            {
                string path = selectFile("txt");
                if (path != null)
                {
                    StreamReader file = new StreamReader(path);
                    string line = "";
                    int i = 0;
                    while ((line = file.ReadLine()) != null)
                    {
                        if (i % 2 == 0 && i > 0)
                        {
                            string[] palabras = line.Split('\t');
                            cla_accs.Add(palabras[3]);
                        }
                        i++;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No se ha seleccionado un archivo correcto");
            }
            return cla_accs;
        }

        public void addOne(string ruc, string razons, string tcv, string dia, string establecimiento, string ptoEmi, string secuencial, string base0, string base12)
        {
            double b0 = base0 == "" ? 0 : double.Parse(base0);
            double b12 = base12 == "" ? 0 : double.Parse(base12);
            dgvV.Rows.Add(
                ruc,
                razons,
                ruc.Length == 10 ? "05" : "04",
                string.Format(dia.PadLeft(2, '0') + "/" + info.Substring(info.Length - 2) + "/" + info.Substring(info.Length - 6, 4)),
                tcv,//Tipo comprobante venta
                 string.Format(establecimiento.PadLeft(3, '0') + "-" + ptoEmi.PadLeft(3, '0') + "-" + secuencial.PadLeft(9, '0')),
                 b0 + b12,
                 b0, b12,
                0, //Monto ICE
                Math.Round(b12 * .12, 2),
                b0 + b12 + Math.Round(b12 * .12, 2),
                0,
                0,
                0,
                0,
                ""
            );
            dgvV.FirstDisplayedScrollingRowIndex = dgvV.Rows.Count - 1;
        }

        public void importRetenciones()
        {
            try
            {
                string path = selectFolder();
                if (path != null)
                {
                    string[] files = Directory.GetFiles(path);
                    List<string> othersRet = new List<string>();

                    foreach (var file in files)
                    {
                        if (file.EndsWith(".xml") || file.EndsWith(".XML"))
                        {
                            XElement XTemp = XElement.Load(file);

                            var queryCDATAXML = from element in XTemp.DescendantNodes()
                                                where element.NodeType == System.Xml.XmlNodeType.CDATA
                                                select element.Parent.Value.Trim();

                            string BodyHtml = queryCDATAXML.ToList<string>()[0].ToString();
                            XDocument xmlDoc = XDocument.Parse(BodyHtml);
                            bool import = false;

                            // Verificar si es retencion
                            if (xmlDoc.Descendants("codDoc").FirstOrDefault().Value.Trim().Equals("07"))
                            {
                                if (xmlDoc.Root.Attribute("version").Value.Trim().Equals("2.0.0"))
                                {
                                    import = importRetVersion2(xmlDoc);
                                }
                                else
                                {
                                    import = importRetVersion1(xmlDoc);
                                }
                            }

                            //Si no se importo las retenciones
                            if (!import)
                            {
                                othersRet.Add("Retención: " + xmlDoc.Descendants("secuencial").FirstOrDefault().Value + " fecha: " + xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value + " archivo: " + file);
                            }
                        }
                    }
                    if (othersRet.Count > 0)
                    {
                        string invoices = othersRet.Count + " retenciones no se han importado\n";
                        othersRet.ForEach(ret => invoices += ret + "\n");

                        MessageBox.Show(invoices, "Retenciones", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    if (dgvV.Rows.Count > 0)
                    {
                        dgvV.FirstDisplayedScrollingRowIndex = dgvV.Rows.Count - 1;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("La carpeta solo debe tener retenciones electrónicas", "Error al importar retenciones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private bool importRetVersion2(XDocument xmlDoc)
        {
            bool import = false;
            string tax = "retencion";

            // Verificar q maximo tenga 2 impuestos
            if (xmlDoc.Descendants(tax).Count() < 3)
            {
                string comprobante = "";
                int porcieniva = 0;
                double porcienrenta = 0;
                double valRetIva = 0;
                double valRetRenta = 0;

                foreach (var impuesto in xmlDoc.Descendants(tax))
                {
                    //NOTA: En la version 2 el numDocSustento esta fuera de las etiquetas retenciones
                    comprobante = xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value;

                    switch (Int32.Parse(impuesto.Descendants("codigo").FirstOrDefault().Value))
                    {
                        case 1:
                            porcienrenta = double.Parse(impuesto.Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            valRetRenta = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec));
                            break;
                        case 2:
                            porcieniva = (int)Math.Round(double.Parse(impuesto.Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec)));
                            valRetIva = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec));
                            break;
                    }
                }

                //Si la retencion es de una factura entonces el numero de compronte va ser mayor a 0
                if (Int64.Parse(comprobante) > 0)
                {
                    foreach (DataGridViewRow row in dgvV.Rows)
                    {
                        if (row.Cells[5].Value != null && row.Cells[5].Value.ToString().Replace("-", "").Equals(comprobante))
                        {
                            row.Cells[12].Value = porcienrenta;
                            row.Cells[13].Value = porcieniva;
                            row.Cells[14].Value = valRetRenta;
                            row.Cells[15].Value = valRetIva;
                            row.Cells[16].Value = xmlDoc.Descendants("secuencial").FirstOrDefault().Value;
                            import = true;
                            break;
                        }
                    }
                }
                //Se emiten retenciones de otros comprobantes y se registra como nuevo registro
                else
                {
                    dgvV.Rows.Add(
                        xmlDoc.Descendants("ruc").FirstOrDefault().Value,
                        "",     //Razon social
                        "04",
                        //xmlDoc.Descendants("tipoIdentificacionSujetoRetenido").FirstOrDefault().Value,
                        xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value,     //Fecha
                        "F",    //Tipo comprobante venta
                        "",
                        0, 0, 0,
                        0,      //Monto ICE
                        0, 0,   //IVA y Total
                        porcienrenta,
                        porcieniva,
                        valRetRenta,
                        valRetIva,
                        xmlDoc.Descendants("secuencial").FirstOrDefault().Value
                    );
                    import = true;
                }
            }

            return import;
        }

        private bool importRetVersion1(XDocument xmlDoc)
        {
            bool import = false;
            string tax = "impuesto";

            // Verificar q maximo tenga 2 impuestos
            if (xmlDoc.Descendants(tax).Count() < 3)
            {
                string tcv = "";
                string comprobante = "";
                int porcieniva = 0;
                double porcienrenta = 0;
                double valRetIva = 0;
                double valRetRenta = 0;

                foreach (var impuesto in xmlDoc.Descendants(tax))
                {
                    //El codDocSustento esta dentro de la etiqueta impuesto y si no existe le ponemos por defecto NULL
                    tcv = impuesto.Descendants("codDocSustento").Any() ? impuesto.Descendants("codDocSustento").FirstOrDefault().Value : null;
                    if (tcv == "01")
                    {
                        comprobante = impuesto.Descendants("numDocSustento").FirstOrDefault().Value;
                    }

                    switch (Int32.Parse(impuesto.Descendants("codigo").FirstOrDefault().Value))
                    {
                        case 1:
                            porcienrenta = double.Parse(impuesto.Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            valRetRenta = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec));
                            break;
                        case 2:
                            porcieniva = (int)Math.Round(double.Parse(impuesto.Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec)));
                            valRetIva = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec));
                            break;
                    }
                }

                //Se emiten retenciones de facturas entonces se debe buscar y registrar su retencion correspondiente
                if (tcv == "01")
                {
                    foreach (DataGridViewRow row in dgvV.Rows)
                    {
                        if (row.Cells[5].Value != null && row.Cells[5].Value.ToString().Replace("-", "").Equals(comprobante))
                        {
                            row.Cells[12].Value = porcienrenta;
                            row.Cells[13].Value = porcieniva;
                            row.Cells[14].Value = valRetRenta;
                            row.Cells[15].Value = valRetIva;
                            row.Cells[16].Value = xmlDoc.Descendants("secuencial").FirstOrDefault().Value;
                            import = true;
                            break;
                        }
                    }
                }
                //Se emiten retenciones de otros comprobantes y se registra como nuevo registro
                else
                {
                    dgvV.Rows.Add(
                        xmlDoc.Descendants("ruc").FirstOrDefault().Value,
                        "",     //Razon social
                        "04",
                        //xmlDoc.Descendants("tipoIdentificacionSujetoRetenido").FirstOrDefault().Value,
                        xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value,     //Fecha
                        "F",    //Tipo comprobante venta
                        "",
                        0, 0, 0,
                        0,      //Monto ICE
                        0, 0,   //IVA y Total
                        porcienrenta,
                        porcieniva,
                        valRetRenta,
                        valRetIva,
                        xmlDoc.Descendants("secuencial").FirstOrDefault().Value
                    );
                    import = true;
                }
            }

            return import;
        }

        public void importRecuperado()
        {
            try
            {
                string path = selectFile("xml");
                if (path != null && (path.EndsWith(".xml") || path.EndsWith(".XML")))
                {
                    XElement xmlDoc = XElement.Load(path);

                    if (decimal.Parse(xmlDoc.Descendants("totalVentas").FirstOrDefault().Value.Replace('.', ',')) == 0)
                    {
                        MessageBox.Show("El ATS tiene ventas en cero", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (xmlDoc.Descendants("detalleVentas").Any())
                    {
                        foreach (var detalleCompras in xmlDoc.Descendants("detalleVentas"))
                        {
                            var tipoComprobante = detalleCompras.Descendants("tipoComprobante").FirstOrDefault().Value;

                            var baseImponible = decimal.Parse(detalleCompras.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', ','));
                            var baseImpGrav = decimal.Parse(detalleCompras.Descendants("baseImpGrav").FirstOrDefault().Value.Replace('.', ','));
                            var montoIce = detalleCompras.Descendants("montoIce").Any() ? decimal.Parse(detalleCompras.Descendants("montoIce").FirstOrDefault().Value.Replace('.', ',')) : 0;
                            var montoIva = decimal.Parse(detalleCompras.Descendants("montoIva").FirstOrDefault().Value.Replace('.', ','));

                            decimal valorRetRenta = decimal.Parse(detalleCompras.Descendants("valorRetRenta").FirstOrDefault().Value.Replace('.', ','));
                            decimal valorRetIva = decimal.Parse(detalleCompras.Descendants("valorRetIva").FirstOrDefault().Value.Replace('.', ','));

                            dgvV.Rows.Add(
                                detalleCompras.Descendants("idCliente").FirstOrDefault().Value,
                                "",//Cliente
                                detalleCompras.Descendants("tpIdCliente").FirstOrDefault().Value,
                                "",//No hay fecha
                                tipoComprobante == "18" ? "F" : (tipoComprobante == "04" ? "N/C" : ""),
                                "",//# Comprobante
                                (double)(baseImponible + baseImpGrav),
                                (double)baseImponible,
                                (double)baseImpGrav,
                                (double)montoIce,
                                (double)montoIva,
                                (double)(baseImponible + baseImpGrav + montoIva),  //Total
                                valorRetRenta > 0 && (baseImponible + baseImpGrav) > 0 ? Math.Round(100 * valorRetRenta / (baseImponible + baseImpGrav), 2) : 0,//%Renta
                                valorRetIva > 0 && montoIva > 0 ? Math.Round(100 * valorRetIva / montoIva) : 0,//%Iva
                                (double)valorRetRenta,
                                (double)valorRetIva,
                                ""
                                );
                        }
                        dgvV.FirstDisplayedScrollingRowIndex = dgvV.Rows.Count - 1;
                    }
                    else
                    {
                        MessageBox.Show("El ATS no tiene ventas", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No es un ATS recuperado del SRI", "Error al importar ATS recuperado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void importRecuperadoMasivo()
        {
            try
            {
                Importa importa = new Importa();
                string pathFolder = importa.selectFolder();
                if (pathFolder != null)
                {
                    string[] files = Directory.GetFiles(pathFolder);

                    foreach (var path in files)
                    {
                        if (path.EndsWith(".xml") || path.EndsWith(".XML"))
                        {
                            XElement xmlDoc = XElement.Load(path);

                            if (xmlDoc.Descendants("detalleVentas").Any())
                            {
                                foreach (var detalleCompras in xmlDoc.Descendants("detalleVentas"))
                                {
                                    var tipoComprobante = detalleCompras.Descendants("tipoComprobante").FirstOrDefault().Value;

                                    var baseImponible = decimal.Parse(detalleCompras.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', ','));
                                    var baseImpGrav = decimal.Parse(detalleCompras.Descendants("baseImpGrav").FirstOrDefault().Value.Replace('.', ','));
                                    var montoIce = detalleCompras.Descendants("montoIce").Any() ? decimal.Parse(detalleCompras.Descendants("montoIce").FirstOrDefault().Value.Replace('.', ',')) : 0;
                                    var montoIva = decimal.Parse(detalleCompras.Descendants("montoIva").FirstOrDefault().Value.Replace('.', ','));

                                    decimal valorRetRenta = decimal.Parse(detalleCompras.Descendants("valorRetRenta").FirstOrDefault().Value.Replace('.', ','));
                                    decimal valorRetIva = decimal.Parse(detalleCompras.Descendants("valorRetIva").FirstOrDefault().Value.Replace('.', ','));

                                    dgvV.Rows.Add(
                                        detalleCompras.Descendants("idCliente").FirstOrDefault().Value,
                                        "",//Cliente
                                        detalleCompras.Descendants("tpIdCliente").FirstOrDefault().Value,
                                        "17-" + xmlDoc.Descendants("Mes").FirstOrDefault().Value + "-" + xmlDoc.Descendants("Anio").FirstOrDefault().Value,//No hay fecha
                                        tipoComprobante == "18" ? "F" : (tipoComprobante == "04" ? "N/C" : ""),
                                        "",//# Comprobante
                                        (double)(baseImponible + baseImpGrav),
                                        (double)baseImponible,
                                        (double)baseImpGrav,
                                        (double)montoIce,
                                        (double)montoIva,
                                        (double)(baseImponible + baseImpGrav + montoIva),  //Total
                                        valorRetRenta > 0 && (baseImponible + baseImpGrav) > 0 ? Math.Round(100 * valorRetRenta / (baseImponible + baseImpGrav), 2) : 0,//%Renta
                                        valorRetIva > 0 && montoIva > 0 ? Math.Round(100 * valorRetIva / montoIva) : 0,//%Iva
                                        (double)valorRetRenta,
                                        (double)valorRetIva,
                                        ""
                                        );
                                }
                                dgvV.FirstDisplayedScrollingRowIndex = dgvV.Rows.Count - 1;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No es un ATS recuperado del SRI", "Error al importar ATS recuperado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        ~DgvVentas() { }
    }
}