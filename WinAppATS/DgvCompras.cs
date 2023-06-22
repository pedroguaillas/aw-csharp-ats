using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
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
    class DgvCompras
    {
        DataGridView dgvCompras;
        string info;
        NumberFormatInfo nfi;

        char dec;

        public DgvCompras(DataGridView dgvCompras)
        {
            this.dgvCompras = dgvCompras;
            //this.nfi = new NumberFormatInfo();
            //nfi.NumberDecimalSeparator = ",";
            nfi = NumberFormatInfo.CurrentInfo;
            dec = ',';
        }

        public void Load(string info)
        {
            this.info = info;
            if (File.Exists(Const.filexml(info)))
            {
                using (XmlReader xml = XmlReader.Create(Const.filexml(info)))
                {
                    DataSet data = new DataSet("compras");
                    data.ReadXml(xml);
                    DataTable dt = data.Tables[0];

                    foreach (DataRow dr in dt.Rows)
                    {
                        dgvCompras.Rows.Add(dr.ItemArray);
                    }
                }
            }
        }

        public async Task<string> toXml()
        {
            DataTable dt = new DataTable("compra");
            for (int i = 1; i < dgvCompras.Columns.Count + 1; i++)
            {
                string column = dgvCompras.Columns[i - 1].Name;
                dt.Columns.Add(column, typeof(string));
            }

            int ColumnCount = dgvCompras.Columns.Count;
            foreach (DataGridViewRow dr in dgvCompras.Rows)
            {
                DataRow dataRow = dt.NewRow();
                for (int i = 0; i < ColumnCount; i++)
                {
                    dataRow[i] = dr.Cells[i].Value;
                }
                dt.Rows.Add(dataRow);
            }

            DataSet ds = new DataSet("compras");
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

        private async Task<string> Guardar()
        {
            string result;
            byte[] bytes = File.ReadAllBytes(Const.filexml(info));

            using (var client = new HttpClient())
            {
                var data = new { file = Convert.ToBase64String(bytes), info = info.Substring(1), tipo = "filecompra" };

                string json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var result1 = await client.PostAsync(Const.URL + "archivos", content);
                result = await result1.Content.ReadAsStringAsync();
            }
            return result;
        }

        public void Importar(CheckBox checkBox)
        {
            try
            {
                Importa importa = new Importa();
                string path = importa.selectFolder();
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

                            if (codDoc == "01" || codDoc == "03" || codDoc == "04" || codDoc == "05")
                            {
                                cargarRegistro(xmlDoc, checkBox.Checked, XTemp.Descendants("numeroAutorizacion").FirstOrDefault().Value);

                                if (!inserts.Contains(xmlDoc.Descendants("ruc").FirstOrDefault().Value))
                                {
                                    var obj = new
                                    {
                                        id = xmlDoc.Descendants("ruc").FirstOrDefault().Value,
                                        denominacion = xmlDoc.Descendants("razonSocial" + (codDoc == "03" ? "Proveedor" : "")).FirstOrDefault().Value.Trim(),
                                        tpId = codDoc == "03" ? "02" : "01",
                                        //tpIdProv = "", //Especial se registra en el servidor con el valor 01-RUC
                                        //tipoProv = "", //Se calcula en servidor en base al tercer digo de la identificacion 01-Persona natural / 02-Sociedad
                                        //xmlDoc.Descendants("contribuyenteEspecial").Any() //Dato extra
                                        contabilidad = xmlDoc.Descendants("obligadoContabilidad").Any() ? xmlDoc.Descendants("obligadoContabilidad").FirstOrDefault().Value : ""
                                    };

                                    insertados.Add(obj);
                                    inserts.Add(obj.id);
                                }
                                i++;
                            }
                        }
                    }
                    ImportContactos import = new ImportContactos();
                    import.registrarList(insertados);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("La carpeta solo debe contener comprobantes electrónicos", "Error al importar comprobantes electrónicos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void cargarRegistro(XDocument xmlDoc, bool gasolinera, string autorizacion)
        {
            double imponible = 0;
            double base0 = 0;
            double base12 = 0;
            double ice = 0;
            double descuentoAdicional = 0;
            double iva = 0;

            foreach (var totalImpuesto in xmlDoc.Descendants("totalImpuesto"))
            {
                switch (Int32.Parse(totalImpuesto.Descendants("codigoPorcentaje").FirstOrDefault().Value))
                {
                    case 0: base0 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    case 2: base12 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    case 3: base12 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    case 8: base12 += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    case 6: imponible += Math.Round(double.Parse(totalImpuesto.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2); break;
                    default:
                        if (Int32.Parse(totalImpuesto.Descendants("codigo").FirstOrDefault().Value) == 3)
                        {
                            ice += Math.Round(double.Parse(totalImpuesto.Descendants("valor").FirstOrDefault().Value.Replace('.', dec)), 2);
                        }
                        break;
                }

                iva += Math.Round(double.Parse(totalImpuesto.Descendants("valor").FirstOrDefault().Value.Replace('.', dec)), 2);

                if (totalImpuesto.Descendants("descuentoAdicional").Any())
                {
                    descuentoAdicional += double.Parse(totalImpuesto.Descendants("descuentoAdicional").FirstOrDefault().Value.Replace('.', dec));
                }
            }

            string codCombustible = "";
            //Si quiere detectar las facturas de gasolina o diesel
            //if (gasolinera && xmlDoc.Elements("detalle").Count() == 1)
            //{
            //    switch (xmlDoc.Descendants("codigoPrincipal").FirstOrDefault().Value)
            //    {
            //        case "0103": codCombustible = "SÚPER"; break;
            //        case "0101": codCombustible = "EXTRA"; break;
            //        case "0174": codCombustible = "EXTRAE"; break;
            //        case "0121": codCombustible = "DIESELP"; break;
            //        case "0104": codCombustible = "DIESEL2"; break;
            //    }
            //}

            string infoA = codCombustible;

            if (xmlDoc.Descendants("tipoIdentificacionComprador").Any() && xmlDoc.Descendants("tipoIdentificacionComprador").FirstOrDefault().Value == "05")
            {
                infoA += " xxx";
            }

            if (descuentoAdicional > 0)
            {
                infoA += " DA";
            }

            string codDoc = xmlDoc.Descendants("codDoc").FirstOrDefault().Value;

            dgvCompras.Rows.Add(
            infoA,
            xmlDoc.Descendants(codDoc == "03" ? "identificacionProveedor" : "ruc").FirstOrDefault().Value,
            removeOtherCharacter(xmlDoc.Descendants("razonSocial" + (codDoc == "03" ? "Proveedor" : "")).FirstOrDefault().Value.Trim()),
            "", //Cod cuenta
            "", //Detalle cuenta
            typeCom(codDoc), //verificar si no por defecto factura.
            xmlDoc.Descendants("fechaEmision").FirstOrDefault().Value,
            xmlDoc.Descendants("estab").FirstOrDefault().Value,
            xmlDoc.Descendants("ptoEmi").FirstOrDefault().Value,
            xmlDoc.Descendants("secuencial").FirstOrDefault().Value,
            autorizacion,
            imponible, base0, base12, 0, ice,
            iva,    //IVA
            base0 + base12 + iva,   //Total

            //Retenciones
            0, 0, 0, 0, 0, 0,

            //Info comprobante de retencion
            "", "", "", "",

            //Retenciones
            "", "", "", 0,

            //Nota de debito o credito
            (codDoc == "04" || codDoc == "05") ? xmlDoc.Descendants("numDocModificado").FirstOrDefault().Value.Substring(0, 3) : "",
            (codDoc == "04" || codDoc == "05") ? xmlDoc.Descendants("numDocModificado").FirstOrDefault().Value.Substring(4, 3) : "",
            (codDoc == "04" || codDoc == "05") ? xmlDoc.Descendants("numDocModificado").FirstOrDefault().Value.Substring(8, 9) : ""
            );

            if (descuentoAdicional > 0)
            {
                dgvCompras.Rows[dgvCompras.RowCount - 1].Cells[17].Style.BackColor = Color.Yellow;
            }
        }

        public string typeCom(string cod)
        {
            string result = "";

            switch (cod)
            {
                case "01": result = "F"; break;
                case "03": result = "L/C"; break;
                case "04": result = "N/C"; break;
                case "05": result = "N/D"; break;
            }

            return result;
        }

        public void importReport(ProgressBar bar)
        {
            Importa importa = new Importa();
            string file = importa.selectFile("txt");

            if (file == null) { return; }

            List<String> claves = extraerAutorizaciones(file);

            if (claves != null && claves.Count > 0)
            {
                try
                {
                    List<string> inserts = new List<string>();
                    List<Object> insertados = new List<Object>();
                    bar.Visible = true;
                    bar.Value = 0;
                    int i = 0;
                    int lent = claves.Count;

                    string problems = "";

                    foreach (var c in claves)
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
                                    cargarRegistro(xmlDoc, false, docResponse.GetElementsByTagName("numeroAutorizacion")[0].InnerXml.Trim());

                                    if (!inserts.Contains(xmlDoc.Descendants("ruc").FirstOrDefault().Value))
                                    {
                                        var obj = new
                                        {
                                            id = xmlDoc.Descendants("ruc").FirstOrDefault().Value,
                                            denominacion = xmlDoc.Descendants("razonSocial").FirstOrDefault().Value.Trim(),
                                            tpId = "01",
                                            contabilidad = xmlDoc.Descendants("obligadoContabilidad").Any() ? xmlDoc.Descendants("obligadoContabilidad").FirstOrDefault().Value : ""
                                        };

                                        insertados.Add(obj);
                                        inserts.Add(obj.id);
                                    }
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
                    bar.Visible = false;

                    if (problems.Length > 0)
                    {
                        file = @"" + file.Remove(file.Length - 4) + "_Error.txt";
                        Archivo archivo = new Archivo();
                        archivo.saveError(problems, file);
                    }

                    dgvCompras.FirstDisplayedScrollingRowIndex = dgvCompras.Rows.Count - 1;

                    ImportContactos import = new ImportContactos();
                    import.registrarList(insertados);
                }
                catch (Exception ex)
                {
                    bar.Visible = false;
                    MessageBox.Show("Se produjo un error al descargar comprobantes electrónicos" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("No existe comprobantes", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //public void descargar(string ruc, ProgressBar bar)
        //{
        //    Importa importa = new Importa();
        //    string path = importa.selectFile("txt");

        //    if (path != null)
        //    {
        //        string line = "";
        //        StreamReader countlines = new StreamReader(path);
        //        int linesLength = 0;
        //        List<string> claves = new List<string>();
        //        while ((line = countlines.ReadLine()) != null)
        //        {
        //            string[] palabras = line.Split('\t');
        //            //Si palabras tiene una sola palabra y la palabra tiene 52 caracteres
        //            if (palabras.Count() == 1 && palabras[0].Length == 52)
        //            {
        //                //extraer solo la clave de acceso
        //                string claveacceso = palabras[0].Substring(3);

        //                // Verificar que todos los caracteres sean digitos
        //                bool verdad = true;

        //                foreach (var letra in claveacceso)
        //                {
        //                    if (!Char.IsDigit(letra))
        //                    {
        //                        verdad = false;
        //                    }
        //                }

        //                if (verdad && !claves.Contains(claveacceso))
        //                {
        //                    claves.Add(claveacceso);
        //                    linesLength++;
        //                }
        //            }
        //        }

        //        bar.Visible = true;
        //        bar.Value = 0;

        //        string problems = "";

        //        for (int i = 0; i < claves.Count; i++)
        //        {
        //            AutorizacionFacturas autorizacion = new AutorizacionFacturas();

        //            if (!autorizacion.Descarga(claves[i], path.Remove(path.Length - 4)))
        //            {
        //                if (!autorizacion.Descarga(claves[i], path.Remove(path.Length - 4)))
        //                {
        //                    if (!autorizacion.Descarga(claves[i], path.Remove(path.Length - 4)))
        //                    {
        //                        problems += "\n" + claves[i];
        //                    }
        //                }
        //            }
        //            bar.Value = (i * 100) / linesLength;
        //        }

        //        if (problems.Length > 0)
        //        {
        //            path = @"" + path.Remove(path.Length - 4) + "/Error.txt";
        //            Archivo archivo = new Archivo();
        //            archivo.saveError(problems, path);
        //        }

        //        bar.Value = 100;
        //        bar.Visible = false;
        //    }
        //}

        public void descargar(string ruc, ProgressBar bar)
        {
            Importa importa = new Importa();
            string path = importa.selectFile("txt");

            if (path != null)
            {
                string line = "";
                StreamReader countlines = new StreamReader(path);
                int linesLength = 0;
                while ((line = countlines.ReadLine()) != null)
                {
                    linesLength++;
                }

                int i = 0;

                StreamReader file = new StreamReader(path);

                bar.Visible = true;
                bar.Value = 0;

                string problems = "";

                while ((line = file.ReadLine()) != null)
                {
                    if (i > 0)
                    {
                        string[] palabras = line.Split('\t');
                        if (palabras[9].Length == 49 || palabras[10].Length == 49)
                        {
                            AutorizacionFacturas autorizacion = new AutorizacionFacturas();
                            string code = palabras[palabras[9].Length == 49 ? 9 : 10];

                            if (!autorizacion.Descarga(code, path.Remove(path.Length - 4)))
                            {
                                if (!autorizacion.Descarga(code, path.Remove(path.Length - 4)))
                                {
                                    if (!autorizacion.Descarga(code, path.Remove(path.Length - 4)))
                                    {
                                        problems += "\n" + code;
                                    }
                                }
                            }
                        }
                    }
                    bar.Value = (i * 100) / linesLength;
                    i++;
                }

                if (problems.Length > 0)
                {
                    path = @"" + path.Remove(path.Length - 4) + "/Error.txt";
                    Archivo archivo = new Archivo();
                    archivo.saveError(problems, path);
                }

                bar.Value = 100;
                bar.Visible = false;
            }
        }

        private List<string> extraerAutorizaciones(string path)
        {
            List<String> cla_accs = new List<string>();
            bool error = false;
            try
            {
                if (path != null)
                {
                    string line = "";
                    StreamReader file = new StreamReader(path);
                    int i = 0;
                    while ((line = file.ReadLine()) != null)
                    {
                        if (i > 0)
                        {
                            string[] palabras = line.Split('\t');
                            if (palabras[9].Length == 49 || palabras[10].Length == 49)
                            {
                                cla_accs.Add(palabras[palabras[9].Length == 49 ? 9 : 10]);
                            }
                            else
                            {
                                error = true;
                                MessageBox.Show("El reporte de compras no tiene una estructura correcta");
                            }
                        }
                        i++;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("El archivo no es reporte del SRI");
            }
            if (error) { return null; }
            return cla_accs;
        }

        public void importRecuperado()
        {
            try
            {
                Importa importa = new Importa();
                string path = importa.selectFile("xml");
                if (path != null && (path.EndsWith(".xml") || path.EndsWith(".XML")))
                {
                    XElement xmlDoc = XElement.Load(path);

                    Retencion retencion = new Retencion();

                    List<Object> insertados = new List<Object>();
                    List<String> inserts = new List<String>();

                    if (xmlDoc.Descendants("detalleCompras").Any())
                    {
                        foreach (var detalleCompras in xmlDoc.Descendants("detalleCompras"))
                        {
                            var idProv = detalleCompras.Descendants("idProv").FirstOrDefault().Value;
                            var razonsocial = detalleCompras.Elements("denoProv").Any() ? detalleCompras.Descendants("denoProv").FirstOrDefault().Value : "";

                            var tipoComprobante = detalleCompras.Descendants("tipoComprobante").FirstOrDefault().Value;
                            var codAts = detalleCompras.Descendants("codRetAir").Any() ? detalleCompras.Descendants("codRetAir").FirstOrDefault().Value : "";

                            var baseImponible = decimal.Parse(detalleCompras.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', ','));
                            var baseImpGrav = decimal.Parse(detalleCompras.Descendants("baseImpGrav").FirstOrDefault().Value.Replace('.', ','));
                            var montoIce = decimal.Parse(detalleCompras.Descendants("montoIce").FirstOrDefault().Value.Replace('.', ','));
                            var montoIva = decimal.Parse(detalleCompras.Descendants("montoIva").FirstOrDefault().Value.Replace('.', ','));

                            DataRow dataRow = retencion.completeRetenciones(codAts);

                            dgvCompras.Rows.Add(
                                "",
                                idProv,
                                razonsocial,
                                "", //Cod cuenta
                                "", //Detalle cuenta
                                tipoComprobante == "01" ? "F" : (tipoComprobante == "02" ? "N/V" : (tipoComprobante == "03" ? "L/C" : (tipoComprobante == "04" ? "N/C" : (tipoComprobante == "05" ? "N/D" : "")))),
                                detalleCompras.Descendants("fechaRegistro").FirstOrDefault().Value,
                                detalleCompras.Descendants("establecimiento").FirstOrDefault().Value,
                                detalleCompras.Descendants("puntoEmision").FirstOrDefault().Value,
                                detalleCompras.Descendants("secuencial").FirstOrDefault().Value,
                                detalleCompras.Descendants("autorizacion").FirstOrDefault().Value,

                                double.Parse(detalleCompras.Descendants("baseNoGraIva").FirstOrDefault().Value.Replace('.', ',')),
                                (double)baseImponible,
                                (double)baseImpGrav,
                                0,
                                (double)montoIce,
                                (double)montoIva,
                                (double)(baseImponible + baseImpGrav + montoIva),  //Total

                                //Retenciones
                                detalleCompras.Elements("valRetBien10").Any() ? double.Parse(detalleCompras.Descendants("valRetBien10").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                detalleCompras.Elements("valRetServ20").Any() ? double.Parse(detalleCompras.Descendants("valRetServ20").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                detalleCompras.Elements("valorRetBienes").Any() ? double.Parse(detalleCompras.Descendants("valorRetBienes").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                detalleCompras.Elements("valRetServ50").Any() ? double.Parse(detalleCompras.Descendants("valRetServ50").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                detalleCompras.Elements("valorRetServicios").Any() ? double.Parse(detalleCompras.Descendants("valorRetServicios").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                detalleCompras.Elements("valRetServ100").Any() ? double.Parse(detalleCompras.Descendants("valRetServ100").FirstOrDefault().Value.Replace('.', ',')) : 0,

                                //Info comprobante de retencion
                                detalleCompras.Descendants("estabRetencion1").Any() ? detalleCompras.Descendants("estabRetencion1").FirstOrDefault().Value : "",
                                detalleCompras.Descendants("ptoEmiRetencion1").Any() ? detalleCompras.Descendants("ptoEmiRetencion1").FirstOrDefault().Value : "",
                                detalleCompras.Descendants("secRetencion1").Any() ? detalleCompras.Descendants("secRetencion1").FirstOrDefault().Value : "",
                                detalleCompras.Descendants("autRetencion1").Any() ? detalleCompras.Descendants("autRetencion1").FirstOrDefault().Value : "",

                                //Retenciones
                                codAts,
                                detalleCompras.Descendants("codRetAir").Any() && dataRow != null ? dataRow[1] : "",
                                detalleCompras.Descendants("porcentajeAir").Any() ? double.Parse(detalleCompras.Descendants("porcentajeAir").FirstOrDefault().Value.Replace('.', ',')).ToString() : "",
                                detalleCompras.Descendants("valRetAir").Any() ? double.Parse(detalleCompras.Descendants("valRetAir").FirstOrDefault().Value.Replace('.', ',')) : 0,

                                //Nota de debito o credito
                                detalleCompras.Descendants("estabModificado").Any() ? detalleCompras.Descendants("estabModificado").FirstOrDefault().Value : "",
                                detalleCompras.Descendants("ptoEmiModificado").Any() ? detalleCompras.Descendants("ptoEmiModificado").FirstOrDefault().Value : "",
                                detalleCompras.Descendants("secModificado").Any() ? detalleCompras.Descendants("secModificado").FirstOrDefault().Value : ""
                                );

                            if (detalleCompras.Descendants("detalleAir").Count() > 1)
                            {
                                var detalleAir = detalleCompras.Descendants("detalleAir");
                                for (int i = 1; i < detalleAir.Count(); i++)
                                {
                                    dgvCompras.Rows[dgvCompras.Rows.Count - 1].Cells[9].Style.BackColor = Color.Red;

                                    DataGridViewRow row = dgvCompras.Rows[dgvCompras.Rows.Count - 1];

                                    //Duplicar fila clone
                                    DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                                    for (Int32 index = 0; index < row.Cells.Count; index++)
                                    {
                                        clonedRow.Cells[index].Value = row.Cells[index].Value;
                                    }

                                    for (int ci = 11; ci < 24; ci++)
                                    {
                                        clonedRow.Cells[ci].Value = 0;
                                    }
                                    codAts = detalleAir.ElementAt(i).Descendants("codRetAir").FirstOrDefault().Value;
                                    clonedRow.Cells[28].Value = codAts;

                                    dataRow = retencion.completeRetenciones(codAts);
                                    clonedRow.Cells[29].Value = dataRow != null ? dataRow[1] : "";
                                    clonedRow.Cells[30].Value = detalleAir.ElementAt(i).Descendants("porcentajeAir").FirstOrDefault().Value.Replace('.', ',');
                                    clonedRow.Cells[31].Value = detalleAir.ElementAt(i).Descendants("valRetAir").FirstOrDefault().Value.Replace('.', ',');
                                    dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                                }
                            }

                            if (razonsocial.Length > 0 && !inserts.Contains(idProv))
                            {
                                var obj = new
                                {
                                    id = idProv,
                                    denominacion = razonsocial,
                                    tpId = detalleCompras.Descendants("tpIdProv").Any() ? detalleCompras.Descendants("tpIdProv").FirstOrDefault().Value : "01",
                                    //tpIdProv = "", //Especial se registra en el servidor con el valor 01-RUC
                                    //tipoProv = "", //Se calcula en servidor en base al tercer digo de la identificacion 01-Persona natural / 02-Sociedad
                                    //xmlDoc.Descendants("contribuyenteEspecial").Any() //Dato extra
                                    contabilidad = ""
                                };

                                insertados.Add(obj);
                                inserts.Add(obj.id);
                            }
                        }
                        if (insertados.Count > 0)
                        {
                            ImportContactos import = new ImportContactos();
                            import.registrarList(insertados);
                        }
                    }
                    else
                    {
                        MessageBox.Show("El ATS no tiene compras", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No se pudo recuperar el ATS", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void importRecuperadoMasiva()
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
                        if (path != null && (path.EndsWith(".xml") || path.EndsWith(".XML")))
                        {
                            XElement xmlDoc = XElement.Load(path);
                            Retencion retencion = new Retencion();

                            if (xmlDoc.Descendants("detalleCompras").Any())
                            {
                                foreach (var detalleCompras in xmlDoc.Descendants("detalleCompras"))
                                {
                                    var idProv = detalleCompras.Descendants("idProv").FirstOrDefault().Value;
                                    var razonsocial = detalleCompras.Elements("denoProv").Any() ? detalleCompras.Descendants("denoProv").FirstOrDefault().Value : "";

                                    var tipoComprobante = detalleCompras.Descendants("tipoComprobante").FirstOrDefault().Value;
                                    var codAts = detalleCompras.Descendants("codRetAir").Any() ? detalleCompras.Descendants("codRetAir").FirstOrDefault().Value : "";

                                    var baseImponible = decimal.Parse(detalleCompras.Descendants("baseImponible").FirstOrDefault().Value.Replace('.', ','));
                                    var baseImpGrav = decimal.Parse(detalleCompras.Descendants("baseImpGrav").FirstOrDefault().Value.Replace('.', ','));
                                    var montoIce = decimal.Parse(detalleCompras.Descendants("montoIce").FirstOrDefault().Value.Replace('.', ','));
                                    var montoIva = decimal.Parse(detalleCompras.Descendants("montoIva").FirstOrDefault().Value.Replace('.', ','));

                                    DataRow dataRow = retencion.completeRetenciones(codAts);

                                    dgvCompras.Rows.Add(
                                        "",
                                        idProv,
                                        razonsocial,
                                        "", //Cod cuenta
                                        "", //Detalle cuenta
                                        tipoComprobante == "01" ? "F" : (tipoComprobante == "02" ? "N/V" : (tipoComprobante == "03" ? "L/C" : (tipoComprobante == "04" ? "N/C" : (tipoComprobante == "05" ? "N/D" : "")))),
                                        detalleCompras.Descendants("fechaRegistro").FirstOrDefault().Value,
                                        detalleCompras.Descendants("establecimiento").FirstOrDefault().Value,
                                        detalleCompras.Descendants("puntoEmision").FirstOrDefault().Value,
                                        detalleCompras.Descendants("secuencial").FirstOrDefault().Value,
                                        detalleCompras.Descendants("autorizacion").FirstOrDefault().Value,

                                        double.Parse(detalleCompras.Descendants("baseNoGraIva").FirstOrDefault().Value.Replace('.', ',')),
                                        (double)baseImponible,
                                        (double)baseImpGrav,
                                        0,
                                        (double)montoIce,
                                        (double)montoIva,
                                        (double)(baseImponible + baseImpGrav + montoIva),  //Total

                                        //Retenciones
                                        detalleCompras.Elements("valRetBien10").Any() ? double.Parse(detalleCompras.Descendants("valRetBien10").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                        detalleCompras.Elements("valRetServ20").Any() ? double.Parse(detalleCompras.Descendants("valRetServ20").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                        detalleCompras.Elements("valorRetBienes").Any() ? double.Parse(detalleCompras.Descendants("valorRetBienes").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                        detalleCompras.Elements("valRetServ50").Any() ? double.Parse(detalleCompras.Descendants("valRetServ50").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                        detalleCompras.Elements("valorRetServicios").Any() ? double.Parse(detalleCompras.Descendants("valorRetServicios").FirstOrDefault().Value.Replace('.', ',')) : 0,
                                        detalleCompras.Elements("valRetServ100").Any() ? double.Parse(detalleCompras.Descendants("valRetServ100").FirstOrDefault().Value.Replace('.', ',')) : 0,

                                        //Info comprobante de retencion
                                        detalleCompras.Descendants("estabRetencion1").Any() ? detalleCompras.Descendants("estabRetencion1").FirstOrDefault().Value : "",
                                        detalleCompras.Descendants("ptoEmiRetencion1").Any() ? detalleCompras.Descendants("ptoEmiRetencion1").FirstOrDefault().Value : "",
                                        detalleCompras.Descendants("secRetencion1").Any() ? detalleCompras.Descendants("secRetencion1").FirstOrDefault().Value : "",
                                        detalleCompras.Descendants("autRetencion1").Any() ? detalleCompras.Descendants("autRetencion1").FirstOrDefault().Value : "",

                                        //Retenciones
                                        codAts,
                                        detalleCompras.Descendants("codRetAir").Any() && dataRow != null ? dataRow[1] : "",
                                        detalleCompras.Descendants("porcentajeAir").Any() ? double.Parse(detalleCompras.Descendants("porcentajeAir").FirstOrDefault().Value.Replace('.', ',')).ToString() : "",
                                        detalleCompras.Descendants("valRetAir").Any() ? double.Parse(detalleCompras.Descendants("valRetAir").FirstOrDefault().Value.Replace('.', ',')) : 0,

                                        //Nota de debito o credito
                                        detalleCompras.Descendants("estabModificado").Any() ? detalleCompras.Descendants("estabModificado").FirstOrDefault().Value : "",
                                        detalleCompras.Descendants("ptoEmiModificado").Any() ? detalleCompras.Descendants("ptoEmiModificado").FirstOrDefault().Value : "",
                                        detalleCompras.Descendants("secModificado").Any() ? detalleCompras.Descendants("secModificado").FirstOrDefault().Value : ""
                                        );

                                    if (detalleCompras.Descendants("detalleAir").Count() > 1)
                                    {
                                        var detalleAir = detalleCompras.Descendants("detalleAir");
                                        for (int i = 1; i < detalleAir.Count(); i++)
                                        {
                                            dgvCompras.Rows[dgvCompras.Rows.Count - 1].Cells[9].Style.BackColor = Color.Red;

                                            DataGridViewRow row = dgvCompras.Rows[dgvCompras.Rows.Count - 1];

                                            //Duplicar fila clone
                                            DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                                            for (Int32 index = 0; index < row.Cells.Count; index++)
                                            {
                                                clonedRow.Cells[index].Value = row.Cells[index].Value;
                                            }

                                            for (int ci = 11; ci < 24; ci++)
                                            {
                                                clonedRow.Cells[ci].Value = 0;
                                            }
                                            codAts = detalleAir.ElementAt(i).Descendants("codRetAir").FirstOrDefault().Value;
                                            clonedRow.Cells[28].Value = codAts;

                                            dataRow = retencion.completeRetenciones(codAts);
                                            clonedRow.Cells[29].Value = dataRow != null ? dataRow[1] : "";
                                            clonedRow.Cells[30].Value = detalleAir.ElementAt(i).Descendants("porcentajeAir").FirstOrDefault().Value.Replace('.', ',');
                                            clonedRow.Cells[31].Value = detalleAir.ElementAt(i).Descendants("valRetAir").FirstOrDefault().Value.Replace('.', ',');
                                            dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                                        }
                                    }
                                }
                                dgvCompras.FirstDisplayedScrollingRowIndex = dgvCompras.Rows.Count - 1;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("No se pudo cargar el reporte del SRI", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string removeOtherCharacter(string ruc)
        {
            string result = "";
            ruc = ruc.ToUpper();
            ruc = ruc.Replace('Á', 'A');
            ruc = ruc.Replace('É', 'E');
            ruc = ruc.Replace('Í', 'I');
            ruc = ruc.Replace('Ó', 'O');
            ruc = ruc.Replace('Ú', 'U');
            ruc = ruc.Replace('Ñ', 'N');
            ruc = ruc.Replace('&', 'Y');

            result = ruc;
            foreach (char c in ruc.ToCharArray())
            {
                if ((c < 'A' || c > 'Z') && c != ' ')
                {
                    result = result.Replace(c.ToString(), string.Empty);
                }
            }
            return result;
        }

        public void ocultarRet(bool vis)
        {
            for (int i = 18; i < 32; i++)
            {
                dgvCompras.Columns[i].Visible = vis;
            }
        }
        public void ocultarNCredit(bool vis)
        {
            for (int i = 32; i < 35; i++)
            {
                dgvCompras.Columns[i].Visible = vis;
            }
        }
        public void ocultarBExeMIce(bool vis)
        {
            for (int i = 14; i < 16; i++)
            {
                dgvCompras.Columns[i].Visible = vis;
            }
        }

        public void generarCodigo()
        {
            if (dgvCompras.Rows.Count > 0)
            {
                dgvCompras.Rows[1].Cells[0].Value = dgvCompras.Rows[1].Cells[0].Value == null || dgvCompras.Rows[1].Cells[0].Value.Equals("") ? " " : dgvCompras.Rows[1].Cells[0].Value;
            }
        }

        public void addOne(string ruc, string razons, string tcv, string dia, string establecimiento, string ptoEmi, string secuencial, string autorizacion, string imponible, string base0, string base12, string estabMod, string ptoemiMod, string secuenMod)
        {
            double imp = double.Parse(imponible.Equals("") ? "0" : imponible);
            double b0 = double.Parse(base0.Equals("") ? "0" : base0);
            double b12 = double.Parse(base12.Equals("") ? "0" : base12);
            double iva = Math.Round(b12 * .12, 2);
            dgvCompras.Rows.Add(
                "",  //Codigo de compra debe ser nulo al final generar el codigo
                ruc,
                razons,
                "", //Cod cuenta
                "", //Detalle cuenta
                tcv, //verificar si no por defecto factura.
                string.Format(dia.PadLeft(2, '0') + "/" + info.Substring(info.Length - 2) + "/" + info.Substring(info.Length - 6, 4)),
                establecimiento.PadLeft(3, '0'),
                ptoEmi.PadLeft(3, '0'),
                secuencial.PadLeft(9, '0'),
                autorizacion,
                imp, b0, b12,
                0, 0,
                iva,//IVA
                Math.Round(b0 + b12 + iva, 2),//Total

                //Retenciones
                0, 0, 0, 0, 0, 0,

                //Info comprobante de retencion
                "", "", "", "",
                "",
                "",
                "",
                0,

                //Nota de debito
                estabMod,
                ptoemiMod,
                secuenMod
                );

            dgvCompras.FirstDisplayedScrollingRowIndex = dgvCompras.Rows.Count - 1;
        }

        public void importRetenciones()
        {
            try
            {
                Importa importa = new Importa();
                string path = importa.selectFolder();
                if (path != null)
                {
                    string[] files = Directory.GetFiles(path);
                    List<string> othersRet = new List<string>();

                    foreach (var file in files)
                    {
                        if (file.EndsWith(".xml"))
                        {
                            XElement XTemp = XElement.Load(file);

                            var queryCDATAXML = from element in XTemp.DescendantNodes()
                                                where element.NodeType == System.Xml.XmlNodeType.CDATA
                                                select element.Parent.Value.Trim();

                            string BodyHtml = queryCDATAXML.ToList<string>()[0].ToString();
                            XDocument xmlDoc = XDocument.Parse(BodyHtml);

                            //Determina que el comprobante sea una retencion
                            if (xmlDoc.Descendants("codDoc").FirstOrDefault().Value == "07" && !xmlDoc.Root.Attribute("version").Value.Trim().Equals("2.0.0"))
                            {
                                bool import = false;
                                //Determina el numero de reteciones que se aplica a la factura
                                switch (xmlDoc.Descendants("impuestos").Elements().Count())
                                {
                                    case 1:
                                        //Caso 1. Validar que la retencion sea de Renta.
                                        if (int.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigo").FirstOrDefault().Value) == 1)
                                        {
                                            import = aplicaCasoUno(XTemp.Descendants("numeroAutorizacion").FirstOrDefault().Value, xmlDoc);
                                        }
                                        break;
                                    case 2:
                                        //Caso 2. Tiene dos retenciones a la renta
                                        if (xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigo").FirstOrDefault().Value == xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("codigo").FirstOrDefault().Value)
                                        {
                                            //Validar que la retencion sea de Renta.
                                            if (int.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigo").FirstOrDefault().Value) == 1)
                                            {
                                                import = aplicaCasoDos(XTemp.Descendants("numeroAutorizacion").FirstOrDefault().Value, xmlDoc);
                                            }
                                        }
                                        //Caso 3. Tiene una rentencion a la renta y uno al IVA
                                        else
                                        {
                                            import = aplicaCasoTres(XTemp.Descendants("numeroAutorizacion").FirstOrDefault().Value, xmlDoc);
                                        }
                                        break;
                                        //case 3:
                                }
                                //En caso de tener mas de 2 retenciones
                                if (!import)
                                {
                                    othersRet.Add(xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value + "\t" + file);
                                    //othersRet.Add(xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value + "\t" + file.Substring(file.LastIndexOf("\"")));
                                    //othersRet.Add(xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value + "\t" + xmlDoc.Descendants("identificacionSujetoRetenido").FirstOrDefault().Value);
                                }
                            }
                        }
                    }
                    if (othersRet.Count > 0)
                    {
                        string invoices = othersRet.Count.ToString() + " retenciones no se han importado\n";
                        invoices += "Serie\t\tRetencion\n";
                        othersRet.ForEach(ret => invoices += ret + "\n");
                        MessageBox.Show(invoices);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("La carpeta solo debe tener retenciones electrónicas", "Error al importar retenciones", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //Casos > 1
        private bool aplicaCasoUno(string autorizacion, XDocument xmlDoc)
        {
            bool enc = false;
            int numDocSustento = int.Parse(xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value.Substring(6));
            string identificacionSujetoRetenido = xmlDoc.Descendants("identificacionSujetoRetenido").FirstOrDefault().Value;

            foreach (DataGridViewRow row in dgvCompras.Rows)
            {
                if (row.Cells[9].Value != null && int.Parse(row.Cells[9].Value.ToString()) == numDocSustento && row.Cells[1].Value.ToString().Equals(identificacionSujetoRetenido))
                {
                    row.Cells[24].Value = xmlDoc.Descendants("estab").FirstOrDefault().Value;
                    row.Cells[25].Value = xmlDoc.Descendants("ptoEmi").FirstOrDefault().Value;
                    row.Cells[26].Value = xmlDoc.Descendants("secuencial").FirstOrDefault().Value;
                    row.Cells[27].Value = autorizacion;
                    row.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                    Retencion retencion = new Retencion();
                    DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                    if (dataRow != null)
                    {
                        row.Cells[29].Value = dataRow["concepto"];
                    }
                    row.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                    row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);
                    enc = true;
                    break;
                }
            }
            return enc;
        }

        private bool aplicaCasoDos(string autorizacion, XDocument xmlDoc)
        {
            bool enc = false;
            int numDocSustento = int.Parse(xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value.Substring(6));
            string identificacionSujetoRetenido = xmlDoc.Descendants("identificacionSujetoRetenido").FirstOrDefault().Value;

            foreach (DataGridViewRow row in dgvCompras.Rows)
            {
                //Aqui se subdivide en subcasos
                if (row.Cells[9].Value != null && int.Parse(row.Cells[9].Value.ToString()) == numDocSustento && row.Cells[1].Value.ToString().Equals(identificacionSujetoRetenido))
                {
                    //Estas 4 columas aplican para todos los subcasos
                    row.Cells[24].Value = autorizacion.Substring(24, 3);//Estab
                    row.Cells[25].Value = autorizacion.Substring(27, 3);//PtoEmi
                    row.Cells[26].Value = autorizacion.Substring(30, 9);//Secuencial
                    row.Cells[27].Value = autorizacion;

                    if (double.Parse(row.Cells[12].Value.ToString()) == 0)
                    {
                        row.Cells[25].Style.BackColor = Color.LightGreen;

                        //.................Clonar la fila
                        DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                        for (Int32 index = 0; index < row.Cells.Count; index++)
                        {
                            clonedRow.Cells[index].Value = row.Cells[index].Value;
                        }

                        double b12 = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2);
                        //..................Priemera fila
                        row.Cells[11].Value = 0;//No objeto de IVA. Se pone 0 xq este valor se pone en la siguiente fila
                        //row.Cells[12].Value = 0;//Base 0%. Antes ya se verifico q es 0 entonces se comenta
                        row.Cells[13].Value = b12;

                        row.Cells[16].Value = Math.Round(b12 * .12, 2);//iva
                        row.Cells[17].Value = Math.Round(b12 * .12, 2) + b12;//No le suma el monto ICE xq le asignamos 0

                        row.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                        Retencion retencion = new Retencion();
                        DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                        if (dataRow != null)
                        {
                            row.Cells[29].Value = dataRow["concepto"];
                        }
                        row.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                        row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                        //............Insertar la nueva fila.........................................................
                        //clonedRow.Cells[11].Value = 0; //No es 0 xq antes de clonar tuvo un valor
                        //clonedRow.Cells[12].Value = 0; //Base 0%. Antes ya se verifico q es 0 entonces se comenta
                        b12 = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2);
                        clonedRow.Cells[13].Value = b12;

                        clonedRow.Cells[16].Value = Math.Round(b12 * .12, 2);//IVA
                        clonedRow.Cells[17].Value = Math.Round(b12 * .12, 2) + b12;//Total

                        clonedRow.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                        dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                        if (dataRow != null)
                        {
                            clonedRow.Cells[29].Value = dataRow["concepto"];
                        }
                        clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                        clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                        dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                    }
                    else if (double.Parse(row.Cells[13].Value.ToString()) == 0)
                    {
                        row.Cells[25].Style.BackColor = Color.Green;

                        //.................Clonar la fila
                        DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                        for (Int32 index = 0; index < row.Cells.Count; index++)
                        {
                            clonedRow.Cells[index].Value = row.Cells[index].Value;
                        }

                        //..................Priemera fila..................................................
                        double b0 = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2);
                        row.Cells[11].Value = 0;//Se asigna 0 xq el valor se pone en la siguiente fila o fila duplicada
                        row.Cells[12].Value = b0;
                        //row.Cells[13].Value = 0;//Se establecio en la condicion q es 0 x ende se comenta

                        //row.Cells[16].Value = 0;//El iva debio ser 0 antes xq la base12 es 0
                        row.Cells[17].Value = b0;//El total es igual a la base0 xq los demas son asignados con 0

                        row.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                        Retencion retencion = new Retencion();
                        DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());

                        if (dataRow != null)
                        {
                            row.Cells[29].Value = dataRow["concepto"];
                        }
                        row.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                        row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                        //............Insertar la nueva fila
                        //clonedRow.Cells[11].Value = 0; //No es 0 xq antes de clonar tuvo un valor
                        b0 = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2);
                        clonedRow.Cells[12].Value = b0;
                        //clonedRow.Cells[13].Value = 0;//Se establecio en la condicion q es 0 x ende se comenta

                        //clonedRow.Cells[16].Value = 0;//El iva debio ser 0 antes xq la base12 es 0
                        clonedRow.Cells[17].Value = b0;//El total solo se suma la base0 + el monto ICE xq se sabe q no hay la base12

                        clonedRow.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                        dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                        if (dataRow != null)
                        {
                            clonedRow.Cells[29].Value = dataRow["concepto"];
                        }
                        clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                        clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                        dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                    }
                    //La base 0 y 12 tiene valores diferente de 0
                    else
                    {
                        //Subcaso que primero aplica a la base 0 y despues a la base 12
                        if (Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2) == double.Parse(row.Cells[12].Value.ToString()))
                        {
                            row.Cells[25].Style.BackColor = Color.DarkGreen;
                            //.................Clonar la fila
                            DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                            for (Int32 index = 0; index < row.Cells.Count; index++)
                            {
                                clonedRow.Cells[index].Value = row.Cells[index].Value;
                            }

                            //..................Priemera fila
                            row.Cells[11].Value = 0; //Se le asignara en la siguiente fila
                            //row.Cells[12].Value = 0; //[Punto Clave] Fue sujeto de comparacion que tenia un valor igual a la base de la retencion
                            row.Cells[13].Value = 0;

                            row.Cells[16].Value = 0;//El iva. Su valor se asigna en la siguiente fila
                            row.Cells[17].Value = row.Cells[12].Value;//El total solo la base0

                            row.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                            Retencion retencion = new Retencion();
                            DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                row.Cells[29].Value = dataRow["concepto"];
                            }
                            row.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            //............Insertar la nueva fila
                            //clonedRow.Cells[11].Value = 0; //Se mantiene por que se clona antes de modificar la fila anterior
                            clonedRow.Cells[12].Value = 0;
                            //La siguiente fila, tambien se mantiene xq antes de clonar contenia un valor
                            //clonedRow.Cells[13].Value = xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec);

                            //clonedRow.Cells[16].Value = 0;//El iva. Su valor se guardo en el clone
                            clonedRow.Cells[17].Value = double.Parse(clonedRow.Cells[13].Value.ToString()) + double.Parse(clonedRow.Cells[16].Value.ToString());//El total solo la base0

                            clonedRow.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                            dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                clonedRow.Cells[29].Value = dataRow["concepto"];
                            }
                            clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                        }
                        else if (double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)) == double.Parse(row.Cells[12].Value.ToString()))
                        {
                            row.Cells[25].Style.BackColor = Color.GreenYellow;
                            //.................Clonar la fila
                            DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                            for (Int32 index = 0; index < row.Cells.Count; index++)
                            {
                                clonedRow.Cells[index].Value = row.Cells[index].Value;
                            }

                            //..................Primera fila
                            row.Cells[11].Value = 0;
                            //row.Cells[12].Value = 0; //[Punto Clave] Fue sujeto de comparacion que tenia un valor igual a la base de la retencion
                            row.Cells[13].Value = 0;

                            row.Cells[16].Value = 0;//El iva. Su valor se guardo en el clone
                            row.Cells[17].Value = row.Cells[12].Value;//El total solo la base0

                            row.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                            Retencion retencion = new Retencion();
                            DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                row.Cells[29].Value = dataRow["concepto"];
                            }
                            row.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            //............Insertar la nueva fila
                            //clonedRow.Cells[11].Value = 0; //Se mantiene por que se clona antes de modificar la fila anterior
                            clonedRow.Cells[12].Value = 0;
                            //La siguiente fila, tambien se mantiene xq antes de clonar contenia un valor
                            //clonedRow.Cells[13].Value = xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec);

                            //clonedRow.Cells[16].Value = 0;//El iva. Su valor se guardo en el clone
                            clonedRow.Cells[17].Value = double.Parse(clonedRow.Cells[13].Value.ToString()) + double.Parse(clonedRow.Cells[16].Value.ToString());//El total solo la base0

                            clonedRow.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                            dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                clonedRow.Cells[29].Value = dataRow["concepto"];
                            }
                            clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                        }
                        //Suma retenciones = suma bas 0 + base 12
                        else if (decimal.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec))
                            + decimal.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec))
                            == decimal.Parse(row.Cells[12].Value.ToString()) + decimal.Parse(row.Cells[13].Value.ToString()))
                        {
                            row.Cells[25].Style.BackColor = Color.LightSeaGreen;
                            //.................Clonar la fila
                            DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                            for (Int32 index = 0; index < row.Cells.Count; index++)
                            {
                                clonedRow.Cells[index].Value = row.Cells[index].Value;
                            }

                            //..................Primera fila
                            row.Cells[11].Value = 0;//Base no objeto de IVA
                            row.Cells[12].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec));
                            row.Cells[13].Value = 0;//Base12


                            row.Cells[16].Value = 0; //El iva. Es 0 xq la Base 12 es 0
                            row.Cells[17].Value = row.Cells[12].Value;//El total solo la base0


                            row.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                            Retencion retencion = new Retencion();
                            DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                row.Cells[29].Value = dataRow["concepto"];
                            }
                            row.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            //............Insertar la nueva fila
                            //clonedRow.Cells[11].Value = 0; //Se mantiene por que se clona antes de modificar la fila anterior
                            clonedRow.Cells[12].Value = 0;
                            double b12 = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec));
                            clonedRow.Cells[13].Value = b12;


                            clonedRow.Cells[16].Value = Math.Round(b12 * .12, 2); //El iva. Se aplica con el porcentaje de la base12
                            clonedRow.Cells[17].Value = b12 + Math.Round(b12 * .12, 2);//No se toma en cuenta la base0 xq no es 0 


                            clonedRow.Cells[28].Value = xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                            dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                clonedRow.Cells[29].Value = dataRow["concepto"];
                            }
                            clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants("impuesto").ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                        }
                    }
                    enc = true;
                    break;
                }
            }
            return enc;
        }

        private bool aplicaCasoTres(string autorizacion, XDocument xmlDoc)
        {
            bool enc = false;
            int numDocSustento = int.Parse(xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value.Substring(6));
            string identificacionSujetoRetenido = xmlDoc.Descendants("identificacionSujetoRetenido").FirstOrDefault().Value;

            string comprobante = xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value.Substring(6);
            foreach (DataGridViewRow row in dgvCompras.Rows)
            {
                if (row.Cells[9].Value != null && int.Parse(row.Cells[9].Value.ToString()) == numDocSustento && row.Cells[1].Value.ToString().Equals(identificacionSujetoRetenido))
                {
                    foreach (var impuesto in xmlDoc.Descendants("impuesto"))
                    {
                        switch (int.Parse(impuesto.Descendants("codigo").FirstOrDefault().Value))
                        {
                            case 1: //Retencion a la renta
                                row.Cells[24].Value = autorizacion.Substring(24, 3);    //Estab
                                row.Cells[25].Value = autorizacion.Substring(27, 3);    //PtoEmi
                                row.Cells[26].Value = autorizacion.Substring(30, 9);    //Secuencial
                                row.Cells[27].Value = autorizacion;

                                row.Cells[28].Value = impuesto.Descendants("codigoRetencion").FirstOrDefault().Value;
                                Retencion retencion = new Retencion();
                                DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                                if (dataRow != null)
                                {
                                    row.Cells[29].Value = dataRow["concepto"];
                                }

                                row.Cells[30].Value = double.Parse(impuesto.Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                                row.Cells[31].Value = Math.Round(double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);
                                break;

                            case 2: //Retencion al IVA
                                switch (int.Parse(impuesto.Descendants("codigoRetencion").FirstOrDefault().Value))
                                {
                                    case 1: row.Cells[20].Value = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)).ToString("N2"); break;
                                    case 2: row.Cells[22].Value = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)).ToString("N2"); break;
                                    case 3: row.Cells[23].Value = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)).ToString("N2"); break;
                                    case 9: row.Cells[18].Value = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)).ToString("N2"); break;
                                    case 10: row.Cells[19].Value = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)).ToString("N2"); break;
                                    case 11: row.Cells[21].Value = double.Parse(impuesto.Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)).ToString("N2"); break;
                                }
                                break;
                        }
                    }
                    enc = true;
                    break;
                }
            }
            return enc;
        }

        public void completeRete()
        {
            try
            {
                foreach (DataGridViewRow row in dgvCompras.Rows)
                {
                    if (row.Cells[5].Value.ToString().Equals("F") || row.Cells[5].Value.ToString().Equals("L/C"))
                    {
                        if (row.Cells[28].Value == null || row.Cells[28].Value.ToString().Equals(""))
                        {
                            row.Cells[28].Value = 332;
                            row.Cells[29].Value = "Otras compras de bienes y servicios no sujetas a retención";
                            row.Cells[30].Value = 0;
                            row.Cells[31].Value = 0;
                        }
                        else if (!row.Cells[28].Value.ToString().Contains("332"))
                        {
                            row.Cells[24].Value = row.Cells[24].Value.ToString().Equals("") ? "999" : row.Cells[24].Value.ToString();
                            row.Cells[25].Value = row.Cells[25].Value.ToString().Equals("") ? "999" : row.Cells[25].Value.ToString();
                            row.Cells[26].Value = row.Cells[26].Value.ToString().Equals("") ? "999" : row.Cells[26].Value.ToString();
                            row.Cells[27].Value = row.Cells[27].Value.ToString().Equals("") ? "9999999999" : row.Cells[27].Value.ToString();
                        }
                    }
                    else if (row.Cells[5].Value.ToString().Equals("N/V"))
                    {
                        row.Cells[28].Value = 332;
                        row.Cells[29].Value = "Otras compras de bienes y servicios no sujetas a retención";
                        row.Cells[30].Value = 0;
                        row.Cells[31].Value = 0;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Se produjo un error al completar las retenciones");
            }
        }

        public async void rellenarProveedores()
        {
            bool enc = false;
            List<string> providers = new List<string>();
            foreach (DataGridViewRow row in dgvCompras.Rows)
            {
                if ((row.Cells[2].Value == null || row.Cells[2].Value.Equals("")) && !providers.Contains(row.Cells[1].Value))
                {
                    enc = true;
                    providers.Add(row.Cells[1].Value.ToString());
                }
            }

            if (enc)
            {
                ImportContactos import = new ImportContactos();
                List<Contacto> contactos = await import.getContactMasive(providers);

                if (contactos != null)
                {
                    foreach (DataGridViewRow row in dgvCompras.Rows)
                    {
                        if (row.Cells[2].Value == null || row.Cells[2].Value.Equals(""))
                        {
                            int pos = -1;
                            int i = 0;
                            while (i < contactos.Count && pos == -1)
                            {
                                if (contactos[i].id.Equals(row.Cells[1].Value))
                                {
                                    row.Cells[2].Value = contactos[i].denominacion;
                                    pos = i;
                                }
                                i++;
                            }
                        }
                    }
                }
            }
        }

        public async void rellenarProveedoresMasivo()
        {
            bool enc = false;
            List<string> providers = new List<string>();
            foreach (DataGridViewRow row in dgvCompras.Rows)
            {
                if ((row.Cells[2].Value == null || row.Cells[2].Value.Equals("")) && row.Cells[1].Value.ToString().Length == 13 && !providers.Contains(row.Cells[1].Value))
                {
                    enc = true;
                    providers.Add(row.Cells[1].Value.ToString());
                }
            }

            if (enc)
            {
                ImportContactos import = new ImportContactos();
                List<Contacto> contactos = new List<Contacto>();

                foreach (var idContacto in providers)
                {
                    ResultSriRuc resultSriRuc = new ResultSriRuc();
                    RestClient client = new RestClient();
                    var request = new RestRequest("https://srienlinea.sri.gob.ec/sri-catastro-sujeto-servicio-internet/rest/Persona/obtenerPorTipoIdentificacion", Method.Get);

                    request.AddParameter("numeroIdentificacion", idContacto);
                    request.AddParameter("tipoIdentificacion", "R");
                    var response = client.Execute(request);

                    if (response.IsSuccessful && !response.Content.Equals(""))
                    {
                        var result = response.Content;
                        var resultado = JsonConvert.DeserializeObject<ResultSriRuc>(result);
                        contactos.Add(new Contacto(idContacto, remplazar(resultado.nombreCompleto.Trim())));
                    }
                }

                if (contactos != null)
                {
                    foreach (DataGridViewRow row in dgvCompras.Rows)
                    {
                        if (row.Cells[2].Value == null || row.Cells[2].Value.Equals(""))
                        {
                            int pos = -1;
                            int i = 0;
                            while (i < contactos.Count && pos == -1)
                            {
                                if (contactos[i].id.Equals(row.Cells[1].Value))
                                {
                                    row.Cells[2].Value = contactos[i].denominacion;
                                    pos = i;
                                }
                                i++;
                            }
                        }
                    }
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
            razonsocial = razonsocial.Replace("&", string.Empty);
            razonsocial = razonsocial.Replace(".", string.Empty);
            razonsocial = razonsocial.Replace(",", string.Empty);

            return razonsocial;
        }

        ~DgvCompras() { }
    }
}