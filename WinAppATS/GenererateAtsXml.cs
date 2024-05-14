using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Text;

namespace WinAppATS
{
    class GenererateAtsXml : Archivo
    {
        String info;
        String establecimiento;
        bool semestral;

        public GenererateAtsXml(string info, string establecimiento, bool semestral)
        {
            this.info = info;
            this.establecimiento = establecimiento;
            this.semestral = semestral;
        }

        public void generate()
        {
            XmlDocument doc = new XmlDocument();
            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);
            XmlElement iva = doc.CreateElement(string.Empty, "iva", string.Empty);
            doc.AppendChild(iva);

            createElement(iva, doc, "TipoIDInformante", "R");
            createElement(iva, doc, "IdInformante", info.Substring(0, 13));

            Cliente cliente = new Cliente();

            createElement(iva, doc, "razonSocial", cliente.getCliente(info.Substring(0, 13)));
            createElement(iva, doc, "Anio", info.Substring(13, 4));
            createElement(iva, doc, "Mes", info.Substring(17));

            if (semestral)
            {
                createElement(iva, doc, "regimenMicroempresa", "SI");
            }

            createElement(iva, doc, "numEstabRuc", establecimiento);
            createElement(iva, doc, "totalVentas", totalVentas());
            createElement(iva, doc, "codigoOperativo", "IVA");

            paraCompras(iva, doc);

            paraVentas(iva, doc);

            paraVentasEstablecimiento(iva, doc);

            paraAnulados(iva, doc);

            string path = miCarpeta("xml", "ATS_" + info.Substring(17));

            if (path != "")
            {
                using (var wr = new XmlTextWriter(path, Encoding.UTF8))
                {
                    wr.Formatting = System.Xml.Formatting.None; // here's the trick !
                    doc.Save(wr);
                }
            }
        }

        private void paraVentas(XmlElement iva, XmlDocument doc)
        {
            if (File.Exists(Const.filexml(@"v" + info)))
            {
                XDocument xmlDoc = XDocument.Load(Const.filexml(@"v" + info));

                XmlElement ventas = doc.CreateElement(string.Empty, "ventas", string.Empty);

                // Ventas agrupadas por tipo de comprobante y cliente
                var groups = from venta in xmlDoc.Root.Elements("venta")
                             group venta by new { TCV = (string)venta.Element("TCV"), tic = (string)venta.Element("tic"), ruc = (string)venta.Element("ruc") } into g
                             select new { g.Key, g };

                foreach (var venta in groups)
                {
                    // Excluye los RUCs en blanco
                    if (venta.Key.ruc != null && venta.Key.ruc != "")
                    {
                        XmlElement detalleVentas = doc.CreateElement(string.Empty, "detalleVentas", string.Empty);
                        ventas.AppendChild(detalleVentas);
                        iva.AppendChild(ventas);

                        createElement(detalleVentas, doc, "tpIdCliente", venta.Key.tic);
                        createElement(detalleVentas, doc, "idCliente", venta.Key.ruc);
                        if (venta.Key.ruc != "9999999999999")
                        {
                            createElement(detalleVentas, doc, "parteRelVtas", "NO");
                        }

                        if (venta.Key.tic == "06")
                        {
                            createElement(detalleVentas, doc, "tipoCliente", "01");
                            createElement(detalleVentas, doc, "denoCli", "PERSONA EXTRANJERA");
                        }
                        createElement(detalleVentas, doc, "tipoComprobante", tcvV(venta.Key.TCV));
                        createElement(detalleVentas, doc, "tipoEmision", ("F"));
                        createElement(detalleVentas, doc, "numeroComprobantes", venta.g.Count().ToString());
                        createElement(detalleVentas, doc, "baseNoGraIva", "0.00");
                        createElement(detalleVentas, doc, "baseImponible", dec(venta.g.Sum(s => decimal.Parse((string)s.Element("b0"))).ToString()));
                        createElement(detalleVentas, doc, "baseImpGrav", dec(venta.g.Sum(s => decimal.Parse((string)s.Element("b12"))).ToString()));
                        createElement(detalleVentas, doc, "montoIva", dec(venta.g.Sum(s => decimal.Parse((string)s.Element("miv"))).ToString()));
                        createElement(detalleVentas, doc, "montoIce", dec(venta.g.Sum(s => decimal.Parse((string)s.Element("mi"))).ToString()));
                        createElement(detalleVentas, doc, "valorRetIva", dec(venta.g.Sum(s => decimal.Parse((string)s.Element("vri"))).ToString()));
                        createElement(detalleVentas, doc, "valorRetRenta", dec(venta.g.Sum(s => decimal.Parse((string)s.Element("vrr"))).ToString()));

                        if (venta.Key.TCV != "N/C")
                        {
                            XmlElement formasDePago = doc.CreateElement(string.Empty, "formasDePago", string.Empty);
                            detalleVentas.AppendChild(formasDePago);
                            createElement(formasDePago, doc, "formaPago", "01");
                        }
                    }
                }
            }
        }

        private void paraAnulados(XmlElement iva, XmlDocument doc)
        {
            if (File.Exists(Const.filexml(@"a" + info)))
            {
                XDocument xmlDoc = XDocument.Load(Const.filexml(@"a" + info));

                XmlElement anulados = doc.CreateElement(string.Empty, "anulados", string.Empty);

                foreach (var anulado in xmlDoc.Descendants("anulado"))
                {
                    XmlElement detalleAnulados = doc.CreateElement(string.Empty, "detalleAnulados", string.Empty);
                    anulados.AppendChild(detalleAnulados);
                    iva.AppendChild(anulados);

                    createElement(detalleAnulados, doc, "tipoComprobante", _(anulado, "tipoComprobante"));
                    createElement(detalleAnulados, doc, "establecimiento", _(anulado, "establecimiento"));
                    createElement(detalleAnulados, doc, "puntoEmision", _(anulado, "puntoEmision"));
                    createElement(detalleAnulados, doc, "secuencialInicio", _(anulado, "secuencialInicio"));
                    createElement(detalleAnulados, doc, "secuencialFin", _(anulado, "secuencialFin"));
                    createElement(detalleAnulados, doc, "autorizacion", _(anulado, "autorizacion"));
                }
            }
        }

        private void paraVentasEstablecimiento(XmlElement iva, XmlDocument doc)
        {
            if (File.Exists(Const.filexml(@"v" + info)))
            {
                XDocument xmlDoc = XDocument.Load(Const.filexml(@"v" + info));

                XmlElement ventasEstablecimiento = doc.CreateElement(string.Empty, "ventasEstablecimiento", string.Empty);

                //Agrupar por establecimiento
                var query = xmlDoc.Root.Elements("venta")
                    .Where(w => w.Element("com") != null && w.Element("com").Value != "" && w.Element("com").Value.Substring(0, 3) != "000" && w.Element("com").Value.ToString().Length > 2)
                     .GroupBy(g => g.Element("com").Value.Substring(0, 3))
                     .Select(s => new
                     {
                         est = s.Descendants("com").FirstOrDefault().Value.Substring(0, 3),
                         bi = s.Sum(su => (decimal.Parse(su.Descendants("b0").FirstOrDefault().Value) + decimal.Parse(su.Descendants("b12").FirstOrDefault().Value)) * (_(su, "TCV") == "N/C" ? -1 : 1))
                     })
                     .ToList();

                for (int i = 1; i <= int.Parse(establecimiento); i++)
                {
                    XmlElement ventaEst = doc.CreateElement(string.Empty, "ventaEst", string.Empty);
                    ventasEstablecimiento.AppendChild(ventaEst);
                    iva.AppendChild(ventasEstablecimiento);

                    var search = query.FirstOrDefault(q => i == int.Parse(q.est));

                    if (search != null)
                    {
                        createElement(ventaEst, doc, "codEstab", search.est);
                        createElement(ventaEst, doc, "ventasEstab", dec(search.bi.ToString()));
                    }
                    else
                    {
                        createElement(ventaEst, doc, "codEstab", i.ToString("000"));
                        createElement(ventaEst, doc, "ventasEstab", "0.00");
                    }
                    createElement(ventaEst, doc, "ivaComp", "0.00");
                }
            }
        }

        private string totalVentas()
        {
            if (File.Exists(Const.filexml(@"v" + info)))
            {
                XDocument xmlDoc = XDocument.Load(Const.filexml(@"v" + info));

                decimal sum = 0;

                foreach (var venta in xmlDoc.Descendants("venta"))
                {
                    decimal b0 = decimal.Parse(_(venta, ("b0")));
                    decimal b12 = decimal.Parse(_(venta, ("b12")));

                    sum += (b0 + b12) * (_(venta, "TCV") == "N/C" ? -1 : 1);
                }

                return dec(sum.ToString());
            }

            return "0.00";
        }

        private void paraCompras(XmlElement iva, XmlDocument doc)
        {
            if (File.Exists(Const.filexml(@"c" + info)))
            {
                XDocument xmlDoc = XDocument.Load(Const.filexml(@"c" + info));

                XmlElement compras = doc.CreateElement(string.Empty, "compras", string.Empty);

                // Compras agrupadas por tipo de comprobante y cliente
                var groups = from compra in xmlDoc.Root.Elements("compra")
                             group compra by new { sec = (string)compra.Element("sec"), pe = (string)compra.Element("pe"), Est = (string)compra.Element("Est"), TCV = (string)compra.Element("TCV"), RUC = (string)compra.Element("RUC") } into g
                             select new { g.Key, g };

                foreach (var comp in groups)
                {
                    XmlElement detalleCompras = doc.CreateElement(string.Empty, "detalleCompras", string.Empty);
                    compras.AppendChild(detalleCompras);
                    iva.AppendChild(compras);

                    var compra = comp.g.FirstOrDefault();

                    createElement(detalleCompras, doc, "codSustento", _(compra, "TCV") == "N/V" ? "02" : codSustento(_(compra, "ccu")));
                    createElement(detalleCompras, doc, "tpIdProv", (comp.Key.RUC.Length == 13 ? "01" : (comp.Key.RUC.Length == 10 ? "02" : "03")));
                    createElement(detalleCompras, doc, "idProv", comp.Key.RUC);
                    createElement(detalleCompras, doc, "tipoComprobante", tcv(comp.Key.TCV));

                    createElement(detalleCompras, doc, "tipoProv", (comp.Key.RUC.Length == 13 && (comp.Key.RUC.Substring(2, 1) == "9" || comp.Key.RUC.Substring(2, 1) == "6")) ? "02" : "01");
                    createElement(detalleCompras, doc, "denoProv", _(compra, ("rs")));
                    createElement(detalleCompras, doc, "parteRel", "NO");
                    createElement(detalleCompras, doc, "fechaRegistro", _(compra, ("fec")));
                    createElement(detalleCompras, doc, "establecimiento", comp.Key.Est);
                    createElement(detalleCompras, doc, "puntoEmision", comp.Key.pe);
                    //createElement(detalleCompras, doc, "establecimiento", _(compra, ("Est")));
                    //createElement(detalleCompras, doc, "puntoEmision", _(compra, ("pe")));
                    createElement(detalleCompras, doc, "secuencial", _(compra, ("sec")));
                    createElement(detalleCompras, doc, "fechaEmision", _(compra, ("fec")));
                    createElement(detalleCompras, doc, "autorizacion", _(compra, ("aut")));

                    // Las agrupaciones de valores sumar, compras que tienen retenciones desde dos retenciones en Renta

                    createElement(detalleCompras, doc, "baseNoGraIva", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("bni"))).ToString()));
                    createElement(detalleCompras, doc, "baseImponible", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("b0"))).ToString()));
                    createElement(detalleCompras, doc, "baseImpGrav", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("b12")) + decimal.Parse((string)s.Element("be"))).ToString()));
                    createElement(detalleCompras, doc, "baseImpExe", "0.00");
                    createElement(detalleCompras, doc, "montoIce", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("mi"))).ToString()));
                    createElement(detalleCompras, doc, "montoIva", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("miv"))).ToString()));
                    createElement(detalleCompras, doc, "valRetBien10", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("r10"))).ToString()));
                    createElement(detalleCompras, doc, "valRetServ20", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("r20"))).ToString()));
                    createElement(detalleCompras, doc, "valorRetBienes", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("r30"))).ToString()));
                    createElement(detalleCompras, doc, "valRetServ50", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("r50"))).ToString()));
                    createElement(detalleCompras, doc, "valorRetServicios", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("r70"))).ToString()));
                    createElement(detalleCompras, doc, "valRetServ100", dec(comp.g.Sum(s => decimal.Parse((string)s.Element("r100"))).ToString()));

                    createElement(detalleCompras, doc, "valorRetencionNc", "0");
                    createElement(detalleCompras, doc, "totbasesImpReemb", "0");

                    pagoExt(detalleCompras, doc);

                    //Formas de pago
                    XmlElement formasDePago = doc.CreateElement(string.Empty, "formasDePago", string.Empty);
                    detalleCompras.AppendChild(formasDePago);

                    foreach (var meth in comp.g)
                    {
                        createElement(formasDePago, doc, "formaPago", (double.Parse((string)meth.Element("bni")) + double.Parse((string)meth.Element("b0")) + double.Parse((string)meth.Element("b12")) > 999.99) ? "20" : "01");
                    }

                    if (_(compra, ("cda")) != null && _(compra, ("cda")) != "")
                    {
                        //Retenciones en compra
                        XmlElement air = doc.CreateElement(string.Empty, "air", string.Empty);
                        detalleCompras.AppendChild(air);

                        foreach (var meth in comp.g)
                        {
                            XmlElement detalleAir = doc.CreateElement(string.Empty, "detalleAir", string.Empty);
                            air.AppendChild(detalleAir);

                            createElement(detalleAir, doc, "codRetAir", _(meth, ("cda")));
                            createElement(detalleAir, doc, "baseImpAir", dec((decimal.Parse((string)meth.Element("b0")) + decimal.Parse((string)meth.Element("b12")) + decimal.Parse((string)meth.Element("be")) - decimal.Parse((string)meth.Element("mi"))).ToString()));
                            createElement(detalleAir, doc, "porcentajeAir", dec(_(meth, ("por"))));
                            createElement(detalleAir, doc, "valRetAir", dec(_(meth, ("vra"))));
                        }

                        //Si las retencionoes no son 332 deben tener mas info
                        if (_(compra, ("es1")) != "")
                        {
                            createElement(detalleCompras, doc, "estabRetencion1", _(compra, ("es1")));
                            createElement(detalleCompras, doc, "ptoEmiRetencion1", _(compra, ("pe1")));
                            createElement(detalleCompras, doc, "secRetencion1", _(compra, ("se1")));
                            createElement(detalleCompras, doc, "autRetencion1", _(compra, ("au1")));
                            createElement(detalleCompras, doc, "fechaEmiRet1", _(compra, ("fec")));
                        }
                    }
                    // NC y ND
                    if (int.Parse(tcv(_(compra, ("TCV")))) > 3)
                    {
                        createElement(detalleCompras, doc, "docModificado", ("01"));
                        createElement(detalleCompras, doc, "estabModificado", _(compra, ("em")));
                        createElement(detalleCompras, doc, "ptoEmiModificado", _(compra, ("pem")));
                        createElement(detalleCompras, doc, "secModificado", _(compra, ("sm")));
                        createElement(detalleCompras, doc, "autModificado", _(compra, ("aut")));
                    }
                }
            }
        }

        private void pagoExt(XmlElement element, XmlDocument doc)
        {
            XmlElement pagoExterior = doc.CreateElement(string.Empty, "pagoExterior", string.Empty);
            element.AppendChild(pagoExterior);

            createElement(pagoExterior, doc, "pagoLocExt", "01");
            createElement(pagoExterior, doc, "paisEfecPago", "NA");
            createElement(pagoExterior, doc, "aplicConvDobTrib", "NA");
            createElement(pagoExterior, doc, "pagExtSujRetNorLeg", "NA");
        }

        private string dec(string num)
        {
            if (num == "")
            {
                return "0.00";
            }

            decimal val = decimal.Parse(num);
            return val.ToString("0.00").Replace(',', '.');
        }

        private string tcvV(string t)
        {
            switch (t)
            {
                case "F": return "18";
                case "N/V": return "02";
                case "N/C": return "04";
                default: return "";
            }
        }

        private string tcv(string t)
        {
            switch (t)
            {
                case "F": return "01";
                case "N/V": return "02";
                case "L/C": return "03";
                case "N/C": return "04";
                case "N/D": return "05";
                default: return "0";
            }
        }
        private string codSustento(string cuenta)
        {
            if (cuenta.Length > 8)
            {
                Count count = new Count();
                return count.getCodSustento(cuenta);
            }
            return "";
        }

        private decimal _d(XElement element, string name)
        {
            decimal n;

            if (element.Element(name) != null && decimal.TryParse(element.Element(name).Value, out n))
            {
                return decimal.Parse(element.Element(name).Value);
            }

            return 0;
        }

        private string _(XElement element, string name)
        {
            if (element.Element(name) != null)
            {
                return element.Element(name).Value;
            }

            return "";
        }

        private void createElement(XmlElement xmlElement, XmlDocument doc, string name, string value)
        {
            XmlElement element = doc.CreateElement(string.Empty, name, string.Empty);
            XmlText textElement = doc.CreateTextNode(value);
            element.AppendChild(textElement);
            xmlElement.AppendChild(element);
        }
    }
}
