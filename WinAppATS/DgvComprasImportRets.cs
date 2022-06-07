using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace WinAppATS
{
    class DgvComprasImportRets : Importa
    {
        DataGridView dgvCompras;
        string info;
        NumberFormatInfo nfi;
        char dec;

        public DgvComprasImportRets(DataGridView dgvCompras)
        {
            this.dgvCompras = dgvCompras;
            dec = ',';
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

                            //Determina que el comprobante sea una retencion
                            if (xmlDoc.Descendants("codDoc").FirstOrDefault().Value == "07")
                            {
                                bool import = false;
                                string taxes = "impuestos";
                                string tax = "impuesto";

                                if (xmlDoc.Root.Attribute("version").Value.Trim().Equals("2.0.0"))
                                {
                                    taxes = "retenciones";
                                    tax = "retencion";
                                }
                                //Determina el numero de reteciones que se aplica a la factura
                                switch (xmlDoc.Descendants(taxes).Elements().Count())
                                {
                                    case 1:
                                        //Caso 1. Validar que la retencion sea de Renta.
                                        if (int.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigo").FirstOrDefault().Value) == 1)
                                        {
                                            import = aplicaCasoUno(XTemp.Descendants("numeroAutorizacion").FirstOrDefault().Value, xmlDoc, tax);
                                        }
                                        break;
                                    case 2:
                                        //Caso 2. Tiene dos retenciones a la renta
                                        if (xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigo").FirstOrDefault().Value == xmlDoc.Descendants(tax).ElementAt(1).Descendants("codigo").FirstOrDefault().Value)
                                        {
                                            //Validar que la retencion sea de Renta.
                                            if (int.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigo").FirstOrDefault().Value) == 1)
                                            {
                                                import = aplicaCasoDos(XTemp.Descendants("numeroAutorizacion").FirstOrDefault().Value, xmlDoc, tax);
                                            }
                                        }
                                        //Caso 3. Tiene una rentencion a la renta y uno al IVA
                                        else
                                        {
                                            import = aplicaCasoTres(XTemp.Descendants("numeroAutorizacion").FirstOrDefault().Value, xmlDoc, tax);
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
        private bool aplicaCasoUno(string autorizacion, XDocument xmlDoc, string tax)
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
                    row.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                    Retencion retencion = new Retencion();
                    DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                    if (dataRow != null)
                    {
                        row.Cells[29].Value = dataRow["concepto"];
                    }
                    row.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                    row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);
                    enc = true;
                    break;
                }
            }
            return enc;
        }

        private bool aplicaCasoDos(string autorizacion, XDocument xmlDoc, string tax)
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

                        double b12 = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2);
                        //..................Priemera fila
                        row.Cells[11].Value = 0;//No objeto de IVA. Se pone 0 xq este valor se pone en la siguiente fila
                        //row.Cells[12].Value = 0;//Base 0%. Antes ya se verifico q es 0 entonces se comenta
                        row.Cells[13].Value = b12;

                        row.Cells[16].Value = Math.Round(b12 * .12, 2);//iva
                        row.Cells[17].Value = Math.Round(b12 * .12, 2) + b12;//No le suma el monto ICE xq le asignamos 0

                        row.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                        Retencion retencion = new Retencion();
                        DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                        if (dataRow != null)
                        {
                            row.Cells[29].Value = dataRow["concepto"];
                        }
                        row.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                        row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                        //............Insertar la nueva fila.........................................................
                        //clonedRow.Cells[11].Value = 0; //No es 0 xq antes de clonar tuvo un valor
                        //clonedRow.Cells[12].Value = 0; //Base 0%. Antes ya se verifico q es 0 entonces se comenta
                        b12 = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2);
                        clonedRow.Cells[13].Value = b12;

                        clonedRow.Cells[16].Value = Math.Round(b12 * .12, 2);//IVA
                        clonedRow.Cells[17].Value = Math.Round(b12 * .12, 2) + b12;//Total

                        clonedRow.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                        dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                        if (dataRow != null)
                        {
                            clonedRow.Cells[29].Value = dataRow["concepto"];
                        }
                        clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                        clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

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
                        double b0 = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2);
                        row.Cells[11].Value = 0;//Se asigna 0 xq el valor se pone en la siguiente fila o fila duplicada
                        row.Cells[12].Value = b0;
                        //row.Cells[13].Value = 0;//Se establecio en la condicion q es 0 x ende se comenta

                        //row.Cells[16].Value = 0;//El iva debio ser 0 antes xq la base12 es 0
                        row.Cells[17].Value = b0;//El total es igual a la base0 xq los demas son asignados con 0

                        row.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                        Retencion retencion = new Retencion();
                        DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());

                        if (dataRow != null)
                        {
                            row.Cells[29].Value = dataRow["concepto"];
                        }
                        row.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                        row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                        //............Insertar la nueva fila
                        //clonedRow.Cells[11].Value = 0; //No es 0 xq antes de clonar tuvo un valor
                        b0 = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2);
                        clonedRow.Cells[12].Value = b0;
                        //clonedRow.Cells[13].Value = 0;//Se establecio en la condicion q es 0 x ende se comenta

                        //clonedRow.Cells[16].Value = 0;//El iva debio ser 0 antes xq la base12 es 0
                        clonedRow.Cells[17].Value = b0;//El total solo se suma la base0 + el monto ICE xq se sabe q no hay la base12

                        clonedRow.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                        dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                        if (dataRow != null)
                        {
                            clonedRow.Cells[29].Value = dataRow["concepto"];
                        }
                        clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                        clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                        dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                    }
                    //La base 0 y 12 tiene valores diferente de 0
                    else
                    {
                        //Subcaso que primero aplica a la base 0 y despues a la base 12
                        if (Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)), 2) == double.Parse(row.Cells[12].Value.ToString()))
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

                            row.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                            Retencion retencion = new Retencion();
                            DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                row.Cells[29].Value = dataRow["concepto"];
                            }
                            row.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            //............Insertar la nueva fila
                            //clonedRow.Cells[11].Value = 0; //Se mantiene por que se clona antes de modificar la fila anterior
                            clonedRow.Cells[12].Value = 0;
                            //La siguiente fila, tambien se mantiene xq antes de clonar contenia un valor
                            //clonedRow.Cells[13].Value = xmlDoc.Descendants(tax).ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec);

                            //clonedRow.Cells[16].Value = 0;//El iva. Su valor se guardo en el clone
                            clonedRow.Cells[17].Value = double.Parse(clonedRow.Cells[13].Value.ToString()) + double.Parse(clonedRow.Cells[16].Value.ToString());//El total solo la base0

                            clonedRow.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                            dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                clonedRow.Cells[29].Value = dataRow["concepto"];
                            }
                            clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                        }
                        else if (double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec)) == double.Parse(row.Cells[12].Value.ToString()))
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

                            row.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                            Retencion retencion = new Retencion();
                            DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                row.Cells[29].Value = dataRow["concepto"];
                            }
                            row.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            //............Insertar la nueva fila
                            //clonedRow.Cells[11].Value = 0; //Se mantiene por que se clona antes de modificar la fila anterior
                            clonedRow.Cells[12].Value = 0;
                            //La siguiente fila, tambien se mantiene xq antes de clonar contenia un valor
                            //clonedRow.Cells[13].Value = xmlDoc.Descendants(tax).ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec);

                            //clonedRow.Cells[16].Value = 0;//El iva. Su valor se guardo en el clone
                            clonedRow.Cells[17].Value = double.Parse(clonedRow.Cells[13].Value.ToString()) + double.Parse(clonedRow.Cells[16].Value.ToString());//El total solo la base0

                            clonedRow.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                            dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                clonedRow.Cells[29].Value = dataRow["concepto"];
                            }
                            clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                        }
                        //Suma retenciones = suma bas 0 + base 12
                        else if (decimal.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec))
                            + decimal.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec))
                            == decimal.Parse(row.Cells[12].Value.ToString()) + decimal.Parse(row.Cells[13].Value.ToString()))
                        {
                            row.Cells[25].Style.BackColor = Color.LightSeaGreen;
                            //.................Clonar la fila
                            DataGridViewRow clonedRow = (DataGridViewRow)row.Clone();
                            for (Int32 index = 0; index < row.Cells.Count; index++)
                            {
                                clonedRow.Cells[index].Value = row.Cells[index].Value;
                            }

                            //..................Priemera fila
                            row.Cells[11].Value = 0;//Base no objeto de IVA
                            row.Cells[12].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec));
                            row.Cells[13].Value = 0;//Base12


                            row.Cells[16].Value = 0; //El iva. Es 0 xq la Base 12 es 0
                            row.Cells[17].Value = row.Cells[12].Value;//El total solo la base0


                            row.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(0).Descendants("codigoRetencion").FirstOrDefault().Value;
                            Retencion retencion = new Retencion();
                            DataRow dataRow = retencion.completeRetenciones(row.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                row.Cells[29].Value = dataRow["concepto"];
                            }
                            row.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            row.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(0).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            //............Insertar la nueva fila
                            //clonedRow.Cells[11].Value = 0; //Se mantiene por que se clona antes de modificar la fila anterior
                            clonedRow.Cells[12].Value = 0;
                            double b12 = double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("baseImponible").FirstOrDefault().Value.Replace('.', dec));
                            clonedRow.Cells[13].Value = b12;


                            clonedRow.Cells[16].Value = Math.Round(b12 * .12, 2); //El iva. Se aplica con el porcentaje de la base12
                            clonedRow.Cells[17].Value = b12 + Math.Round(b12 * .12, 2);//No se toma en cuenta la base0 xq no es 0 


                            clonedRow.Cells[28].Value = xmlDoc.Descendants(tax).ElementAt(1).Descendants("codigoRetencion").FirstOrDefault().Value;
                            dataRow = retencion.completeRetenciones(clonedRow.Cells[28].Value.ToString());
                            if (dataRow != null)
                            {
                                clonedRow.Cells[29].Value = dataRow["concepto"];
                            }
                            clonedRow.Cells[30].Value = double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("porcentajeRetener").FirstOrDefault().Value.Replace('.', dec));
                            clonedRow.Cells[31].Value = Math.Round(double.Parse(xmlDoc.Descendants(tax).ElementAt(1).Descendants("valorRetenido").FirstOrDefault().Value.Replace('.', dec)), 2);

                            dgvCompras.Rows.Insert(row.Index + 1, clonedRow);
                        }
                    }
                    enc = true;
                    break;
                }
            }
            return enc;
        }

        private bool aplicaCasoTres(string autorizacion, XDocument xmlDoc, string tax)
        {
            bool enc = false;
            int numDocSustento = int.Parse(xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value.Substring(6));
            string identificacionSujetoRetenido = xmlDoc.Descendants("identificacionSujetoRetenido").FirstOrDefault().Value;

            string comprobante = xmlDoc.Descendants("numDocSustento").FirstOrDefault().Value.Substring(6);
            foreach (DataGridViewRow row in dgvCompras.Rows)
            {
                if (row.Cells[9].Value != null && int.Parse(row.Cells[9].Value.ToString()) == numDocSustento && row.Cells[1].Value.ToString().Equals(identificacionSujetoRetenido))
                {
                    foreach (var impuesto in xmlDoc.Descendants(tax))
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

        ~DgvComprasImportRets() { }
    }
}