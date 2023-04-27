using System;
using System.IO;
using System.Net;
using System.Text;
using System.Xml;

namespace WinAppATS
{
    class AutorizacionFacturas
    {
        public bool Descarga(string claveAcceso, string PathServer)
        {
            try
            {
                string result = LlamarSri(claveAcceso);
                if (result != String.Empty)
                {
                    XmlDocument docResponse = new XmlDocument();
                    XmlDocument docSave = new XmlDocument();

                    docResponse.LoadXml(result);
                    if (!Directory.Exists(PathServer))
                    {
                        Directory.CreateDirectory(PathServer);
                    }

                    var autorizacion = docSave.CreateElement("autorizacion");
                    docSave.AppendChild(autorizacion);

                    var estado = docSave.CreateElement("estado");
                    estado.InnerText = docResponse.GetElementsByTagName("estado")[0].InnerText;
                    autorizacion.AppendChild(estado);

                    var numeroAutorizacion = docSave.CreateElement("numeroAutorizacion");
                    numeroAutorizacion.InnerText = docResponse.GetElementsByTagName("numeroAutorizacion")[0].InnerText;
                    autorizacion.AppendChild(numeroAutorizacion);

                    var fechaAutorizacion = docSave.CreateElement("fechaAutorizacion");
                    fechaAutorizacion.InnerText = docResponse.GetElementsByTagName("fechaAutorizacion")[0].InnerText;
                    autorizacion.AppendChild(fechaAutorizacion);

                    var ambiente = docSave.CreateElement("ambiente");
                    ambiente.InnerText = docResponse.GetElementsByTagName("ambiente")[0].InnerText;
                    autorizacion.AppendChild(ambiente);

                    var comprobante = docSave.CreateElement("comprobante");
                    autorizacion.AppendChild(comprobante);

                    var comprobanteCData = docSave.CreateCDataSection("comprobante");
                    string tagComprobante = docResponse.GetElementsByTagName("comprobante")[0].InnerXml.Trim();
                    tagComprobante = tagComprobante.Replace("&lt;", "<");
                    tagComprobante = tagComprobante.Replace("&gt;", ">");
                    comprobanteCData.InnerText = tagComprobante;
                    //comprobanteCData.InnerText = "<?xml version=" + "\"1.0\"" + " encoding=" + "\"UTF-8\"" + "?>" + docResponse.GetElementsByTagName("comprobante")[0].InnerXml.Trim();
                    comprobante.AppendChild(comprobanteCData);

                    XmlDeclaration xmlDeclaration = docSave.CreateXmlDeclaration("1.0", "UTF-8", "yes");
                    XmlElement root = docSave.DocumentElement;
                    docSave.InsertBefore(xmlDeclaration, root);

                    docSave.Save(@PathServer + @"\" + claveAcceso + ".xml");

                    //using (XmlTextWriter wr = new XmlTextWriter(@PathServer + @"\" + claveAcceso + ".xml", Encoding.UTF8))
                    //{
                    //    wr.Formatting = Formatting.None; // here's the trick !
                    //    //wr.Indentation = 0;
                    //    docSave.Save(wr);
                    //}

                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
        /// <summary>
        /// The llamar sri.
        /// </summary>
        /// <param name="claveAcceso">The claveAcceso.</param>
        /// <returns>The <see cref="DatosTributarios"/>.</returns>
        public string LlamarSri(string claveAcceso)
        {
            try
            {
                string result = null;
                string url = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl";

                string xml = "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:ec=\"http://ec.gob.sri.ws.autorizacion\">";
                xml = xml + "<soapenv:Header/>";
                xml = xml + "<soapenv:Body>";
                xml = xml + "<ec:autorizacionComprobante>";
                xml = xml + "<claveAccesoComprobante>" + claveAcceso + "</claveAccesoComprobante>";
                xml = xml + "</ec:autorizacionComprobante>";
                xml = xml + "</soapenv:Body>";
                xml = xml + "</soapenv:Envelope>";

                byte[] bytes = Encoding.ASCII.GetBytes(xml);
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                request.Method = "POST";
                request.ContentLength = bytes.Length;
                request.ContentType = "text/xml";

                Stream requestStream = request.GetRequestStream();
                requestStream.Write(bytes, 0, bytes.Length);
                requestStream.Close();

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream responseStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(responseStream, Encoding.GetEncoding("ISO-8859-1"));
                    result = reader.ReadToEnd();
                }

                response.Close();

                return result;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// The llamar sri.
        /// </summary>
        /// <param name="claveAcceso">The claveAcceso.</param>
        /// <returns>The <see cref="DatosTributarios"/>.</returns>
        public string LlamarSriLote(string claveAcceso)
        {
            //try
            //{
            string result = null;
            string url = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl";

            string xml = "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:ec=\"http://ec.gob.sri.ws.autorizacion\">";
            xml = xml + "<soapenv:Header/>";
            xml = xml + "<soapenv:Body>";
            xml = xml + "<ec:autorizacionComprobanteLote>";
            xml = xml + "<claveAccesoLote>" + claveAcceso + "</claveAccesoLote>";
            xml = xml + "</ec:autorizacionComprobanteLote>";
            xml = xml + "</soapenv:Body>";
            xml = xml + "</soapenv:Envelope>";

            byte[] bytes = Encoding.ASCII.GetBytes(xml);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

            request.Method = "POST";
            request.ContentLength = bytes.Length;
            request.ContentType = "text/xml";

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream, Encoding.GetEncoding("ISO-8859-1"));
                result = reader.ReadToEnd();
            }

            response.Close();

            return result;
            //}
            //catch (Exception)
            //{
            //    return null;
            //}
        }

        ~AutorizacionFacturas() { }
    }
}
