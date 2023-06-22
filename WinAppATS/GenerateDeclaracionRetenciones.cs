using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace WinAppATS
{
    class GenerateDeclaracionRetenciones : Archivo
    {
        String info;
        JObject jsoncontent;
        JObject jObject;

        public GenerateDeclaracionRetenciones(string info)
        {
            this.info = info;
            jsoncontent = new JObject();
            jObject = new JObject();
            jsoncontent["detallesDeclaracion"] = jObject;
        }
        public void generate()
        {
            compras();

            string path = miCarpeta("json", "Retenciones_" + info.Substring(17));

            if (path != "")
            {
                // write JSON directly to a file
                using (StreamWriter file = File.CreateText(path))
                using (JsonTextWriter writer = new JsonTextWriter(file))
                {
                    jsoncontent.WriteTo(writer);
                }
            }
        }

        private void compras()
        {
            if (File.Exists(Const.filexml(@"c" + info)))
            {
                XDocument xmlDoc = XDocument.Load(Const.filexml(@"c" + info));

                // Compras agrupadas por codigo de retencion
                var groups = from compra in xmlDoc.Root.Elements("compra")
                             where compra.Element("cda") != null && !compra.Element("cda").Equals("")
                             group compra by new { cda = (string)compra.Element("cda") } into g
                             select new { g.Key, g };

                foreach (var comp in groups)
                {
                    var compra = comp.g.FirstOrDefault();
                    string cda = _(compra, ("cda"));

                    if (!cda.Equals(""))
                    {
                        Retencion retencion = new Retencion();
                        DataRow row = retencion.completeRetenciones(cda);

                        decimal ret = comp.g.Sum(s => decimal.Parse((string)s.Element("b0")) + decimal.Parse((string)s.Element("b12")) - decimal.Parse((string)s.Element("mi")));

                        if (jObject.ContainsKey((string)row[3]))
                        {
                            ret += decimal.Parse(jObject[row[3]].ToString().Replace('.', ','));
                        }

                        jObject[row != null ? row[3] : ""] = dec(ret.ToString());
                    }
                }
            }
        }

        private string _(XElement element, string name)
        {
            if (element.Element(name) != null)
            {
                return element.Element(name).Value;
            }

            return "";
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

        ~GenerateDeclaracionRetenciones() { }
    }
}
