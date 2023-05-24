using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Xml.Linq;

namespace WinAppATS
{
    class GenerateDeclaracionIVA : Archivo
    {
        String info;
        JObject jsoncontent;
        JObject jObject;

        public GenerateDeclaracionIVA(string info)
        {
            this.info = info;
            jsoncontent = new JObject();
            jObject = new JObject();
            jsoncontent["detallesDeclaracion"] = jObject;
        }
        public void generate()
        {
            ventas();
            compras();

            string path = miCarpeta("json", "IVA_" + info.Substring(17));

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

        private void ventas()
        {
            if (File.Exists(Const.filexml(@"v" + info)))
            {
                XDocument xmlDoc = XDocument.Load(Const.filexml(@"v" + info));
                decimal sumfact12 = 0;
                decimal sumfact0 = 0;
                decimal sumb12 = 0;
                decimal sumb0 = 0;
                decimal sumbRetIva = 0;

                foreach (var venta in xmlDoc.Descendants("venta"))
                {
                    if (_(venta, "TCV") == "F")
                    {
                        sumfact12 += decimal.Parse(_(venta, ("b12")));
                        sumfact0 += decimal.Parse(_(venta, ("b0")));
                    }

                    sumb12 += decimal.Parse(_(venta, ("b12"))) * (_(venta, "TCV") == "N/C" ? -1 : 1);
                    sumb0 += decimal.Parse(_(venta, ("b0"))) * (_(venta, "TCV") == "N/C" ? -1 : 1);
                    sumbRetIva += decimal.Parse(_(venta, ("vri")));
                }

                if (sumfact12 > 0)
                {
                    jObject["450"] = dec(sumfact12.ToString());
                }
                if (sumfact0 > 0)
                {
                    jObject["570"] = dec(sumfact0.ToString());
                }
                if (sumb12 > 0)
                {
                    jObject["460"] = dec(sumb12.ToString());
                }
                //jObject["670"] = sumfact0.ToString();
                //jObject["680"] = sumb0.ToString();
                if (sumb0 > 0)
                {
                    jObject["580"] = dec(sumb0.ToString());
                }
                if (sumbRetIva > 0)
                {
                    jObject["2200"] = dec(sumbRetIva.ToString());
                }
            }
        }

        private void compras()
        {
            if (File.Exists(Const.filexml(@"c" + info)))
            {
                XDocument xmlDoc = XDocument.Load(Const.filexml(@"c" + info));
                decimal sumfact12 = 0;
                decimal sumfact0 = 0;
                decimal sumb12 = 0;
                decimal sumb0 = 0;
                decimal sumbNV = 0;
                decimal sumbNoIva = 0;

                //retenciones
                decimal r10 = 0;
                decimal r20 = 0;
                decimal r30 = 0;
                decimal r50 = 0;
                decimal r70 = 0;
                decimal r100 = 0;

                foreach (var venta in xmlDoc.Descendants("compra"))
                {
                    string tcv = _(venta, "TCV");

                    if (tcv == "F" || tcv == "L/C" || tcv == "N/D")
                    {
                        sumfact12 += decimal.Parse(_(venta, ("b12")));
                        sumfact0 += decimal.Parse(_(venta, ("b0")));
                    }

                    if (tcv == "F")
                    {
                        r10 += decimal.Parse(_(venta, ("r10")));
                        r20 += decimal.Parse(_(venta, ("r20")));
                        r30 += decimal.Parse(_(venta, ("r30")));
                        r50 += decimal.Parse(_(venta, ("r50")));
                        r70 += decimal.Parse(_(venta, ("r70")));
                        r100 += decimal.Parse(_(venta, ("r100")));
                    }

                    sumb12 += decimal.Parse(_(venta, ("b12"))) * (tcv == "N/C" ? -1 : 1);

                    sumbNoIva += (tcv == "F" ? decimal.Parse(_(venta, ("bni"))) : 0);

                    if (tcv == "N/V")
                    {
                        sumbNV += decimal.Parse(_(venta, ("b0")));
                    }
                    else
                    {
                        sumb0 += decimal.Parse(_(venta, ("b0"))) * (tcv == "N/C" ? -1 : 1);
                    }
                }

                if (sumfact12 > 0)
                    jObject["1270"] = dec(sumfact12.ToString());
                if (sumb12 > 0)
                    jObject["1280"] = dec(sumb12.ToString());
                if (sumfact0 > 0)
                    jObject["1720"] = dec(sumfact0.ToString());
                if (sumb0 > 0)
                    jObject["1730"] = dec(sumb0.ToString());
                if (sumbNV > 0)
                {
                    jObject["1735"] = dec(sumbNV.ToString());
                    jObject["1740"] = dec(sumbNV.ToString());
                }
                if (sumbNoIva > 0)
                {
                    jObject["1818"] = dec(sumbNoIva.ToString());
                    jObject["1820"] = dec(sumbNoIva.ToString());
                }

                //Retenciones
                if (r10 > 0)
                    jObject["2515"] = dec(r10.ToString());
                if (r20 > 0)
                    jObject["2525"] = dec(r20.ToString());
                if (r30 > 0)
                    jObject["2520"] = dec(r30.ToString());
                if (r50 > 0)
                    jObject["2960"] = dec(r50.ToString());
                if (r70 > 0)
                    jObject["2530"] = dec(r70.ToString());
                if (r100 > 0)
                    jObject["2540"] = dec(r100.ToString());
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

        ~GenerateDeclaracionIVA() { }
    }
}
