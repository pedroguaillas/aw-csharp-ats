using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;

namespace WinAppATS
{
    class Retencion
    {

        string filename = string.Format("{0}//rentencion.dat", System.Windows.Forms.Application.StartupPath);

        public DataRow completeRetenciones(string code)
        {
            DataSet data = new DataSet("retenciones");
            using (XmlReader xml = XmlReader.Create(filename))
            {
                data.ReadXml(xml);

                DataRow[] results = data.Tables[0].Select(string.Format("code = '{0}'", code));

                if (results.Length > 0)
                {
                    return results[0];
                }
            }

            return null;
        }

        public List<String> getColecs()
        {
            DataSet data = new DataSet("retenciones");
            using (XmlReader xml = XmlReader.Create(filename))
            {
                data.ReadXml(xml);

                List<String> list = (from row in data.Tables[0].AsEnumerable()
                                     select row.Field<String>("code")).ToList<String>();
                return list;
            }

            return null;
        }

        ~Retencion() { }
    }
}