using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;

namespace WinAppATS
{
    class Count
    {

        string filename = string.Format("{0}//data.dat", System.Windows.Forms.Application.StartupPath);

        public string completeCounts(string code)
        {
            DataSet data = new DataSet("cuentas");
            using (XmlReader xml = XmlReader.Create(filename))
            {
                data.ReadXml(xml);

                DataRow[] results = data.Tables[0].Select(string.Format("code = '{0}'", code));

                if (results.Length > 0)
                {
                    return results[0]["description"].ToString();
                }
            }
            return "";
        }

        public string getCodSustento(string code)
        {
            DataSet data = new DataSet("cuentas");
            using (XmlReader xml = XmlReader.Create(filename))
            {
                data.ReadXml(xml);

                DataRow[] results = data.Tables[0].Select(string.Format("code = '{0}'", code));

                if (results.Length > 0)
                {
                    return results[0]["tst"].ToString();
                }
            }
            return "";
        }

        public List<String> getColecs()
        {

            DataSet data = new DataSet("cuentas");
            using (XmlReader xml = XmlReader.Create(filename))
            {
                data.ReadXml(xml);

                List<String> list = (from row in data.Tables[0].AsEnumerable()
                                     select row.Field<String>("code")).ToList<String>();
                return list;
            }

            return null;
        }

        ~Count() { }
    }
}