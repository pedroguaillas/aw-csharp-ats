using System.Data;
using System.Xml;

namespace WinAppATS
{
    class Cliente
    {
        public string getCliente(string ruc)
        {
            string rs = "";
            DataSet data = new DataSet("clientes");
            using (XmlReader xml = XmlReader.Create(Const.filexml("clientes")))
            {
                data.ReadXml(xml);

                DataRow[] results = data.Tables[0].Select(string.Format("ruc = '{0}'", ruc));

                if (results.Length > 0)
                {
                    rs = results[0]["razonsocial"].ToString();
                }
            }
            return rs;
        }

        public DataTable filter(string val)
        {
            DataSet data = new DataSet("clientes");
            using (XmlReader xml = XmlReader.Create(Const.filexml("clientes")))
            {
                data.ReadXml(xml);

                //DataRow[] results = data.Tables[0].Select(string.Format("ruc = '{0}'", val));
                var query = data.Tables[0].AsEnumerable()
                .Where(r => r.Field<string>("ruc").Contains(val) ||
                r.Field<string>("razonsocial").Contains(val)).CopyToDataTable<DataRow>();

                if (query != null)
                {
                    return query;
                }
            }
            return null;
        }

        public DataRow getClienteRow(string ruc)
        {
            DataSet data = new DataSet("clientes");
            using (XmlReader xml = XmlReader.Create(Const.filexml("clientes")))
            {
                data.ReadXml(xml);

                DataRow[] results = data.Tables[0].Select(string.Format("ruc = '{0}'", ruc));

                if (results.Length > 0)
                {
                    return results[0];
                }
            }
            return null;
        }
    }
}
