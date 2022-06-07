using System;
using System.Data;

namespace WinAppATS
{
    class ConvertionDecimal
    {
        public void convertion(DataTable dataTable)
        {
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                String b0 = dataTable.Rows[i]["b0"].ToString();
                dataTable.Rows[i]["b0"] = b0.Replace(Const.nodec, Const.dec);
                String b12 = dataTable.Rows[i]["b12"].ToString();
                dataTable.Rows[i]["b12"] = b12.Replace(Const.nodec, Const.dec);
                String be = dataTable.Rows[i]["be"].ToString();
                dataTable.Rows[i]["be"] = be.Replace(Const.nodec, Const.dec);
                String mi = dataTable.Rows[i]["mi"].ToString();
                dataTable.Rows[i]["mi"] = mi.Replace(Const.nodec, Const.dec);
                String miv = dataTable.Rows[i]["miv"].ToString();
                dataTable.Rows[i]["miv"] = miv.Replace(Const.nodec, Const.dec);
                String tot = dataTable.Rows[i]["tot"].ToString();
                dataTable.Rows[i]["tot"] = tot.Replace(Const.nodec, Const.dec);
                String r10 = dataTable.Rows[i]["r10"].ToString();
                dataTable.Rows[i]["r10"] = r10.Replace(Const.nodec, Const.dec);
                String r20 = dataTable.Rows[i]["r20"].ToString();
                dataTable.Rows[i]["r20"] = r20.Replace(Const.nodec, Const.dec);
                String r30 = dataTable.Rows[i]["r30"].ToString();
                dataTable.Rows[i]["r30"] = r30.Replace(Const.nodec, Const.dec);
                String r50 = dataTable.Rows[i]["r50"].ToString();
                dataTable.Rows[i]["r50"] = r50.Replace(Const.nodec, Const.dec);
                String r70 = dataTable.Rows[i]["r70"].ToString();
                dataTable.Rows[i]["r70"] = r70.Replace(Const.nodec, Const.dec);
                String r100 = dataTable.Rows[i]["r100"].ToString();
                dataTable.Rows[i]["r100"] = r100.Replace(Const.nodec, Const.dec);
                String vra = dataTable.Rows[i]["vra"].ToString();
                dataTable.Rows[i]["vra"] = vra.Replace(Const.nodec, Const.dec);
            }
        }

        public void conVentas(DataTable dataTable)
        {
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                String bi = dataTable.Rows[i]["bi"].ToString();
                dataTable.Rows[i]["bi"] = bi.Replace(Const.nodec, Const.dec);
                String b0 = dataTable.Rows[i]["b0"].ToString();
                dataTable.Rows[i]["b0"] = b0.Replace(Const.nodec, Const.dec);
                String b12 = dataTable.Rows[i]["b12"].ToString();
                dataTable.Rows[i]["b12"] = b12.Replace(Const.nodec, Const.dec);
                String mi = dataTable.Rows[i]["mi"].ToString();
                dataTable.Rows[i]["mi"] = mi.Replace(Const.nodec, Const.dec);
                String miv = dataTable.Rows[i]["miv"].ToString();
                dataTable.Rows[i]["miv"] = miv.Replace(Const.nodec, Const.dec);
                String tot = dataTable.Rows[i]["tot"].ToString();
                dataTable.Rows[i]["tot"] = tot.Replace(Const.nodec, Const.dec);
                String vrr = dataTable.Rows[i]["vrr"].ToString();
                dataTable.Rows[i]["vrr"] = vrr.Replace(Const.nodec, Const.dec);
                String vri = dataTable.Rows[i]["vri"].ToString();
                dataTable.Rows[i]["vri"] = vri.Replace(Const.nodec, Const.dec);
            }
        }
    }
}
