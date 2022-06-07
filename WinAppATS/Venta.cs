using System;

namespace WinAppATS
{
    class Venta
    {
        public String idCliente { set; get; }
        public String fecha { set; get; }
        public String comprobante { set; get; }
        public String tipoComprobante { set; get; }
        public String tipoEmision { set; get; }
        public String baseNoGraIva { set; get; }
        public String baseImponible { set; get; }
        public String baseImpGrav { set; get; }
        public String montoIva { set; get; }
        public String montoIce { set; get; }
        public String valorRetIva { set; get; }
        public String valorRetRenta { set; get; }
    }
}