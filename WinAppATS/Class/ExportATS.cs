using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinAppATS.Class
{
    class ExportATS
    {
        public void export(string info)
        {
            //Comprobar si existe los archivos

            //Si no existe archivo de compras dejar en blanco
            if (File.Exists(Const.filexml("c" + info)))
            {

            }

            //Si no existe archivo de ventas generar en 0
            if (File.Exists(Const.filexml("v" + info)))
            {

            }
            //Si no existe archivo de anulados dejar en blanco
            if (File.Exists(Const.filexml("a" + info)))
            {

            }
        }
    }
}
