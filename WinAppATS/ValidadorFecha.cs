using System;
using System.Windows.Forms;

namespace WinAppATS
{
    class ValidadorFecha
    {
        public void validarmes(KeyPressEventArgs e, string Name, string tbDia)
        {
            if (char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back)
            {
                if (e.KeyChar == (char)Keys.Back)
                {
                    e.Handled = false;
                    return;
                }

                string month = Name.Substring(Name.Length - 2, 2);
                string year = Name.Substring(Name.Length - 6, 4);

                int dia = int.Parse((tbDia + e.KeyChar));
                int anio = int.Parse(year);
                int m = int.Parse(month);

                DateTime firstDayOfMonth = new DateTime(anio, m, 1);

                //Falta validar el ultimo mes del anio de la siguiente linea

                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
                int lasDay = int.Parse(lastDayOfMonth.ToString("dd"));
                if (dia > 0 && dia <= lasDay)
                {
                    e.Handled = false;
                }
                else
                {
                    string mes = lastDayOfMonth.ToLongDateString();
                    mes = mes.Substring(mes.IndexOf(' ') + 1);
                    mes = mes.Substring(mes.IndexOf(' ') + 1);
                    MessageBox.Show("El mes " + mes + " solo tiene " + lasDay.ToString() + " dias");
                    e.Handled = true;
                    return;
                }
            }
            else
            {
                e.Handled = true;
            }
        }
    }
}
