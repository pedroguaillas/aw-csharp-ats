namespace WinAppATS
{
    class Const
    {
        //Ruta archivos xml
        public static string filexml(string tipo)
        {
            return string.Format("{0}//" + tipo + ".xml", System.Windows.Forms.Application.StartupPath);
        }

        //Ruta reportes
        public static string filereport(string file)
        {
            return "../../Reports/" + file + ".rdlc";
        }

        //Ruta servidor
        public static string URL = @"http://ats.auditwhole.com/";
        //public static string URL = @"https://qr.auditwhole.com/";

        public static double IVA = .12;

        public static char dec = ',';

        public static char nodec = '.';

        public static double round(double num)
        {
            double x = 0;
            string str_num = num.ToString().Trim();
            if (!str_num.Contains(","))
            {
                return num;
            }

            // pasa a un vector de dimensión 2, donde v[0]=numero entero y v[1]=numero decimal
            string[] vec_num = str_num.Split(',');

            // Si el numero decimal tiene 3 cifras y le ultimo digito es 5
            if (vec_num[1].Length == 3 && int.Parse(vec_num[1].Substring(2, 1).Trim()) == 5)
            {
                // Si el punto decimal es 99 incrementar el valor entero en 1
                if (vec_num[1].Substring(0, 2).Equals("99"))
                {
                    x = int.Parse(vec_num[0]) + 1;
                }
                // Si el primer digito del decimal es 0
                else if (vec_num[1].Substring(0, 1).Equals("0"))
                {
                    // Incrementar 1 al digito 2 del numero decimal
                    int num_x = int.Parse(vec_num[1].Substring(0, 2)) + 1;
                    // Si el digito incrementado es 10 mantener si no agregar el 0 en la primera position 
                    vec_num[1] = num_x == 10 ? num_x.ToString() : "0" + num_x;
                    // Asignación de valor redondeado
                    x = double.Parse(vec_num[0] + ',' + vec_num[1]);
                }
                else
                {
                    vec_num[1] = (int.Parse(vec_num[1].Substring(0, 2)) + 1).ToString();
                    // Asignación de valor redondeado
                    x = double.Parse(vec_num[0] + ',' + vec_num[1]);
                }
            }
            else
            {
                x = System.Math.Round(num, 2);
            }

            return x;
        }
    }
}