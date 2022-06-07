using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinAppATS
{
    class Validate
    {
        public bool validateMessage(int tpId, string idContacto)
        {
            bool error = false;
            if (tpId < 2)
            {
                if (tpId == 0)
                {
                    if (idContacto.Length < 13)
                    {
                        MessageBox.Show("El RUC debe tener 13 dígitos", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        error = true;
                    }

                    //if (!error && !validateRuc(idContacto))
                    //{
                    //    MessageBox.Show("No es un RUC válido", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    error = true;
                    //}
                }
                else if (tpId == 1)
                {
                    if (idContacto.Length < 10)
                    {
                        MessageBox.Show("La cédula debe tener 10 dígitos", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        error = true;
                    }

                    if (!error && !validateCedula(idContacto))
                    {
                        MessageBox.Show("La cédula no es válido", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        error = true;
                    }
                }
            }

            return error;
        }

        public bool validateRuc(string ruc)
        {
            char[] vs = ruc.ToCharArray();
            int establecimiento = int.Parse(ruc.Substring(10));
            if (establecimiento > 0)
            {
                return validateCedula(ruc.Substring(0, 10));
            }

            return false;
        }

        public bool validateCedula(string cedula)
        {
            char[] vs = cedula.ToCharArray();
            int provincia = int.Parse(cedula.Substring(0, 2));

            if ((provincia > 0 && provincia < 25) || provincia == 30)
            {
                //Persona natural
                if (int.Parse(vs[2].ToString()) >= 0 && int.Parse(vs[2].ToString()) < 6)
                {
                    int verificar = 0;
                    int coeficiente = 2;
                    int sum = 0;
                    for (int i = 0; i < 9; i++)
                    {
                        int result = coeficiente * int.Parse(vs[i].ToString());
                        result = result > 9 ? result - 9 : result;
                        sum += result;
                        coeficiente = coeficiente == 2 ? 1 : 2;
                    }
                    if (sum % 10 == 0)
                    {
                        verificar = 0;
                    }
                    else
                    {
                        verificar = 10 - (sum % 10);
                    }
                    if (verificar == int.Parse(vs[9].ToString()))
                    {
                        return true;
                    }
                }
                //R.U.C. Públicos:
                else if (int.Parse(vs[2].ToString()) == 6)
                {
                    int coeficiente = 3;
                    int sum = 0;
                    for (int i = 0; i < 8; i++)
                    {
                        int result = coeficiente * int.Parse(vs[i].ToString());
                        sum += result;
                        coeficiente--;
                        if (coeficiente == 1)
                        {
                            coeficiente = 7;
                        }
                    }
                    //Si el residuo es 0
                    if (sum % 11 == 0)
                    {
                        //el digito verificador es 0
                        return int.Parse(vs[0].ToString()) == 0;
                    }
                    sum %= 11;
                    if (11 - sum == int.Parse(vs[8].ToString()))
                    {
                        return true;
                    }
                }
                //R.U.C. Jurídicos y extranjeros sin cédula:
                else if (int.Parse(vs[2].ToString()) == 9)
                {
                    int coeficiente = 4;
                    int sum = 0;
                    for (int i = 0; i < 9; i++)
                    {
                        int result = coeficiente * int.Parse(vs[i].ToString());
                        sum += result;
                        coeficiente--;
                        if (coeficiente == 1)
                        {
                            coeficiente = 7;
                        }
                    }
                    //Si el residuo es 0
                    if (sum % 11 == 0)
                    {
                        //el digito verificador es 0
                        return int.Parse(vs[9].ToString()) == 0;
                    }
                    sum %= 11;
                    if (11 - sum == int.Parse(vs[9].ToString()))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        ~Validate() { }
    }
}
