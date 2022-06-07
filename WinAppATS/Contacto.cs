
namespace WinAppATS
{
    class Contacto
    {
        public string id;
        public string denominacion;
        public string tpId;
        public string tpContacto;
        public string contabilidad;
        public string created_at;
        public string updated_at;

        public Contacto(string id, string denominacion)
        {
            this.id = id;
            this.denominacion = denominacion;
        }
    }

}
