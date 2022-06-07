using System.Windows.Forms;

namespace WinAppATS
{
    class Importa
    {
        public string selectFolder()
        {
            string path = null;

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    path = fbd.SelectedPath;
                }
            }
            return path;
        }

        public string selectFile(string type)
        {
            string path = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = type;
            openFileDialog.Filter = "Archivos " + type + " (*." + type + ")|*." + type;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog.FileName;
            }
            return path;
        }

        public virtual void Importar(DataGridView dgvResult) { }

        public string addCeroFirst(string cadena)
        {
            if (cadena.StartsWith("."))
                cadena = "0" + cadena;
            return cadena.PadLeft(2);
        }

        ~Importa() { }
    }
}
