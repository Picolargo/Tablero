using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tablero
{
    public static class LoadingScreen
    {
        private static LoadingForm loadingForm;

        // Método para mostrar el formulario de cargando
        public static void ShowLoading()
        {
            if (loadingForm == null)
            {
                loadingForm = new LoadingForm();
                loadingForm.Show();
                Application.DoEvents(); // Permite que el formulario se actualice
            }
        }

        // Método para ocultar el formulario de cargando
        public static void HideLoading()
        {
            if (loadingForm != null)
            {
                loadingForm.Close();
                loadingForm = null;
            }
        }
    }
}
