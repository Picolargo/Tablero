// Archivo adicional de la clase Form_principal.
// Agrégalo al proyecto como "Form1.ReportesSelector.cs" (Agregar > Elemento existente).
// En Form1.cs, agrega una llamada al final del método Form_principal_Load:
//     InicializarSelectorDeReportes();
// IMPORTANTE: después de las condiciones de nivel_user, fuera de todos los if.

using System;
using System.Windows.Forms;

namespace Tablero
{
    public partial class Form_principal
    {
        // Conservar el objeto TabPage permite seleccionar un reporte sin perder
        // los controles, filtros, eventos o políticas existentes.
        private sealed class OpcionReporte
        {
            public TabPage Pagina { get; private set; }

            public OpcionReporte(TabPage pagina)
            {
                Pagina = pagina;
            }

            public override string ToString()
            {
                return Pagina.Text;
            }
        }

        private void InicializarSelectorDeReportes()
        {
            // Algunas cuentas no tienen acceso al módulo Reportes completo.
            if (!materialTabControl1.TabPages.Contains(tabPage10))
                return;

            // Evita suscripciones duplicadas si se reinicializa el selector.
            cb_seleccionar_reporte.SelectedIndexChanged -= cb_seleccionar_reporte_SelectedIndexChanged;
            tapcontrol.SelectedIndexChanged -= tapcontrol_SelectedIndexChanged_Reportes;

            cb_seleccionar_reporte.BeginUpdate();
            try
            {
                cb_seleccionar_reporte.Items.Clear();

                // Los permisos YA se aplicaron en Form_principal_Load mediante
                // tapcontrol.TabPages.Remove(...). Solo mostramos las que quedan.
                foreach (TabPage pagina in tapcontrol.TabPages)
                    cb_seleccionar_reporte.Items.Add(new OpcionReporte(pagina));

                cb_seleccionar_reporte.Enabled = cb_seleccionar_reporte.Items.Count > 0;

                if (tapcontrol.TabPages.Count > 0)
                {
                    if (tapcontrol.SelectedTab == null)
                        tapcontrol.SelectedIndex = 0;

                    SincronizarComboConReporte();
                }
                else
                {
                    lbl_ayuda_reportes.Text = "No hay reportes habilitados para este usuario.";
                }
            }
            finally
            {
                cb_seleccionar_reporte.EndUpdate();
            }

            cb_seleccionar_reporte.SelectedIndexChanged += cb_seleccionar_reporte_SelectedIndexChanged;
            tapcontrol.SelectedIndexChanged += tapcontrol_SelectedIndexChanged_Reportes;

            // Se adapta al cambio de tamaño de ventana o escalado de Windows.
            pnl_contenido_reportes.Resize -= pnl_contenido_reportes_Resize;
            pnl_contenido_reportes.Resize += pnl_contenido_reportes_Resize;
            AjustarAreaVisibleReportes();
        }

        private void pnl_contenido_reportes_Resize(object sender, EventArgs e)
        {
            AjustarAreaVisibleReportes();
        }

        private void AjustarAreaVisibleReportes()
        {
            // La altura de las pestañas depende del escalado/DPI y del tema.
            // El panel contenedor recorta esa franja, sin quitar TabPages.
            int altoPestanas = (tapcontrol.SelectedTab != null)
                ? tapcontrol.SelectedTab.Top
                : 38;
            if (altoPestanas < 0 || altoPestanas > 120)
                altoPestanas = 38;

            tapcontrol.Location = new System.Drawing.Point(0, -altoPestanas);
            tapcontrol.Size = new System.Drawing.Size(
                pnl_contenido_reportes.ClientSize.Width,
                pnl_contenido_reportes.ClientSize.Height + altoPestanas + 4);
        }

        private void cb_seleccionar_reporte_SelectedIndexChanged(object sender, EventArgs e)
        {
            OpcionReporte opcion = cb_seleccionar_reporte.SelectedItem as OpcionReporte;

            // Doble comprobación: solo permite pestañas que siguen autorizadas.
            if (opcion == null || !tapcontrol.TabPages.Contains(opcion.Pagina))
                return;

            if (tapcontrol.SelectedTab != opcion.Pagina)
                tapcontrol.SelectedTab = opcion.Pagina;
        }

        private void tapcontrol_SelectedIndexChanged_Reportes(object sender, EventArgs e)
        {
            // También sincroniza el ComboBox si otro proceso cambia la pestaña.
            SincronizarComboConReporte();
        }

        private void SincronizarComboConReporte()
        {
            TabPage seleccionada = tapcontrol.SelectedTab;
            if (seleccionada == null)
                return;

            for (int i = 0; i < cb_seleccionar_reporte.Items.Count; i++)
            {
                OpcionReporte opcion = cb_seleccionar_reporte.Items[i] as OpcionReporte;
                if (opcion != null && opcion.Pagina == seleccionada)
                {
                    if (cb_seleccionar_reporte.SelectedIndex != i)
                        cb_seleccionar_reporte.SelectedIndex = i;
                    return;
                }
            }
        }
    }
}
