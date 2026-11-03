using ClosedXML.Excel;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Tablero
{
    public partial class Form_principal
    {
        private DataTable _datosCumplimientoSupervisor;
        private DataTable _datosCumplimientoJefe;
        private bool _inicializandoCumplimientoPersonal;

        private bool TienePermisoCumplimientoPersonal()
        {
            return (nivel_user == "Administrador" || nivel_user == "Super Administrador")
                && materialTabControl1.TabPages.Contains(tabPage10);
        }

        // Ejecutar una vez al FINAL de Form_principal_Load, después de aplicar permisos,
        // pero ANTES de InicializarSelectorDeReportes().
        private void InicializarCumplimientoPersonal()
        {
            if (!TienePermisoCumplimientoPersonal())
            {
                // Nunca dejar estas nuevas vistas disponibles a otros niveles por accidente.
                tapcontrol.TabPages.Remove(tabPageCumplSup);
                tapcontrol.TabPages.Remove(tabPageCumplJefe);
                return;
            }

            _inicializandoCumplimientoPersonal = true;
            try
            {
                var servicio = new Reporte_Cumplimiento_Personal(connectionString);
                List<int> anios = servicio.ObtenerAniosDisponibles();
                cbCumplSupAnio.Items.Clear();
                cbCumplJefeAnio.Items.Clear();
                foreach (int anio in anios)
                {
                    cbCumplSupAnio.Items.Add(anio);
                    cbCumplJefeAnio.Items.Add(anio);
                }
                if (anios.Count > 0)
                {
                    cbCumplSupAnio.SelectedIndex = 0;
                    cbCumplJefeAnio.SelectedIndex = 0;
                }

                PrepararSemanasCumplimiento(clbCumplSupSemanas);
                PrepararSemanasCumplimiento(clbCumplJefeSemanas);
                PrepararGridCumplimiento(dgvCumplSupervisor);
                PrepararGridCumplimiento(dgvCumplJefe);

                lblCumplSupEstado.Text = "Selecciona las semanas y pulsa Consultar.";
                lblCumplJefeEstado.Text = "Selecciona las semanas y pulsa Consultar.";
                btnCumplSupExcel.Enabled = false;
                btnCumplJefeExcel.Enabled = false;
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "No fue posible inicializar los reportes por persona: " + ex.Message,
                    "Reportes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tapcontrol.TabPages.Remove(tabPageCumplSup);
                tapcontrol.TabPages.Remove(tabPageCumplJefe);
                return;
            }
            finally
            {
                _inicializandoCumplimientoPersonal = false;
            }

            // Evitar datos obsoletos si el usuario modifica los filtros.
            cbCumplSupAnio.SelectedIndexChanged += (s, e) => InvalidarCumplimiento(false);
            cbCumplJefeAnio.SelectedIndexChanged += (s, e) => InvalidarCumplimiento(true);
            clbCumplSupSemanas.ItemCheck += (s, e) => InvalidarCumplimiento(false);
            clbCumplJefeSemanas.ItemCheck += (s, e) => InvalidarCumplimiento(true);
        }

        private void PrepararSemanasCumplimiento(CheckedListBox lista)
        {
            lista.Items.Clear();
            for (int semana = 1; semana <= 53; semana++)
                lista.Items.Add(semana.ToString("00", CultureInfo.InvariantCulture));

            // Selección inicial: última semana completa.
            DateTime hoy = DateTime.Today;
            int diferencia = (hoy.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)hoy.DayOfWeek) - 1;
            DateTime domingoAnterior = hoy.AddDays(-diferencia - 1);
            int semanaAnterior = ObtenerNumeroSemana(domingoAnterior);
            if (semanaAnterior >= 1 && semanaAnterior <= lista.Items.Count)
                lista.SetItemChecked(semanaAnterior - 1, true);
        }

        private static void PrepararGridCumplimiento(DataGridView grid)
        {
            grid.AutoGenerateColumns = true;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.MultiSelect = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            grid.RowHeadersVisible = false;
            grid.EnableHeadersVisualStyles = false;
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.None;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 120, 35);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            grid.ColumnHeadersHeight = 38;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            grid.DefaultCellStyle.Font = new Font("Segoe UI", 10F);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(35, 50, 75);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(247, 249, 252);
            grid.RowTemplate.Height = 30;
            grid.CellFormatting -= FormatearCeldaCumplimiento;
            grid.CellFormatting += FormatearCeldaCumplimiento;
        }

        private static void FormatearCeldaCumplimiento(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var grid = (DataGridView)sender;
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;
            string nombre = grid.Columns[e.ColumnIndex].Name;
            if (nombre != "Total (%)" && nombre != "Promedio período (%)")
                return;
            if (e.Value == null || e.Value == DBNull.Value)
                return;

            decimal porcentaje;
            if (!decimal.TryParse(Convert.ToString(e.Value, CultureInfo.InvariantCulture),
                    NumberStyles.Number, CultureInfo.InvariantCulture, out porcentaje))
                return;

            // Priorizar selección para que el texto continúe siendo legible.
            if (grid.Rows[e.RowIndex].Selected)
                return;
            e.CellStyle.BackColor = porcentaje >= 90m ? Color.FromArgb(213, 240, 220)
                                 : porcentaje >= 80m ? Color.FromArgb(255, 244, 200)
                                 : Color.FromArgb(255, 224, 224);
            e.CellStyle.ForeColor = Color.FromArgb(40, 50, 65);
        }

        private void InvalidarCumplimiento(bool esJefe)
        {
            if (_inicializandoCumplimientoPersonal)
                return;

            if (esJefe)
            {
                _datosCumplimientoJefe = null;
                dgvCumplJefe.DataSource = null;
                btnCumplJefeExcel.Enabled = false;
                lblCumplJefeEstado.Text = "Filtros modificados. Pulsa Consultar para actualizar.";
            }
            else
            {
                _datosCumplimientoSupervisor = null;
                dgvCumplSupervisor.DataSource = null;
                btnCumplSupExcel.Enabled = false;
                lblCumplSupEstado.Text = "Filtros modificados. Pulsa Consultar para actualizar.";
            }
        }

        private static List<int> SemanasMarcadas(CheckedListBox control)
        {
            return control.CheckedItems.Cast<object>()
                .Select(i => int.Parse(i.ToString(), CultureInfo.InvariantCulture))
                .OrderBy(s => s)
                .ToList();
        }

        private void EjecutarCumplimientoPersonal(bool esJefe)
        {
            if (!TienePermisoCumplimientoPersonal())
            {
                MetroFramework.MetroMessageBox.Show(this, "No tienes permiso para consultar este reporte.",
                    "Acceso restringido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ComboBox combo = esJefe ? cbCumplJefeAnio : cbCumplSupAnio;
            CheckedListBox semanasControl = esJefe ? clbCumplJefeSemanas : clbCumplSupSemanas;
            DataGridView grid = esJefe ? dgvCumplJefe : dgvCumplSupervisor;
            Label estado = esJefe ? lblCumplJefeEstado : lblCumplSupEstado;
            Button botonExcel = esJefe ? btnCumplJefeExcel : btnCumplSupExcel;
            string nivel = esJefe ? "Jefe de Turno" : "Supervisor";

            if (combo.SelectedItem == null || semanasControl.CheckedItems.Count == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "Selecciona un año y al menos una semana.",
                    "Faltan filtros", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                botonExcel.Enabled = false;
                grid.DataSource = null;
                int anio = Convert.ToInt32(combo.SelectedItem);
                List<int> semanas = SemanasMarcadas(semanasControl);
                var servicio = new Reporte_Cumplimiento_Personal(connectionString);
                DataTable datos = servicio.Consultar(nivel, anio, semanas);

                if (esJefe)
                    _datosCumplimientoJefe = datos;
                else
                    _datosCumplimientoSupervisor = datos;

                grid.DataSource = datos;
                if (grid.Columns.Contains("ID")) grid.Columns["ID"].Visible = false;
                foreach (string nombre in new[] { "Despegue (%)", "Demás áreas (%)", "Total (%)", "Promedio período (%)" })
                    if (grid.Columns.Contains(nombre))
                        grid.Columns[nombre].DefaultCellStyle.Format = "N2";
                if (grid.Columns.Contains("Nombre"))
                    grid.Columns["Nombre"].MinimumWidth = 185;

                int conDatos = datos.AsEnumerable().Count(r => Convert.ToInt32(r["Fichas"]) > 0);
                estado.Text = string.Format(CultureInfo.CurrentCulture,
                    "{0} filas · {1} con registros · {2} semana(s) · Año {3}. Sin registros se deja el porcentaje en blanco.",
                    datos.Rows.Count, conDatos, semanas.Count, anio);
                botonExcel.Enabled = datos.Rows.Count > 0;

                if (datos.Rows.Count == 0)
                    MetroFramework.MetroMessageBox.Show(this, "No hay usuarios activos de nivel " + nivel + ".",
                        "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (esJefe) _datosCumplimientoJefe = null;
                else _datosCumplimientoSupervisor = null;
                estado.Text = "Error al consultar.";
                MetroFramework.MetroMessageBox.Show(this, "Error al consultar cumplimiento: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void ExportarCumplimientoPersonal(bool esJefe)
        {
            if (!TienePermisoCumplimientoPersonal())
                return;
            DataTable datos = esJefe ? _datosCumplimientoJefe : _datosCumplimientoSupervisor;
            if (datos == null || datos.Rows.Count == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "Primero genera el reporte con el botón Consultar.",
                    "Exportar Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string titulo = esJefe ? "Cumplimiento por Jefe de Turno" : "Cumplimiento por Supervisor";
            string nombreSugerido = (esJefe ? "Cumplimiento_Jefe_de_Turno_" : "Cumplimiento_Supervisor_") + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx";
            using (var dialogo = new SaveFileDialog())
            {
                dialogo.Title = "Guardar " + titulo;
                dialogo.Filter = "Archivo de Excel (*.xlsx)|*.xlsx";
                dialogo.DefaultExt = "xlsx";
                dialogo.AddExtension = true;
                dialogo.FileName = nombreSugerido;
                if (dialogo.ShowDialog(this) != DialogResult.OK)
                    return;

                try
                {
                    // Ya se utiliza ClosedXML en el proyecto. No requiere enviar correo ni EPPlus.
                    using (var libro = new XLWorkbook())
                    {
                        var hoja = libro.Worksheets.Add("Cumplimiento");
                        string[] columnas = datos.Columns.Cast<DataColumn>()
                            .Where(c => c.ColumnName != "ID")
                            .Select(c => c.ColumnName).ToArray();
                        int cantidadColumnas = columnas.Length;
                        hoja.Cell(1, 1).Value = titulo.ToUpperInvariant();
                        hoja.Range(1, 1, 1, cantidadColumnas).Merge();
                        hoja.Range(1, 1, 1, cantidadColumnas).Style.Fill.BackgroundColor = XLColor.FromHtml("#34495E");
                        hoja.Range(1, 1, 1, cantidadColumnas).Style.Font.FontColor = XLColor.White;
                        hoja.Range(1, 1, 1, cantidadColumnas).Style.Font.Bold = true;
                        hoja.Range(1, 1, 1, cantidadColumnas).Style.Font.FontSize = 15;
                        hoja.Row(1).Height = 30;
                        hoja.Cell(2, 1).Value = "Generado: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                        hoja.Cell(3, 1).Value = "Despegue: 50 % / demás áreas: 50 % si hay Despegue; sin Despegue: 100 % demás áreas.";
                        hoja.Range(3, 1, 3, cantidadColumnas).Merge();

                        for (int col = 0; col < columnas.Length; col++)
                        {
                            hoja.Cell(5, col + 1).Value = columnas[col];
                            hoja.Cell(5, col + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#EB7823");
                            hoja.Cell(5, col + 1).Style.Font.FontColor = XLColor.White;
                            hoja.Cell(5, col + 1).Style.Font.Bold = true;
                        }

                        for (int fila = 0; fila < datos.Rows.Count; fila++)
                        {
                            DataRow origen = datos.Rows[fila];
                            for (int col = 0; col < columnas.Length; col++)
                            {
                                object valor = origen[columnas[col]];
                                var celda = hoja.Cell(fila + 6, col + 1);
                                if (valor == DBNull.Value || valor == null)
                                    continue;
                                if (columnas[col].EndsWith("(%)", StringComparison.Ordinal))
                                {
                                    celda.Value = Convert.ToDecimal(valor) / 100m;
                                    celda.Style.NumberFormat.Format = "0.00%";
                                }
                                else if (valor is int)
                                    celda.Value = (int)valor;
                                else if (valor is long)
                                    celda.Value = (long)valor;
                                else if (valor is decimal)
                                    celda.Value = (decimal)valor;
                                else
                                    celda.Value = Convert.ToString(valor);
                            }
                            if (fila % 2 == 1)
                                hoja.Range(fila + 6, 1, fila + 6, cantidadColumnas)
                                    .Style.Fill.BackgroundColor = XLColor.FromHtml("#F3F7FA");
                        }

                        hoja.SheetView.FreezeRows(5);
                        hoja.Range(5, 1, 5 + datos.Rows.Count, cantidadColumnas).SetAutoFilter();
                        hoja.Columns().AdjustToContents();
                        libro.SaveAs(dialogo.FileName);
                    }

                    MetroFramework.MetroMessageBox.Show(this, "Archivo guardado:\n" + dialogo.FileName,
                        "Excel generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo exportar el archivo: " + ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnCumplSupConsultar_Click(object sender, EventArgs e) => EjecutarCumplimientoPersonal(false);
        private void btnCumplJefeConsultar_Click(object sender, EventArgs e) => EjecutarCumplimientoPersonal(true);
        private void btnCumplSupExcel_Click(object sender, EventArgs e) => ExportarCumplimientoPersonal(false);
        private void btnCumplJefeExcel_Click(object sender, EventArgs e) => ExportarCumplimientoPersonal(true);

        private void btnCumplSupTodas_Click(object sender, EventArgs e) => MarcarSemanas(clbCumplSupSemanas, true);
        private void btnCumplJefeTodas_Click(object sender, EventArgs e) => MarcarSemanas(clbCumplJefeSemanas, true);
        private void btnCumplSupLimpiar_Click(object sender, EventArgs e) => MarcarSemanas(clbCumplSupSemanas, false);
        private void btnCumplJefeLimpiar_Click(object sender, EventArgs e) => MarcarSemanas(clbCumplJefeSemanas, false);

        private static void MarcarSemanas(CheckedListBox lista, bool marcar)
        {
            lista.BeginUpdate();
            try
            {
                for (int i = 0; i < lista.Items.Count; i++)
                    lista.SetItemChecked(i, marcar);
            }
            finally
            {
                lista.EndUpdate();
            }
        }
    }
}
