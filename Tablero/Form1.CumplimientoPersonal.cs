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
        private DataTable _desgloseCumplimientoSupervisor;
        private DataTable _desgloseCumplimientoJefe;
        private int? _anioConsultadoSupervisor;
        private int? _anioConsultadoJefe;
        private List<int> _semanasConsultadasSupervisor = new List<int>();
        private List<int> _semanasConsultadasJefe = new List<int>();
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
                PrepararGridCumplimiento(dgvCumplSupDesglose);
                PrepararGridCumplimiento(dgvCumplJefeDesglose);
                // El usuario ve el valor original de cada área, no ponderado.
                dgvCumplSupDesglose.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(39, 99, 166);
                dgvCumplJefeDesglose.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(39, 99, 166);

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
                _desgloseCumplimientoJefe = null;
                _anioConsultadoJefe = null;
                _semanasConsultadasJefe.Clear();
                dgvCumplJefe.DataSource = null;
                dgvCumplJefeDesglose.DataSource = null;
                btnCumplJefeExcel.Enabled = false;
                lblCumplJefeEstado.Text = "Filtros modificados. Pulsa Consultar para actualizar.";
            }
            else
            {
                _datosCumplimientoSupervisor = null;
                _desgloseCumplimientoSupervisor = null;
                _anioConsultadoSupervisor = null;
                _semanasConsultadasSupervisor.Clear();
                dgvCumplSupervisor.DataSource = null;
                dgvCumplSupDesglose.DataSource = null;
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
            CheckedListBox lista = esJefe ? clbCumplJefeSemanas : clbCumplSupSemanas;
            DataGridView gridResumen = esJefe ? dgvCumplJefe : dgvCumplSupervisor;
            DataGridView gridDetalle = esJefe ? dgvCumplJefeDesglose : dgvCumplSupDesglose;
            Label estado = esJefe ? lblCumplJefeEstado : lblCumplSupEstado;
            Button botonExcel = esJefe ? btnCumplJefeExcel : btnCumplSupExcel;
            string nivel = esJefe ? "Jefe de Turno" : "Supervisor";

            if (combo.SelectedItem == null || lista.CheckedItems.Count == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "Selecciona un año y al menos una semana.",
                    "Faltan filtros", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                botonExcel.Enabled = false;
                gridResumen.DataSource = null;
                gridDetalle.DataSource = null;
                if (esJefe)
                {
                    _datosCumplimientoJefe = null;
                    _desgloseCumplimientoJefe = null;
                }
                else
                {
                    _datosCumplimientoSupervisor = null;
                    _desgloseCumplimientoSupervisor = null;
                }

                int anio = Convert.ToInt32(combo.SelectedItem);
                List<int> semanas = SemanasMarcadas(lista);
                var servicio = new Reporte_Cumplimiento_Personal(connectionString);

                // Consultar los dos juegos de datos con filtros idénticos.
                DataTable resumen = servicio.Consultar(nivel, anio, semanas);
                DataTable desglose = servicio.ConsultarDesglose(nivel, anio, semanas);

                // Protección contra versiones de consultas inconsistentes.
                if (resumen.Rows.Count != desglose.Rows.Count)
                    throw new InvalidOperationException("El resumen y el desglose devolvieron diferente cantidad de filas. Reintenta la consulta.");

                if (esJefe)
                {
                    _datosCumplimientoJefe = resumen;
                    _desgloseCumplimientoJefe = desglose;
                    _anioConsultadoJefe = anio;
                    _semanasConsultadasJefe = semanas.ToList();
                }
                else
                {
                    _datosCumplimientoSupervisor = resumen;
                    _desgloseCumplimientoSupervisor = desglose;
                    _anioConsultadoSupervisor = anio;
                    _semanasConsultadasSupervisor = semanas.ToList();
                }

                gridResumen.DataSource = resumen;
                gridDetalle.DataSource = desglose;
                PrepararColumnasCumplimiento(gridResumen);
                PrepararColumnasCumplimiento(gridDetalle);

                int filasConDatos = resumen.AsEnumerable()
                    .Count(r => r.Table.Columns.Contains("Fichas") && r["Fichas"] != DBNull.Value && Convert.ToInt32(r["Fichas"]) > 0);
                estado.Text = string.Format(CultureInfo.CurrentCulture,
                    "{0} persona-semana · {1} con registros · {2} semana(s) · Año {3}. El desglose muestra porcentajes sin ponderar.",
                    resumen.Rows.Count, filasConDatos, semanas.Count, anio);
                botonExcel.Enabled = resumen.Rows.Count > 0;

                if (resumen.Rows.Count == 0)
                    MetroFramework.MetroMessageBox.Show(this, "No hay usuarios activos del nivel " + nivel + ".",
                        "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (esJefe)
                {
                    _datosCumplimientoJefe = null;
                    _desgloseCumplimientoJefe = null;
                }
                else
                {
                    _datosCumplimientoSupervisor = null;
                    _desgloseCumplimientoSupervisor = null;
                }
                gridResumen.DataSource = null;
                gridDetalle.DataSource = null;
                estado.Text = "Error al consultar.";
                MetroFramework.MetroMessageBox.Show(this, "Error al consultar cumplimiento: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private static void PrepararColumnasCumplimiento(DataGridView grid)
        {
            if (grid.Columns.Contains("ID"))
                grid.Columns["ID"].Visible = false;
            foreach (DataGridViewColumn columna in grid.Columns)
            {
                if (columna.Name.EndsWith("(%)", StringComparison.Ordinal))
                {
                    // Los valores se almacenan como 0..100; el signo % es literal.
                    columna.DefaultCellStyle.Format = "0.00'%'";
                    columna.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else if (columna.Name == "Nombre")
                    columna.MinimumWidth = 165;
            }
            if (grid.Columns.Contains("Promedio demás áreas (%)"))
                grid.Columns["Promedio demás áreas (%)"].DefaultCellStyle.BackColor = Color.FromArgb(226, 236, 247);
        }


        private void ExportarCumplimientoPersonal(bool esJefe)
        {
            if (!TienePermisoCumplimientoPersonal()) return;
            DataTable resumen = esJefe ? _datosCumplimientoJefe : _datosCumplimientoSupervisor;
            DataTable desglose = esJefe ? _desgloseCumplimientoJefe : _desgloseCumplimientoSupervisor;
            int? anio = esJefe ? _anioConsultadoJefe : _anioConsultadoSupervisor;
            List<int> semanas = esJefe ? _semanasConsultadasJefe : _semanasConsultadasSupervisor;
            if (resumen == null || desglose == null || !anio.HasValue || resumen.Rows.Count == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "Primero genera ambos reportes con Consultar.",
                    "Exportar Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string titulo = esJefe ? "Cumplimiento por Jefe de Turno" : "Cumplimiento por Supervisor";
            string nombre = (esJefe ? "Cumplimiento_Jefe_de_Turno_" : "Cumplimiento_Supervisor_")
                            + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx";
            using (var dialogo = new SaveFileDialog())
            {
                dialogo.Title = "Guardar " + titulo;
                dialogo.Filter = "Libro de Excel (*.xlsx)|*.xlsx";
                dialogo.DefaultExt = "xlsx";
                dialogo.AddExtension = true;
                dialogo.FileName = nombre;
                if (dialogo.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    using (var libro = new XLWorkbook())
                    {
                        var hoja = libro.Worksheets.Add("Cumplimiento");
                        hoja.Style.Font.FontName = "Calibri";
                        hoja.Style.Font.FontSize = 10;
                        hoja.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        int cols = Math.Max(resumen.Columns.Count - (resumen.Columns.Contains("ID") ? 1 : 0),
                                            desglose.Columns.Count - (desglose.Columns.Contains("ID") ? 1 : 0));
                        hoja.Range(1, 1, 2, cols).Merge();
                        var franja = hoja.Range(1, 1, 2, cols);
                        franja.Style.Fill.BackgroundColor = XLColor.FromHtml("#283B4D");
                        franja.Style.Font.FontColor = XLColor.White;
                        franja.Style.Font.FontSize = 17;
                        franja.Style.Font.Bold = true;
                        franja.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        hoja.Cell(1, 1).Value = titulo.ToUpperInvariant();
                        hoja.Row(1).Height = 23;
                        hoja.Row(2).Height = 17;

                        hoja.Cell(3, 1).Value = "AÑO: " + anio.Value + "   |   SEMANAS: " + string.Join(", ", semanas.Select(x => x.ToString("00")));
                        hoja.Range(3, 1, 3, cols).Merge();
                        hoja.Range(3, 1, 3, cols).Style.Font.Bold = true;
                        hoja.Range(3, 1, 3, cols).Style.Font.FontColor = XLColor.FromHtml("#283B4D");
                        hoja.Row(3).Height = 24;
                        hoja.Cell(4, 1).Value = "Generado: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm") +
                            "  |  Despegue 50 % + demás áreas 50 % si hay Despegue; sin Despegue, demás áreas 100 %.";
                        hoja.Range(4, 1, 4, cols).Merge();
                        hoja.Range(4, 1, 4, cols).Style.Font.FontColor = XLColor.FromHtml("#64748B");

                        int siguiente = EscribirTablaCumplimientoExcel(hoja, resumen, 6,
                            "01  |  RESUMEN PONDERADO", "#EB7823", cols);
                        siguiente += 3;
                        EscribirTablaCumplimientoExcel(hoja, desglose, siguiente,
                            "02  |  DESGLOSE POR ÁREA (PORCENTAJES SIN PONDERAR)", "#2763A6", cols, true);

                        // No usar AutoFiltro: son dos tablas en una misma hoja de trabajo.
                        hoja.Column(1).Width = 11;
                        hoja.Column(2).Width = 12;
                        hoja.Columns(3, cols).AdjustToContents();
                        foreach (var columna in hoja.Columns(1, cols))
                            if (columna.Width > 25) columna.Width = 25;

                        // Dar espacio a los encabezados coloreados del desglose.
                        string[] columnasDetalle = desglose.Columns.Cast<DataColumn>()
                            .Where(c => c.ColumnName != "ID" && c.ColumnName != "Estado")
                            .Select(c => c.ColumnName).ToArray();
                        int indiceDespegue = Array.IndexOf(columnasDetalle, "Despegue (%)") + 1;
                        int indicePromedio = Array.IndexOf(columnasDetalle, "Promedio demás áreas (%)") + 1;
                        if (indiceDespegue > 0 && hoja.Column(indiceDespegue).Width < 19)
                            hoja.Column(indiceDespegue).Width = 19;
                        if (indicePromedio > 0 && hoja.Column(indicePromedio).Width < 25)
                            hoja.Column(indicePromedio).Width = 25;
                        hoja.SheetView.FreezeRows(4);
                        hoja.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                        hoja.PageSetup.FitToPages(1, 0);
                        libro.SaveAs(dialogo.FileName);
                    }
                    MetroFramework.MetroMessageBox.Show(this, "Las dos tablas se guardaron en una sola hoja:\n" + dialogo.FileName,
                        "Excel generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo exportar el archivo: " + ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // En esDesglose=true se identifican visualmente los TRES bloques:
        // NARANJA: Despegue original (si existe, tiene peso 50 % en el resumen).
        // AZUL:    las siete áreas que forman el promedio, solo entre las que tienen datos.
        // VERDE:   Promedio demás áreas (no se incluye Despegue).
        // La separación es SOLO visual: no modifica fórmulas, datos ni ponderaciones.
        private static int EscribirTablaCumplimientoExcel(
            IXLWorksheet hoja, DataTable datos, int filaInicio, string titulo,
            string colorHeader, int anchoTitulo, bool esDesglose = false)
        {
            string[] columnas = datos.Columns.Cast<DataColumn>()
                .Where(c => c.ColumnName != "ID" && c.ColumnName != "Estado")
                .Select(c => c.ColumnName).ToArray();
            int n = columnas.Length;

            int colDespegue = Array.IndexOf(columnas, "Despegue (%)") + 1;
            int colPromedio = Array.IndexOf(columnas, "Promedio demás áreas (%)") + 1;
            string[] otrasAreas = {
                "Evaporado (%)", "Grind (%)", "Inspección (%)", "Empacado (%)",
                "Polvos (%)", "Revolturas (%)", "Máquinas (%)"
            };
            var indicesOtrasAreas = otrasAreas
                .Select(a => Array.IndexOf(columnas, a) + 1)
                .Where(i => i > 0).ToArray();

            // Se aplica agrupación cuando están las columnas esperadas de la consulta.
            bool gruposCompletos = esDesglose && colDespegue > 1 && colPromedio > 0 &&
                                   indicesOtrasAreas.Length == otrasAreas.Length &&
                                   indicesOtrasAreas.Max() - indicesOtrasAreas.Min() == otrasAreas.Length - 1 &&
                                   indicesOtrasAreas.Min() == colDespegue + 1 &&
                                   colPromedio == indicesOtrasAreas.Max() + 1;

            var banda = hoja.Range(filaInicio, 1, filaInicio, anchoTitulo);
            banda.Merge();
            banda.Style.Fill.BackgroundColor = XLColor.FromHtml("#34495E");
            banda.Style.Font.FontColor = XLColor.White;
            banda.Style.Font.Bold = true;
            banda.Style.Font.FontSize = 12;
            hoja.Cell(filaInicio, 1).Value = titulo;
            hoja.Row(filaInicio).Height = 27;

            // Solo el desglose lleva leyenda y franja de grupos encima de los encabezados.
            // Así queda claro que las 7 columnas azules alimentan el promedio verde.
            if (esDesglose)
            {
                int filaLeyenda = filaInicio + 1;
                var leyenda = hoja.Range(filaLeyenda, 1, filaLeyenda, anchoTitulo);
                leyenda.Merge();
                hoja.Cell(filaLeyenda, 1).Value =
                    "NARANJA: Despegue (valor original; pesa 50 % si existe)  |  " +
                    "AZUL: demás áreas (se promedian solo las que tienen registros)  |  " +
                    "VERDE: promedio de esas áreas (pesa 50 % con Despegue o 100 % sin él).";
                leyenda.Style.Fill.BackgroundColor = XLColor.FromHtml("#F0F4F8");
                leyenda.Style.Font.FontColor = XLColor.FromHtml("#334155");
                leyenda.Style.Font.FontSize = 9;
                leyenda.Style.Alignment.WrapText = true;
                leyenda.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                hoja.Row(filaLeyenda).Height = 31;

                int filaGrupos = filaInicio + 2;
                if (gruposCompletos)
                {
                    PintarBandaGrupoExcel(hoja, filaGrupos, 1, colDespegue - 1,
                        "IDENTIFICACIÓN", "#E2E8F0", "#334155");
                    PintarBandaGrupoExcel(hoja, filaGrupos, colDespegue, colDespegue,
                        "DESPEGUE\nPESO 50 %", "#F8C88A", "#743D09");
                    PintarBandaGrupoExcel(hoja, filaGrupos,
                        indicesOtrasAreas.Min(), indicesOtrasAreas.Max(),
                        "7 ÁREAS: BASE DEL PROMEDIO (SOLO LAS QUE TIENEN REGISTROS)",
                        "#D4E7FB", "#174478");
                    PintarBandaGrupoExcel(hoja, filaGrupos, colPromedio, colPromedio,
                        "PROMEDIO\nDEMÁS ÁREAS", "#BFEBDD", "#075B50");
                }
                hoja.Row(filaGrupos).Height = 38;
            }

            int filaEncabezado = filaInicio + (esDesglose ? 3 : 1);
            for (int c = 0; c < n; c++)
            {
                string columna = columnas[c];
                int numCol = c + 1;
                var celda = hoja.Cell(filaEncabezado, numCol);
                celda.Value = columna;
                celda.Style.Fill.BackgroundColor = XLColor.FromHtml(colorHeader);
                celda.Style.Font.FontColor = XLColor.White;
                celda.Style.Font.Bold = true;
                celda.Style.Alignment.WrapText = true;
                celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                celda.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                celda.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                if (gruposCompletos)
                {
                    if (numCol == colDespegue)
                        celda.Style.Fill.BackgroundColor = XLColor.FromHtml("#CB761C");
                    else if (indicesOtrasAreas.Contains(numCol))
                        celda.Style.Fill.BackgroundColor = XLColor.FromHtml("#2763A6");
                    else if (numCol == colPromedio)
                        celda.Style.Fill.BackgroundColor = XLColor.FromHtml("#087C70");
                }
            }
            hoja.Row(filaEncabezado).Height = esDesglose ? 40 : 32;

            int primeraFila = filaEncabezado + 1;
            for (int f = 0; f < datos.Rows.Count; f++)
            {
                DataRow origen = datos.Rows[f];
                int fila = primeraFila + f;
                bool alternada = f % 2 == 1;

                for (int c = 0; c < n; c++)
                {
                    string columna = columnas[c];
                    int numCol = c + 1;
                    object valor = origen[columna];
                    var celda = hoja.Cell(fila, numCol);

                    // PRIORIDAD DE COLORES: incluso en filas alternas y celdas vacías.
                    // Un área sin registros permanece en blanco, pero conserva el color de grupo.
                    if (gruposCompletos)
                    {
                        if (numCol == colDespegue)
                            celda.Style.Fill.BackgroundColor = XLColor.FromHtml(
                                alternada ? "#FBE3C5" : "#FFF1E0");
                        else if (indicesOtrasAreas.Contains(numCol))
                            celda.Style.Fill.BackgroundColor = XLColor.FromHtml(
                                alternada ? "#E7F0FD" : "#F4F8FE");
                        else if (numCol == colPromedio)
                        {
                            celda.Style.Fill.BackgroundColor = XLColor.FromHtml(
                                alternada ? "#CFF0E4" : "#E4F8F0");
                            celda.Style.Font.Bold = true;
                            celda.Style.Font.FontColor = XLColor.FromHtml("#075B50");
                        }
                        else if (alternada)
                            celda.Style.Fill.BackgroundColor = XLColor.FromHtml("#F3F7FA");
                    }
                    else if (alternada && columna != "Total (%)" && columna != "Promedio período (%)")
                    {
                        celda.Style.Fill.BackgroundColor = XLColor.FromHtml("#F3F7FA");
                    }

                    if (valor != null && valor != DBNull.Value)
                    {
                        if (columna.EndsWith("(%)", StringComparison.Ordinal))
                        {
                            decimal porcentaje = Convert.ToDecimal(valor);
                            celda.Value = porcentaje / 100m;
                            celda.Style.NumberFormat.Format = "0.00%";
                            celda.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            if (columna == "Total (%)" || columna == "Promedio período (%)")
                            {
                                celda.Style.Fill.BackgroundColor = XLColor.FromHtml(
                                    porcentaje >= 90m ? "#D5F0DC" : porcentaje >= 80m ? "#FFF4C8" : "#FFE0E0");
                                celda.Style.Font.Bold = true;
                            }
                        }
                        else if (valor is int) celda.Value = (int)valor;
                        else if (valor is long) celda.Value = (long)valor;
                        else if (valor is decimal) celda.Value = (decimal)valor;
                        else celda.Value = Convert.ToString(valor);
                    }

                    if (gruposCompletos && (numCol == colDespegue || numCol == colPromedio))
                    {
                        // Separadores visibles entre Despegue, las 7 áreas y su resultado.
                        celda.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                        celda.Style.Border.LeftBorderColor = XLColor.FromHtml(
                            numCol == colDespegue ? "#CB761C" : "#087C70");
                    }
                }
                hoja.Row(fila).Height = 21;
            }

            if (datos.Rows.Count == 0)
            {
                hoja.Cell(primeraFila, 1).Value = "Sin registros para los filtros seleccionados.";
                hoja.Range(primeraFila, 1, primeraFila, Math.Max(1, n)).Merge();
                return primeraFila + 1;
            }

            var tabla = hoja.Range(filaEncabezado, 1, primeraFila + datos.Rows.Count - 1, n);
            tabla.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tabla.Style.Border.OutsideBorderColor = XLColor.FromHtml("#D6E0EA");
            return primeraFila + datos.Rows.Count;
        }

        private static void PintarBandaGrupoExcel(IXLWorksheet hoja, int fila,
            int columnaInicial, int columnaFinal, string texto, string colorFondo, string colorTexto)
        {
            var rango = hoja.Range(fila, columnaInicial, fila, columnaFinal);
            if (columnaInicial < columnaFinal)
                rango.Merge();
            hoja.Cell(fila, columnaInicial).Value = texto.Replace("\n", Environment.NewLine);
            rango.Style.Fill.BackgroundColor = XLColor.FromHtml(colorFondo);
            rango.Style.Font.FontColor = XLColor.FromHtml(colorTexto);
            rango.Style.Font.Bold = true;
            rango.Style.Font.FontSize = 9;
            rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            rango.Style.Alignment.WrapText = true;
            rango.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            rango.Style.Border.BottomBorderColor = XLColor.FromHtml("#FFFFFF");
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
