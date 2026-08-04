using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;
using Telerik.WinControls.UI;

namespace Tablero
{
    public partial class FormTrazabilidad : MaterialForm
    {
        private string connectionString;
        private DatabaseHelper dbHelper;

        public FormTrazabilidad(string connectionString)
        {
            InitializeComponent();
            this.connectionString = connectionString;
            dbHelper = new DatabaseHelper(connectionString);

            // Initialize MaterialSkinManager and set the theme and color scheme  
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Orange600, Primary.Orange600, Primary.BlueGrey800, Accent.Blue700, TextShade.WHITE);

            ConfigurarRadGridView();
        }

        private void ConfigurarRadGridView()
        {
            rgvTrazabilidad.EnableGrouping = false;
            rgvTrazabilidad.EnableHotTracking = true;
            rgvTrazabilidad.ShowFilteringRow = false;
            rgvTrazabilidad.EnableFiltering = true;
            rgvTrazabilidad.EnableCustomFiltering = true;
            rgvTrazabilidad.MasterTemplate.AllowAddNewRow = false;
            rgvTrazabilidad.MasterTemplate.AllowDeleteRow = false;
            rgvTrazabilidad.MasterTemplate.AllowEditRow = false;
            rgvTrazabilidad.MasterTemplate.AutoSizeColumnsMode = GridViewAutoSizeColumnsMode.Fill;
        }

        private void FormTrazabilidad_Load(object sender, EventArgs e)
        {
            CargarUsuarios();
            dtpDesde.Value = DateTime.Now.AddMonths(-1);
            dtpHasta.Value = DateTime.Now;
            btnBuscar.PerformClick();
        }

        private void CargarUsuarios()
        {
            try
            {
                string query = @"SELECT DISTINCT ""Usuario_Edito"" 
                                FROM public.""Trazabilidad_Ediciones_Ficha"" 
                                ORDER BY ""Usuario_Edito""";
                DataTable dt = dbHelper.ExecuteSelectQuery(query);

                cmbUsuario.Items.Clear();
                cmbUsuario.Items.Add("Todos");
                foreach (DataRow row in dt.Rows)
                {
                    cmbUsuario.Items.Add(row["Usuario_Edito"].ToString());
                }
                cmbUsuario.SelectedIndex = 0;
            }
            catch
            {
                cmbUsuario.Items.Clear();
                cmbUsuario.Items.Add("Todos");
                cmbUsuario.SelectedIndex = 0;
            }
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            try
            {
                string fechaInicio = dtpDesde.Value.Date.ToString("yyyy-MM-dd");
                string fechaFin = dtpHasta.Value.Date.AddDays(1).ToString("yyyy-MM-dd");
                string usuario = cmbUsuario.SelectedItem?.ToString();

                if (usuario == "Todos") usuario = "";

                DataTable dt = dbHelper.ObtenerTrazabilidadConFiltros(fechaInicio, fechaFin, usuario);

                rgvTrazabilidad.DataSource = dt;

                if (dt != null && dt.Rows.Count > 0)
                {
                    // Configurar columnas
                    if (rgvTrazabilidad.Columns.Contains("ID_Trazabilidad"))
                        rgvTrazabilidad.Columns["ID_Trazabilidad"].IsVisible = false;

                    if (rgvTrazabilidad.Columns.Contains("Fecha_Edicion"))
                    {
                        rgvTrazabilidad.Columns["Fecha_Edicion"].FormatString = "{0:dd/MM/yyyy HH:mm:ss}";
                        rgvTrazabilidad.Columns["Fecha_Edicion"].TextAlignment = ContentAlignment.MiddleCenter;
                    }

                    if (rgvTrazabilidad.Columns.Contains("ID_Ficha"))
                    {
                        rgvTrazabilidad.Columns["ID_Ficha"].HeaderText = "ID Ficha";
                        rgvTrazabilidad.Columns["ID_Ficha"].TextAlignment = ContentAlignment.MiddleCenter;
                    }

                    if (rgvTrazabilidad.Columns.Contains("Usuario_Edito"))
                        rgvTrazabilidad.Columns["Usuario_Edito"].HeaderText = "Usuario que Editó";

                    if (rgvTrazabilidad.Columns.Contains("Nivel_Usuario"))
                        rgvTrazabilidad.Columns["Nivel_Usuario"].HeaderText = "Nivel";

                    if (rgvTrazabilidad.Columns.Contains("No_Empleado"))
                        rgvTrazabilidad.Columns["No_Empleado"].HeaderText = "No. Empleado";

                    lblTotalRegistros.Text = $"Total: {dt.Rows.Count} registros";
                    btnExportExcel.Enabled = true;
                    btnLimpiar.Enabled = true;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron registros de trazabilidad con los filtros seleccionados.",
                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnExportExcel.Enabled = false;
                    btnLimpiar.Enabled = false;
                    lblTotalRegistros.Text = "Total: 0 registros";
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al buscar trazabilidad: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            dtpDesde.Value = DateTime.Now.AddMonths(-1);
            dtpHasta.Value = DateTime.Now;
            cmbUsuario.SelectedIndex = 0;
            rgvTrazabilidad.DataSource = null;
            rgvTrazabilidad.Rows.Clear();
            rgvTrazabilidad.Columns.Clear();
            lblTotalRegistros.Text = "Total: 0 registros";
            btnExportExcel.Enabled = false;
            btnLimpiar.Enabled = false;
            CargarUsuarios();
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Guardar trazabilidad";
                    saveFileDialog.FileName = $"Trazabilidad_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportarRadGridViewAExcel(rgvTrazabilidad, saveFileDialog.FileName);

                        MetroFramework.MetroMessageBox.Show(this,
                            "Exportación completada con éxito.",
                            "Exportación", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al exportar: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportarRadGridViewAExcel(RadGridView radGridView, string filePath)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Trazabilidad");

                int colIndex = 1;
                foreach (GridViewDataColumn column in radGridView.Columns)
                {
                    if (column.IsVisible)
                    {
                        var cell = worksheet.Cell(1, colIndex);
                        cell.Value = column.HeaderText;
                        cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.Orange;
                        cell.Style.Font.Bold = true;
                        cell.Style.Font.FontColor = ClosedXML.Excel.XLColor.Black;
                        cell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                        colIndex++;
                    }
                }

                int rowIndex = 2;
                foreach (GridViewRowInfo row in radGridView.ChildRows)
                {
                    if (row.IsVisible && !(row is GridViewGroupRowInfo))
                    {
                        colIndex = 1;
                        foreach (GridViewDataColumn column in radGridView.Columns)
                        {
                            if (column.IsVisible)
                            {
                                worksheet.Cell(rowIndex, colIndex).Value = row.Cells[column.Name].Value?.ToString();
                                colIndex++;
                            }
                        }
                        rowIndex++;
                    }
                }

                var usedRange = worksheet.Range(1, 1, Math.Max(rowIndex - 1, 2), colIndex - 1);
                usedRange.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                usedRange.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }
        }
    }
}