using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;

namespace Tablero
{
    public partial class Editar : Telerik.WinControls.UI.RadForm
    {
        //variable para la conexión a la base de datos
        string connectionString = string.Empty;
        // 🔹 Evento público que enviará dos valores: id_global y area
        public event Action<string, string> FichaSeleccionada;
        // 🔹 Bandera interna para saber si se hizo doble clic
        private bool fichaSeleccionada = false;
        public Editar(string connectionString, bool tipo)
        {
            InitializeComponent();
            this.connectionString = connectionString;
            materialExpansionPanel1.SaveClick += MaterialExpansionPanel1_OnActionButtonClick;
            materialExpansionPanel1.CancelClick += MaterialExpansionPanel1_OnCancelButtonClick;
        }
        // 🔹 Evento del botón "Validar"
        private void MaterialExpansionPanel1_OnActionButtonClick(object sender, EventArgs e)
        {
            MessageBox.Show("✅ Validación completada correctamente.",
                            "Validación",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
            this.Controls.Remove(materialExpansionPanel1); // lo elimina visualmente
            materialExpansionPanel1.Dispose();              // libera recursos
            radGridView1.Visible = true;          // muestra el grid
            actualiza_fichas();
        }

        // 🔹 Evento del botón "Cancelar"
        private void MaterialExpansionPanel1_OnCancelButtonClick(object sender, EventArgs e)
        {
            this.Close(); // 🔹 Cierra el formulario actual
        }
        private void actualiza_fichas()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT ""ID_Ficha"", ""Fecha"", ""Turno"", ""OP"", ""Area"" FROM public.""Ficha"" ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridViewTelerik(querySimple, radGridView1, null);

            // Configurar el DataGridView
            radGridView1.Columns[0].IsVisible = false; // Ocultar la columna ID
            radGridView1.Columns["Fecha"].FormatString = "{0:dd/MM/yyyy}";
        }

        private void radGridView1_CellDoubleClick(object sender, GridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                string id_global = radGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                string area = radGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

                // 🔹 Marca que sí se seleccionó una ficha
                fichaSeleccionada = true;

                // 🔹 Dispara el evento
                FichaSeleccionada?.Invoke(id_global, area);

                // 🔹 Cierra el formulario
                this.Close();
            }
        }

        private void Editar_FormClosed(object sender, FormClosedEventArgs e)
        {
            // 🔹 Si no se seleccionó nada, dispara un evento con valores nulos opcionalmente
            if (!fichaSeleccionada)
            {
                MessageBox.Show("❌ Se canceló la operación.",
                            "Cancelado",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                FichaSeleccionada?.Invoke(null, null);
            }
        }
    }
}
