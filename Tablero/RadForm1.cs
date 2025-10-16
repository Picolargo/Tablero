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
        public Editar(string connectionString, bool tipo)
        {
            InitializeComponent();
            this.connectionString = connectionString;
        }

        private void RadForm1_Load(object sender, EventArgs e)
        {
            actualiza_fichas();
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
                // 🔹 Columna 0 → ID_Ficha
                string id_global = radGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();

                // 🔹 Columna 4 → Área (recuerda que el índice comienza en 0)
                string area = radGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

                // 🔹 Dispara el evento enviando ambos valores
                FichaSeleccionada?.Invoke(id_global, area);

                // 🔹 Cierra el formulario
                this.Close();
            }
        }
    }
}
