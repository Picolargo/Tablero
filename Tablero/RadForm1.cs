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
    public partial class RadForm1 : Telerik.WinControls.UI.RadForm
    {
        //variable para la conexión a la base de datos
        string connectionString = string.Empty;
        public RadForm1(string connectionString)
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
    }
}
