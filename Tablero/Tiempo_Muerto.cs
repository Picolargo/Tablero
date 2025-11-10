using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;

namespace Tablero
{
    public partial class Tiempo_Muerto : Telerik.WinControls.UI.RadForm
    {
        //variable para la conexión a la base de datos
        string connectionString = string.Empty;
        string id_ficha = string.Empty;
        public Tiempo_Muerto(string connectionString, string id_ficha)
        {
            this.connectionString = connectionString;
            this.id_ficha = id_ficha;
            InitializeComponent();
        }

        private void Tiempo_Muerto_Load(object sender, EventArgs e)
        {
            actualiza_fichas();
        }
        private void actualiza_fichas()
        {
            int id_ficha_int = int.Parse(id_ficha);
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
    'Operativo' AS ""Tipo de Tiempo Muerto"",
    ""Min_Detenidos"" AS ""Minutos Detenidos"", 
    ""Motivos"" 
FROM public.""Tiempo_Muerto_Operativo"" 
WHERE ""ID_Ficha"" = @idficha

UNION ALL

SELECT 
    'Mecánico' AS ""Tipo de Tiempo Muerto"",
    ""Min_Detenidos"" AS ""Minutos Detenidos"", 
    ""Motivos"" 
FROM public.""Tiempo_muerto_Mecanico"" 
WHERE ""ID_Ficha"" = @idficha

ORDER BY ""Tipo de Tiempo Muerto"" DESC;";
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@idficha", id_ficha_int),
            };
            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridViewTelerik(querySimple, radGridView1, parameters);
        }
    }
}
