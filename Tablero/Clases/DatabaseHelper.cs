using Npgsql;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Telerik.WinControls.UI;

namespace Tablero
{
    class DatabaseHelper
    {
        private string connectionString;

        public DatabaseHelper(string connectionString)
        {
            this.connectionString = connectionString;
        }

        // Método para probar la conexión
        public bool TestConnection()
        {
            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error de conexión: " + ex.Message);
                    return false;
                }
            }
        }

        // Ejecutar consulta SELECT
        public DataTable ExecuteSelectQuery(string query, NpgsqlParameter[] parameters = null)
        {
            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }

                        NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(command);
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);
                        return dataTable;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error en SELECT: " + ex.Message);
                    return null;
                }
            }
        }

        // Agrega estos métodos a tu clase DatabaseHelper
        public int ExecuteScalarInt(string query, NpgsqlParameter[] parameters = null)
        {
            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }

                        object result = command.ExecuteScalar();
                        return result != null ? Convert.ToInt32(result) : -1;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error en ExecuteScalar: " + ex.Message);
                    return -1;
                }
            }
        }

        // Ejecutar INSERT, UPDATE, DELETE
        public int ExecuteNonQuery(string query, NpgsqlParameter[] parameters = null)
        {
            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        return command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error en operación: " + ex.Message);
                    return -1;
                }
            }
        }

        // Cargar DataGridView
        public void LoadDataIntoDataGridView(string query, DataGridView dataGridView, NpgsqlParameter[] parameters = null)
        {
            try
            {
                DataTable dataTable = ExecuteSelectQuery(query, parameters);
                if (dataTable != null)
                {
                    dataGridView.DataSource = dataTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar datos: " + ex.Message);
            }
        }

        // Cargar RadGridView
        public void LoadDataIntoDataGridViewTelerik(string query, RadGridView radGridView, NpgsqlParameter[] parameters = null)
        {
            try
            {
                DataTable dataTable = ExecuteSelectQuery(query, parameters);
                if (dataTable != null)
                {
                    radGridView.DataSource = dataTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar datos: " + ex.Message);
            }
        }

        // Validación de usuario corregida para PostgreSQL
        public bool ValidateUser(string usuario, string password)
        {
            // Consulta corregida con comillas para nombres de columnas con mayúsculas
            string query = @"SELECT COUNT(1) FROM public.""Usuarios"" 
                           WHERE ""Usuario"" ILIKE @usuario AND ""Password"" = @password";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@usuario", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = usuario },
                new NpgsqlParameter("@password", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = password }
            };

            try
            {
                DataTable result = ExecuteSelectQuery(query, parameters);
                if (result != null && result.Rows.Count > 0)
                {
                    return Convert.ToInt32(result.Rows[0][0]) > 0;
                }
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al validar usuario: " + ex.Message);
                return false;
            }
        }

        // Método adicional para obtener información completa del usuario
        public DataRow GetUserInfo(string usuario)
        {
            string query = @"SELECT ""ID_User"", ""Usuario"", ""No_Empleado"", ""Nivel"" 
                           FROM public.""Usuarios"" WHERE ""Usuario"" ILIKE @usuario";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@usuario", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = usuario }
            };

            try
            {
                DataTable result = ExecuteSelectQuery(query, parameters);
                if (result != null && result.Rows.Count > 0)
                {
                    return result.Rows[0];
                }
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al obtener información del usuario: " + ex.Message);
                return null;
            }
        }

        // Método principal que acepta array de parámetros
        public void LoadDataIntoComboBox(string query, ComboBox comboBox, string displayMember, string valueMember, NpgsqlParameter[] parameters = null)
        {
            try
            {
                DataTable dataTable = ExecuteSelectQuery(query, parameters);

                if (dataTable != null && dataTable.Rows.Count > 0)
                {
                    comboBox.DataSource = dataTable;
                    comboBox.DisplayMember = displayMember;
                    comboBox.ValueMember = valueMember;
                    comboBox.SelectedIndex = -1;
                }
                else
                {
                    comboBox.DataSource = null;
                    comboBox.Items.Clear();
                    comboBox.Text = "No hay datos disponibles";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar ComboBox: {ex.Message}");
            }
        }


        //public void LoadDataIntoComboBox(out bool bande_consulta, string query, ComboBox comboBox, string displayMember, string valueMember, NpgsqlParameter[] parameters = null)
        //{
        //    bande_consulta = false; // Inicializar

        //    try
        //    {
        //        DataTable dataTable = ExecuteSelectQuery(query, parameters);

        //        if (dataTable != null && dataTable.Rows.Count > 0)
        //        {
        //            comboBox.DataSource = dataTable;
        //            comboBox.DisplayMember = displayMember;
        //            comboBox.ValueMember = valueMember;
        //            comboBox.SelectedIndex = -1;
        //            bande_consulta = true;
        //        }
        //        else
        //        {
        //            comboBox.DataSource = null;
        //            comboBox.Items.Clear();
        //            comboBox.Text = "No hay datos disponibles";
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Error al cargar ComboBox: {ex.Message}");
        //    }
        //}

        // Método auxiliar para determinar tipos de datos
        private NpgsqlTypes.NpgsqlDbType GetNpgsqlDbType(object value)
        {
            if (value == null) return NpgsqlTypes.NpgsqlDbType.Varchar;

            Type valueType = value.GetType();

            switch (Type.GetTypeCode(valueType))
            {
                case TypeCode.Int32: return NpgsqlTypes.NpgsqlDbType.Integer;
                case TypeCode.String: return NpgsqlTypes.NpgsqlDbType.Varchar;
                case TypeCode.Decimal: return NpgsqlTypes.NpgsqlDbType.Numeric;
                case TypeCode.Double: return NpgsqlTypes.NpgsqlDbType.Double;
                case TypeCode.Single: return NpgsqlTypes.NpgsqlDbType.Real;
                case TypeCode.Boolean: return NpgsqlTypes.NpgsqlDbType.Boolean;
                case TypeCode.Int64: return NpgsqlTypes.NpgsqlDbType.Bigint;
                case TypeCode.DateTime: return NpgsqlTypes.NpgsqlDbType.Date;
                default: return NpgsqlTypes.NpgsqlDbType.Varchar;
            }
        }

        // Método sobrecargado para un solo parámetro
        public void LoadDataIntoComboBox(string query, ComboBox comboBox, string displayMember, string valueMember, string paramName, object paramValue)
        {
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter(paramName, GetNpgsqlDbType(paramValue)) { Value = paramValue }
            };

            LoadDataIntoComboBox(query, comboBox, displayMember, valueMember, parameters);
        }
    }
}