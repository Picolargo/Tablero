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
        // Método auxiliar para detectar errores de red - SOLO para el login
        private bool IsNetworkError(Exception ex)
        {
            string errorMessage = ex.Message.ToLower();
            return errorMessage.Contains("network") ||
                   errorMessage.Contains("connection") ||
                   errorMessage.Contains("timeout") ||
                   errorMessage.Contains("host") ||
                   errorMessage.Contains("unable to connect") ||
                   errorMessage.Contains("no such host") ||
                   errorMessage.Contains("connection refused");
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
                    MessageBox.Show("Tipo de Eror: " + ex.Message+ ", Favor de verificar su conexión a Internet","Error de Red",MessageBoxButtons.OK,MessageBoxIcon.Error);
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

        // Validación de usuario corregida para PostgreSQL - acepta Usuario o No_Empleado
        //public bool ValidateUser(string identificador, string password)
        //{
        //    // Consulta que acepta tanto Usuario como No_Empleado
        //    string query = @"SELECT COUNT(1) FROM public.""Usuarios"" 
        //           WHERE (""Usuario"" ILIKE @identificador OR ""No_Empleado"" = @identificador) 
        //           AND ""Password"" = @password";

        //    NpgsqlParameter[] parameters = new NpgsqlParameter[]
        //    {
        //        new NpgsqlParameter("@identificador", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = identificador },
        //        new NpgsqlParameter("@password", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = password }
        //    };

        //    try
        //    {
        //        DataTable result = ExecuteSelectQuery(query, parameters);
        //        if (result != null && result.Rows.Count > 0)
        //        {
        //            return Convert.ToInt32(result.Rows[0][0]) > 0;
        //        }
        //        return false;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error al validar usuario: " + ex.Message);
        //        return false;
        //    }
        //}
        // Validación de usuario corregida para PostgreSQL - acepta Usuario o No_Empleado
        public bool ValidateUser(string identificador, string password)
        {
            return ValidateUser(identificador, password, out _);
        }

        // Nueva versión con manejo de error de conexión
        public bool ValidateUser(string identificador, string password, out string errorMessage)
        {
            errorMessage = string.Empty;

            // Consulta que acepta tanto Usuario como No_Empleado
            string query = @"SELECT COUNT(1) FROM public.""Usuarios"" 
           WHERE (""Usuario"" ILIKE @identificador OR ""No_Empleado"" = @identificador) 
           AND ""Password"" = @password";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@identificador", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = identificador },
        new NpgsqlParameter("@password", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = password }
            };

            try
            {
                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        command.Parameters.AddRange(parameters);

                        object result = command.ExecuteScalar();
                        return result != null ? Convert.ToInt32(result) > 0 : false;
                    }
                }
            }
            catch (Exception ex)
            {
                if (IsNetworkError(ex))
                {
                    errorMessage = "CONNECTION_ERROR";
                }
                else
                {
                    errorMessage = ex.Message;
                }
                return false;
            }
        }

        // Método adicional para obtener información completa del usuario
        //public DataRow GetUserInfo(string identificador)
        //{
        //    string query = @"SELECT ""ID_User"", ""Usuario"", ""No_Empleado"", ""Nivel"" FROM public.""Usuarios"" 
        //           WHERE ""Usuario"" ILIKE @identificador OR ""No_Empleado"" = @identificador";

        //    NpgsqlParameter[] parameters = new NpgsqlParameter[]
        //    {
        //        new NpgsqlParameter("@identificador", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = identificador },
        //    };

        //    try
        //    {
        //        DataTable result = ExecuteSelectQuery(query, parameters);
        //        if (result != null && result.Rows.Count > 0)
        //        {
        //            return result.Rows[0];
        //        }
        //        return null;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error al obtener información del usuario: " + ex.Message);
        //        return null;
        //    }
        //}
        // Método adicional para obtener información completa del usuario
        public DataRow GetUserInfo(string identificador)
        {
            return GetUserInfo(identificador, out _);
        }

        // Nueva versión con manejo de error de conexión
        public DataRow GetUserInfo(string identificador, out string errorMessage)
        {
            errorMessage = string.Empty;

            string query = @"SELECT ""ID_User"", ""Usuario"", ""No_Empleado"", ""Nivel"" FROM public.""Usuarios"" 
           WHERE ""Usuario"" ILIKE @identificador OR ""No_Empleado"" = @identificador";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@identificador", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = identificador },
            };

            try
            {
                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        command.Parameters.AddRange(parameters);

                        using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            if (dataTable.Rows.Count > 0)
                            {
                                return dataTable.Rows[0];
                            }
                            return null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (IsNetworkError(ex))
                {
                    errorMessage = "CONNECTION_ERROR";
                }
                else
                {
                    errorMessage = ex.Message;
                }
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