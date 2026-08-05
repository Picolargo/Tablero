using Npgsql;
using System;
using System.Collections.Generic;
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
        /// <summary>
        /// Registra la trazabilidad de una edición de ficha (versión simplificada)
        /// </summary>
        public int RegistrarTrazabilidadEdicion(int idFicha, int idUsuario, string usuario, string nivel)
        {
            string query = @"INSERT INTO public.""Trazabilidad_Ediciones_Ficha"" 
        (""ID_Ficha"", ""ID_Usuario"", ""Usuario_Edito"", ""Nivel_Usuario"", ""Fecha_Edicion"")
        VALUES (@id_ficha, @id_usuario, @usuario, @nivel, @fecha)
        RETURNING ""ID_Trazabilidad"";";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
        new NpgsqlParameter("@id_usuario", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
        new NpgsqlParameter("@usuario", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = usuario ?? (object)DBNull.Value },
        new NpgsqlParameter("@nivel", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = nivel ?? (object)DBNull.Value },
        new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Timestamp) { Value = DateTime.Now }
            };

            try
            {
                return ExecuteScalarInt(query, parameters);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al registrar trazabilidad: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// Obtiene la trazabilidad de ediciones de una ficha específica
        /// </summary>
        public DataTable ObtenerTrazabilidadPorFicha(int idFicha)
        {
            string query = @"SELECT 
        t.""ID_Trazabilidad"",
        t.""Fecha_Edicion"",
        t.""Usuario_Edito"",
        t.""Nivel_Usuario"",
        u.""No_Empleado""
    FROM public.""Trazabilidad_Ediciones_Ficha"" t
    INNER JOIN public.""Usuarios"" u ON t.""ID_Usuario"" = u.""ID_User""
    WHERE t.""ID_Ficha"" = @id_ficha
    ORDER BY t.""Fecha_Edicion"" DESC";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha }
            };

            return ExecuteSelectQuery(query, parameters);
        }

        /// <summary>
        /// Obtiene toda la trazabilidad con filtros (versión simplificada)
        /// </summary>
        public DataTable ObtenerTrazabilidadConFiltros(string fechaInicio, string fechaFin, string usuario)
        {
            string query = @"SELECT 
        t.""ID_Trazabilidad"",
        t.""ID_Ficha"",
        t.""Fecha_Edicion"",
        t.""Usuario_Edito"",
        t.""Nivel_Usuario"",
        u.""No_Empleado""
    FROM public.""Trazabilidad_Ediciones_Ficha"" t
    INNER JOIN public.""Usuarios"" u ON t.""ID_Usuario"" = u.""ID_User""
    WHERE 1=1";

            List<NpgsqlParameter> parameters = new List<NpgsqlParameter>();

            if (!string.IsNullOrEmpty(fechaInicio) && !string.IsNullOrEmpty(fechaFin))
            {
                query += " AND t.\"Fecha_Edicion\" BETWEEN @fecha_inicio AND @fecha_fin";
                parameters.Add(new NpgsqlParameter("@fecha_inicio", NpgsqlTypes.NpgsqlDbType.Timestamp) { Value = Convert.ToDateTime(fechaInicio) });
                parameters.Add(new NpgsqlParameter("@fecha_fin", NpgsqlTypes.NpgsqlDbType.Timestamp) { Value = Convert.ToDateTime(fechaFin) });
            }

            if (!string.IsNullOrEmpty(usuario))
            {
                query += " AND t.\"Usuario_Edito\" ILIKE @usuario";
                parameters.Add(new NpgsqlParameter("@usuario", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = $"%{usuario}%" });
            }

            query += " ORDER BY t.\"Fecha_Edicion\" DESC";

            return ExecuteSelectQuery(query, parameters.ToArray());
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
        // Método para validar la contraseña de un Jefe de Turno y obtener su ID
        public int ValidateJefePasswordAndGetId(string identificador, string password, out string errorMessage)
        {
            errorMessage = string.Empty;

            string query = @"SELECT ""ID_User"" FROM public.""Usuarios"" 
                     WHERE (""Usuario"" ILIKE @identificador OR ""No_Empleado"" = @identificador) 
                     AND ""Password"" = @password 
                     AND ""Nivel"" = 'Jefe de Turno'
                     AND ""Activo"" = true";

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

                        if (result != null)
                        {
                            return Convert.ToInt32(result); // Retorna el ID del Jefe de Turno
                        }
                        else
                        {
                            return -1; // Credenciales incorrectas o no es Jefe de Turno
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return -2; // Error de conexión o BD
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

        // Nueva versión con manejo de error de conexión
        //public bool ValidateUser(string identificador, string password, out string errorMessage)
        //{
        //    errorMessage = string.Empty;

        //    // Consulta que acepta tanto Usuario como No_Empleado
        //    string query = @"SELECT COUNT(1) FROM public.""Usuarios"" 
        //   WHERE (""Usuario"" ILIKE @identificador OR ""No_Empleado"" = @identificador) 
        //   AND ""Password"" = @password AND ""Activo"" = true";

        //    NpgsqlParameter[] parameters = new NpgsqlParameter[]
        //    {
        //new NpgsqlParameter("@identificador", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = identificador },
        //new NpgsqlParameter("@password", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = password }
        //    };

        //    try
        //    {
        //        using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
        //        {
        //            connection.Open();
        //            using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
        //            {
        //                command.Parameters.AddRange(parameters);

        //                object result = command.ExecuteScalar();
        //                return result != null ? Convert.ToInt32(result) > 0 : false;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        if (IsNetworkError(ex))
        //        {
        //            errorMessage = "CONNECTION_ERROR";
        //        }
        //        else
        //        {
        //            errorMessage = ex.Message;
        //        }
        //        return false;
        //    }
        //}
        // Nueva versión que retorna un código de estado en lugar de solo bool
        public int ValidateUserWithStatus(string identificador, string password, out string errorMessage)
        {
            errorMessage = string.Empty;

            // Primero verificar si el usuario existe
            string queryCheckUser = @"SELECT ""Activo"" FROM public.""Usuarios"" 
           WHERE (""Usuario"" ILIKE @identificador OR ""No_Empleado"" = @identificador)";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@identificador", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = identificador }
            };

            try
            {
                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();

                    // Verificar si el usuario existe
                    using (NpgsqlCommand command = new NpgsqlCommand(queryCheckUser, connection))
                    {
                        command.Parameters.AddRange(parameters);
                        object result = command.ExecuteScalar();

                        // Si no existe el usuario
                        if (result == null)
                        {
                            return -1; // Usuario no existe
                        }

                        // Verificar si está activo
                        bool activo = Convert.ToBoolean(result);
                        if (!activo)
                        {
                            return -2; // Usuario desactivado
                        }
                    }

                    // Si existe y está activo, verificar la contraseña
                    string queryValidate = @"SELECT COUNT(1) FROM public.""Usuarios"" 
                   WHERE (""Usuario"" ILIKE @identificador OR ""No_Empleado"" = @identificador) 
                   AND ""Password"" = @password
                   AND ""Activo"" = true";

                    using (NpgsqlCommand command = new NpgsqlCommand(queryValidate, connection))
                    {
                        command.Parameters.AddRange(new NpgsqlParameter[]
                        {
                    new NpgsqlParameter("@identificador", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = identificador },
                    new NpgsqlParameter("@password", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = password }
                        });

                        object result = command.ExecuteScalar();
                        int count = result != null ? Convert.ToInt32(result) : 0;

                        if (count > 0)
                        {
                            return 1; // Usuario válido y activo
                        }
                        else
                        {
                            return 0; // Credenciales incorrectas
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
                return -3; // Error de conexión o excepción
            }
        }

        // Mantén el método original para compatibilidad si es necesario
        public bool ValidateUser(string identificador, string password, out string errorMessage)
        {
            int status = ValidateUserWithStatus(identificador, password, out errorMessage);
            return status == 1;
        }
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
        // Versión más genérica para cualquier cantidad de columnas
        public void LoadDataIntoMultiColumnComboBox(string query, RadMultiColumnComboBox multiColumnCombo,
            string[] columnNames, string[] columnHeaders, bool[] columnVisibility, NpgsqlParameter[] parameters = null)
        {
            try
            {
                DataTable dataTable = ExecuteSelectQuery(query, parameters);

                if (dataTable != null && dataTable.Rows.Count > 0)
                {
                    // Asignar el DataSource
                    multiColumnCombo.DataSource = dataTable;

                    // Configurar el editor (GridView)
                    var gridView = multiColumnCombo.EditorControl as RadGridView;
                    if (gridView != null)
                    {
                        // Limpiar columnas existentes
                        gridView.Columns.Clear();

                        // Agregar columnas según la configuración
                        for (int i = 0; i < columnNames.Length; i++)
                        {
                            if (i >= columnNames.Length) break;

                            GridViewTextBoxColumn column = new GridViewTextBoxColumn(columnNames[i]);
                            column.FieldName = columnNames[i];
                            column.HeaderText = (columnHeaders != null && i < columnHeaders.Length)
                                ? columnHeaders[i]
                                : columnNames[i];
                            column.IsVisible = (columnVisibility != null && i < columnVisibility.Length)
                                ? columnVisibility[i]
                                : true;
                            column.HeaderTextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
                            column.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;

                            gridView.Columns.Add(column);
                        }

                        // Configurar opciones de visualización
                        gridView.AutoSizeColumnsMode = GridViewAutoSizeColumnsMode.Fill;
                        gridView.EnableFiltering = true;
                        gridView.ShowHeaderCellButtons = true;
                        gridView.ReadOnly = true;
                    }

                    // Configurar propiedades del combobox
                    if (columnNames.Length > 0)
                    {
                        multiColumnCombo.DisplayMember = columnNames[1]; // La segunda columna como display (Lote)
                        multiColumnCombo.ValueMember = columnNames[0]; // La primera columna como valor (ID)
                    }

                    multiColumnCombo.DropDownMaxSize = new Size(400, 300);
                    multiColumnCombo.DropDownSizingMode = Telerik.WinControls.UI.SizingMode.RightBottom;
                    multiColumnCombo.AutoFilter = true;
                    multiColumnCombo.AutoSizeDropDownToBestFit = true;
                }
                else
                {
                    multiColumnCombo.DataSource = null;
                    multiColumnCombo.Text = "No hay datos disponibles";

                    var gridView = multiColumnCombo.EditorControl as RadGridView;
                    gridView?.Columns.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar RadMultiColumnComboBox: {ex.Message}");
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

        /// <summary>
        /// Registra una imagen de evidencia en la base de datos
        /// </summary>
        public int RegistrarImagenEvidencia(int idFicha, string nombreArchivo, string rutaCompleta,
            int numeroImagen, int idUsuario, long tamaño)
        {
            string query = @"INSERT INTO public.""Imagenes_Evidencia"" 
        (""ID_Ficha"", ""Nombre_Archivo"", ""Ruta_Completa"", ""Numero_Imagen"", 
         ""Fecha_Subida"", ""ID_Usuario"", ""Tamaño"")
        VALUES (@id_ficha, @nombre_archivo, @ruta_completa, @numero_imagen, 
                @fecha_subida, @id_usuario, @tamaño)
        RETURNING ""ID_Imagen"";";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
        new NpgsqlParameter("@nombre_archivo", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = nombreArchivo },
        new NpgsqlParameter("@ruta_completa", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = rutaCompleta },
        new NpgsqlParameter("@numero_imagen", NpgsqlTypes.NpgsqlDbType.Integer) { Value = numeroImagen },
        new NpgsqlParameter("@fecha_subida", NpgsqlTypes.NpgsqlDbType.Timestamp) { Value = DateTime.Now },
        new NpgsqlParameter("@id_usuario", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
        new NpgsqlParameter("@tamaño", NpgsqlTypes.NpgsqlDbType.Bigint) { Value = tamaño }
            };

            try
            {
                return ExecuteScalarInt(query, parameters);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al registrar imagen: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// Obtiene todas las imágenes de evidencia de una ficha
        /// </summary>
        public DataTable ObtenerImagenesPorFicha(int idFicha)
        {
            string query = @"SELECT 
        ""ID_Imagen"",
        ""Nombre_Archivo"",
        ""Ruta_Completa"",
        ""Numero_Imagen"",
        ""Fecha_Subida"",
        ""Tamaño""
    FROM public.""Imagenes_Evidencia""
    WHERE ""ID_Ficha"" = @id_ficha
    ORDER BY ""Numero_Imagen"" ASC";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha }
            };

            return ExecuteSelectQuery(query, parameters);
        }

        /// <summary>
        /// Elimina imágenes de evidencia de una ficha
        /// </summary>
        public bool EliminarImagenesPorFicha(int idFicha)
        {
            string query = @"DELETE FROM public.""Imagenes_Evidencia"" 
        WHERE ""ID_Ficha"" = @id_ficha";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha }
            };

            try
            {
                int result = ExecuteNonQuery(query, parameters);
                return result >= 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al eliminar imágenes: {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// Elimina una imagen específica por su ID
        /// </summary>
        public bool EliminarImagenPorId(int idImagen, out string rutaArchivo)
        {
            rutaArchivo = string.Empty;

            // Primero obtener la ruta del archivo
            string queryGet = @"SELECT ""Ruta_Completa"" FROM public.""Imagenes_Evidencia"" 
                        WHERE ""ID_Imagen"" = @id_imagen";

            NpgsqlParameter[] paramGet = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_imagen", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idImagen }
            };

            DataTable dt = ExecuteSelectQuery(queryGet, paramGet);
            if (dt != null && dt.Rows.Count > 0)
            {
                rutaArchivo = dt.Rows[0]["Ruta_Completa"].ToString();
            }

            // Eliminar el registro de la BD
            string queryDelete = @"DELETE FROM public.""Imagenes_Evidencia"" 
                           WHERE ""ID_Imagen"" = @id_imagen";

            NpgsqlParameter[] paramDelete = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_imagen", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idImagen }
            };

            try
            {
                int result = ExecuteNonQuery(queryDelete, paramDelete);
                return result > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al eliminar imagen: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Obtiene la ruta de una imagen por su ID
        /// </summary>
        public string ObtenerRutaImagenPorId(int idImagen)
        {
            string query = @"SELECT ""Ruta_Completa"" FROM public.""Imagenes_Evidencia"" 
                     WHERE ""ID_Imagen"" = @id_imagen";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_imagen", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idImagen }
            };

            DataTable dt = ExecuteSelectQuery(query, parameters);
            if (dt != null && dt.Rows.Count > 0)
            {
                return dt.Rows[0]["Ruta_Completa"].ToString();
            }
            return string.Empty;
        }
    }
}