using MaterialSkin;
using MaterialSkin.Controls;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tablero
{
    public partial class Form_principal : MaterialForm
    {
        private string var_no_empledo = string.Empty;
        private string var_nom_empledo = string.Empty;
        private string id_global_users = string.Empty; // Variable para almacenar el ID del usuario seleccionado en el DataGridView
        //variable para la conexión a la base de datos
        string connectionString = "Host=localhost;Username=postgres;Password=Picolargo789;Database=Reporteo";

        public Form_principal(string var_no_empledo, string var_nom_empledo)
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
            this.var_no_empledo = var_no_empledo;// Asignar el número de empleado al campo correspondiente
            this.var_nom_empledo = var_nom_empledo;// Asignar el nombre del empleado al campo correspondiente

            lbl_no_emp2.Text = var_no_empledo; // Mostrar el número de empleado en el label correspondiente
            lbl_nom2.Text = var_nom_empledo; // Mostrar el nombre del empleado en el label correspondiente

            // Initialize MaterialSkinManager and set the theme and color scheme  
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Orange600, Primary.Orange600, Primary.BlueGrey800, Accent.Blue700, TextShade.WHITE);

            // Forzar el color después de inicializar MaterialSkin  
            pnl_hr_inicio.BackColor = Color.FromArgb(192, 255, 192);
            pnl_Metahora.BackColor = Color.FromArgb(192, 255, 192);


            // Personalización de dgv_mecanico
            dgv_mecanico.EnableHeadersVisualStyles = false;
            dgv_mecanico.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_mecanico.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_mecanico.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_mecanico.BackgroundColor = Color.White; // Fondo blanco
            dgv_mecanico.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_mecanico.DefaultCellStyle.ForeColor = Color.Black;
            dgv_mecanico.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_mecanico.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_mecanico.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_mecanico.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_mecanico.RowHeadersVisible = false;
            dgv_mecanico.AllowUserToAddRows = true;
            dgv_mecanico.AllowUserToDeleteRows = true;
            dgv_mecanico.AllowUserToResizeRows = false;
            dgv_mecanico.ReadOnly = false;
            dgv_mecanico.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dgv_mecanico.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv_mecanico.MultiSelect = false;
            dgv_mecanico.BorderStyle = BorderStyle.None;

            // Personalización de dgv_almacen
            dgv_almacen.EnableHeadersVisualStyles = false;
            dgv_almacen.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_almacen.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_almacen.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_almacen.BackgroundColor = Color.White; // Fondo blanco
            dgv_almacen.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_almacen.DefaultCellStyle.ForeColor = Color.Black;
            dgv_almacen.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_almacen.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_almacen.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_almacen.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_almacen.RowHeadersVisible = false;
            dgv_almacen.AllowUserToAddRows = true;
            dgv_almacen.AllowUserToDeleteRows = true;
            dgv_almacen.AllowUserToResizeRows = false;
            dgv_almacen.ReadOnly = false;
            dgv_almacen.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dgv_almacen.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv_almacen.MultiSelect = false;
            dgv_almacen.BorderStyle = BorderStyle.None;

            // Personalización de dgv_operativo
            dgv_operativo.EnableHeadersVisualStyles = false;
            dgv_operativo.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_operativo.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_operativo.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_operativo.BackgroundColor = Color.White; // Fondo blanco
            dgv_operativo.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_operativo.DefaultCellStyle.ForeColor = Color.Black;
            dgv_operativo.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_operativo.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_operativo.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_operativo.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_operativo.RowHeadersVisible = false;
            dgv_operativo.AllowUserToAddRows = true;
            dgv_operativo.AllowUserToDeleteRows = true;
            dgv_operativo.AllowUserToResizeRows = false;
            dgv_operativo.ReadOnly = false;
            dgv_operativo.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dgv_operativo.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv_operativo.MultiSelect = false;
            dgv_operativo.BorderStyle = BorderStyle.None;

            // Personalización de dgv_calidad
            dgv_calidad.EnableHeadersVisualStyles = false;
            dgv_calidad.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_calidad.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_calidad.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_calidad.BackgroundColor = Color.White; // Fondo blanco
            dgv_calidad.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_calidad.DefaultCellStyle.ForeColor = Color.Black;
            dgv_calidad.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_calidad.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_calidad.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_calidad.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_calidad.RowHeadersVisible = false;
            dgv_calidad.AllowUserToAddRows = true;
            dgv_calidad.AllowUserToDeleteRows = true;
            dgv_calidad.AllowUserToResizeRows = false;
            dgv_calidad.ReadOnly = false;
            dgv_calidad.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dgv_calidad.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv_calidad.MultiSelect = false;
            dgv_calidad.BorderStyle = BorderStyle.None;

            // Personalización de dgv_users
            dgv_users.EnableHeadersVisualStyles = false;
            dgv_users.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_users.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_users.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_users.BackgroundColor = Color.White; // Fondo blanco
            dgv_users.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_users.DefaultCellStyle.ForeColor = Color.Black;
            dgv_users.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_users.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_users.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_users.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_users.RowHeadersVisible = false;
            dgv_users.BorderStyle = BorderStyle.None;

            // Personalización de dgv_metas
            dgv_metas.EnableHeadersVisualStyles = false;
            dgv_metas.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_metas.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_metas.BackgroundColor = Color.White; // Fondo blanco
            dgv_metas.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_metas.DefaultCellStyle.ForeColor = Color.Black;
            dgv_metas.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_metas.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_metas.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_metas.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas.RowHeadersVisible = false;
            dgv_metas.BorderStyle = BorderStyle.None;
        }

        private void dgv_mecanico_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_mecanico.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Quita la restricción en otras columnas
                }
            }
        }

        private void SoloNumeros_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permite solo dígitos y teclas de control (como backspace)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void dgv_almacen_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_almacen.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Quita la restricción en otras columnas
                }
            }
        }

        private void dgv_operativo_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_operativo.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Quita la restricción en otras columnas
                }
            }
        }

        private void dgv_calidad_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_calidad.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Quita la restricción en otras columnas
                }
            }
        }

        private void txt_Tiempo_comida_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permite solo dígitos y teclas de control (como backspace)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void Form_principal_Load(object sender, EventArgs e)
        {
            lbl_user_no_emp.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
            lbl_Nom.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);

            actualiza_grid_users(); // Llamar al método para actualizar el DataGridView de usuarios

        }

        //private void actualiza_grid_users()
        //{
        //    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
        //    // string querySimple = "SELECT * FROM public.\"Usuarios\"";
        //    string querySimple = "SELECT \"ID_User\" as \"ID\", \"No_Empleado\" as \"No Empleado\", \"Usuario\" as \"Nombre de usuario\", \"Password\" as \"Contraseña\", \"Nivel\" FROM public.\"Usuarios\" ORDER BY \"\"No_Empleado\"\" ASC;";
        //    // Cargar los datos de la tabla Usuarios en el DataGridView
        //    dbHelper.LoadDataIntoDataGridView(querySimple, dgv_users, null);
        //    // Configurar el DataGridView
        //    dgv_users.Columns[0].Visible = false;
        //    dgv_users.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;// Ajustar automáticamente el ancho de la columna "No Empleado"
        //}
        private void actualiza_grid_users()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_User"" as ""ID"", 
                            ""No_Empleado"" as ""No Empleado"", 
                            ""Usuario"" as ""Nombre de usuario"", 
                            ""Password"" as ""Contraseña"", 
                            ""Nivel"" 
                          FROM public.""Usuarios""
                          ORDER BY ""No_Empleado"" ASC;";  // ← ASC para orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_users, null);

            // Configurar el DataGridView
            dgv_users.Columns[0].Visible = false;
            dgv_users.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void cb_Area_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                reiniciar_textboxes();// reiniciar los textbox
                txt_1.EmbeddedLabelText = "Relacion fresco seco meta";
                txt_2.EmbeddedLabelText = "Relacion fresco seco";
                txt_3.EmbeddedLabelText = "Kilogramos por hora hombre";
                txt_4.EmbeddedLabelText = "Kg a Túnel";
                txt_5.EmbeddedLabelText = "Kg entrada a sumergidor";
                txt_6.EmbeddedLabelText = "Kg en despegue";
                txt_7.EmbeddedLabelText = "Kg de repelado";
                //hacer visible los componentes
                txt_1.Visible = true;
                txt_2.Visible = true;
                txt_3.Visible = true;
                txt_4.Visible = true;
                txt_5.Visible = true;
                txt_6.Visible = true;
                txt_7.Visible = true;
                card_datos.Visible = true;
                card_TM.Visible = true;
                // hacer visible los controles de turno y fecha
                cb_Turno.Enabled = true;
                dtp1.Enabled = true;
            }
            if (cb_Area.SelectedIndex == 1)
            {
                reiniciar_textboxes();// reiniciar los textbox
                txt_1.EmbeddedLabelText = "Kilogramos OK";
                txt_4.EmbeddedLabelText = "Kg fuera de especificación";
                txt_5.EmbeddedLabelText = "Merma";
                txt_6.EmbeddedLabelText = "Lote 1";
                txt_7.EmbeddedLabelText = "Kg de lote 1";
                txt_8.EmbeddedLabelText = "Lote 2";
                txt_9.EmbeddedLabelText = "Kg de lote 2";
                // hacer visibles los componentes
                txt_1.Visible = true;
                txt_4.Visible = true;
                txt_5.Visible = true;
                txt_6.Visible = true;
                txt_7.Visible = true;
                txt_8.Visible = true;
                txt_9.Visible = true;
                card_datos.Visible = true;
                card_TM.Visible = true;
                // Habilitar los controles de turno y fecha
                cb_Turno.Enabled = true;
                dtp1.Enabled = true;
            }
            if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4)
            {
                reiniciar_textboxes();// reiniciar los textbox
                txt_1.EmbeddedLabelText = "Kilogramos OK";
                txt_4.EmbeddedLabelText = "Kg fuera de especificación";
                txt_5.EmbeddedLabelText = "Merma";
                // hacer visibles los componentes
                txt_1.Visible = true;
                txt_4.Visible = true;
                txt_5.Visible = true;
                card_datos.Visible = true;
                card_TM.Visible = true;
                // Habilitar los controles de turno y fecha
                cb_Turno.Enabled = true;
                dtp1.Enabled = true;
            }
            if (cb_Area.SelectedIndex == 5)
            {
                reiniciar_textboxes();// reiniciar los textbox
                txt_1.EmbeddedLabelText = "Kilogramos OK";
                txt_4.EmbeddedLabelText = "Kg fuera de especificación";
                txt_5.EmbeddedLabelText = "Merma";
                txt_6.EmbeddedLabelText = "Costales empacados";
                // hacer visibles los componentes
                txt_1.Visible = true;
                txt_4.Visible = true;
                txt_5.Visible = true;
                txt_6.Visible = true;
                card_datos.Visible = true;
                card_TM.Visible = true;
                // Habilitar los controles de turno y fecha
                cb_Turno.Enabled = true;
                dtp1.Enabled = true;
            }
            if (cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7 || cb_Area.SelectedIndex == 8)
            {
                reiniciar_textboxes();// reiniciar los textbox

                // Habilitar los controles
                cb_Turno.Enabled = true;
                dtp1.Enabled = true;
                cb_OP.Enabled = true;
            }
        }

        public void reiniciar_textboxes()
        {
            cb_Turno.Enabled = false;
            cb_OP.Enabled = false;
            dtp1.Enabled = false;
            // Reiniciar los textos de los textboxes
            txt_1.EmbeddedLabelText = string.Empty;
            txt_2.EmbeddedLabelText = string.Empty;
            txt_3.EmbeddedLabelText = string.Empty;
            txt_4.EmbeddedLabelText = string.Empty;
            txt_5.EmbeddedLabelText = string.Empty;
            txt_6.EmbeddedLabelText = string.Empty;
            txt_7.EmbeddedLabelText = string.Empty;
            txt_8.EmbeddedLabelText = string.Empty;
            txt_9.EmbeddedLabelText = string.Empty;
            // Limpiar los valores de los textboxes
            txt_1.Text = string.Empty;
            txt_2.Text = string.Empty;
            txt_3.Text = string.Empty;
            txt_4.Text = string.Empty;
            txt_5.Text = string.Empty;
            txt_6.Text = string.Empty;
            txt_7.Text = string.Empty;
            txt_8.Text = string.Empty;
            txt_9.Text = string.Empty;
            // Hacer invisibles los componentes
            txt_1.Visible = false;
            txt_2.Visible = false;
            txt_3.Visible = false;
            txt_4.Visible = false;
            txt_5.Visible = false;
            txt_6.Visible = false;
            txt_7.Visible = false;
            txt_8.Visible = false;
            txt_9.Visible = false;
            card_datos.Visible = false;
            card_TM.Visible = false;
        }

        private void dgv_users_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Obtener el valor de la primera columna del renglón clicado
                id_global_users = dgv_users.Rows[e.RowIndex].Cells[0].Value.ToString();
                // MetroFramework.MetroMessageBox.Show(this, "Usuario id", "id es: "+ id_global_users, MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Cargar los datos del usuario seleccionado en los campos de texto
                txt_no_emp.Text = dgv_users.Rows[e.RowIndex].Cells[1].Value.ToString(); // No Empleado
                txt_usuario.Text = dgv_users.Rows[e.RowIndex].Cells[2].Value.ToString(); // Nombre de usuario
                txt_contra.Text = dgv_users.Rows[e.RowIndex].Cells[3].Value.ToString(); // Contraseña
                cmb_nivel_user.SelectedItem = dgv_users.Rows[e.RowIndex].Cells[4].Value.ToString(); // Nivel

                btn_edit.Enabled = true;
                cmb_nivel_user.Enabled = true;
                cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
                cmb_nivel_user.Enabled = false;
            }
        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            if (txt_no_emp.Text == string.Empty || txt_usuario.Text == string.Empty || txt_contra.Text == string.Empty || cmb_nivel_user.SelectedIndex == -1)
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_users))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idUsuario = Convert.ToInt32(id_global_users);

                    queryInsertUpdate = "UPDATE public.\"Usuarios\" SET \"Usuario\" = @Usuario, \"Password\" = @Password, \"Nivel\" = @Nivel WHERE \"ID_User\" = @ID_usuario;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Usuario", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_usuario.Text
                        },
                        new NpgsqlParameter("@Password", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_contra.Text
                        },
                        new NpgsqlParameter("@Nivel", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = cmb_nivel_user.SelectedItem?.ToString() ?? ""
                        },
                        new NpgsqlParameter("@ID_usuario", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idUsuario  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    
                    // Verificar si el usuario ya existe
                    string queryCheckUser = "SELECT COUNT(*) FROM public.\"Usuarios\" WHERE \"Usuario\"  ILIKE @Usuario or \"No_Empleado\" ILIKE @no_emp;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@Usuario", txt_usuario.Text),
                    new NpgsqlParameter("@no_emp", txt_no_emp.Text)
                    };
                    DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckUser, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El nombre de usuario y/o numero de emplado ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Usuarios\" (\"No_Empleado\", \"Usuario\", \"Password\", \"Nivel\") VALUES (@No_Empleado, @Usuario, @Password, @Nivel);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@No_Empleado", txt_no_emp.Text),
                    new NpgsqlParameter("@Usuario", txt_usuario.Text),
                    new NpgsqlParameter("@Password", txt_contra.Text),
                    new NpgsqlParameter("@Nivel", cmb_nivel_user.SelectedItem.ToString())
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                
                if (result > 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Usuario guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    actualiza_grid_users(); // Actualizar el DataGridView de usuarios
                    limpiarCampos(); // Limpiar los campos de texto
                }
            }
        }

        private void limpiarCampos()
        {
            txt_no_emp.Text = string.Empty;
            txt_usuario.Text = string.Empty;
            txt_contra.Text = string.Empty;
            cmb_nivel_user.SelectedIndex = -1;
            id_global_users = string.Empty; // Limpiar la variable global
            cb_Area.SelectedIndex = -1; // Limpiar el ComboBox de áreas
            cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
            txt_no_emp.Focus(); // Enfocar el campo de No Empleado
        }

        private void btn_new_user_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = true;
            btn_cancel.Enabled = true;
            btn_edit.Enabled = false;
            limpiarCampos();
            txt_no_emp.Enabled = true;
            txt_usuario.Enabled = true;
            txt_contra.Enabled = true;
            cmb_nivel_user.Enabled = true;
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = false;
            btn_cancel.Enabled = false;
            limpiarCampos();
            txt_no_emp.Enabled = false;
            txt_usuario.Enabled = false;
            txt_contra.Enabled = false;
            btn_edit.Enabled = false;
            id_global_users = string.Empty;

            cmb_nivel_user.Enabled = true;
            cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
            cmb_nivel_user.Enabled = false;
        }

        private void btn_edit_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = true; 
            btn_cancel.Enabled = true;
            txt_usuario.Enabled = true;
            txt_contra.Enabled = true;
            cmb_nivel_user.Enabled = true;
            cmb_nivel_user.Focus();
            txt_usuario.Focus();
        }
    }
}
