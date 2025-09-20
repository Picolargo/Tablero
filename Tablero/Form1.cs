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
        private string id_global_meta_deshidratado = string.Empty;
        private bool filtroUsuariosActivo = false;
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
            dgv_metas_des.EnableHeadersVisualStyles = false;
            dgv_metas_des.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_des.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_metas_des.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_metas_des.BackgroundColor = Color.White; // Fondo blanco
            dgv_metas_des.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_metas_des.DefaultCellStyle.ForeColor = Color.Black;
            dgv_metas_des.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_metas_des.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_metas_des.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_metas_des.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_des.RowHeadersVisible = false;
            dgv_metas_des.BorderStyle = BorderStyle.None;
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
            actualiza_grid_Deshitratado();

        }
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

        private void actualiza_grid_Deshitratado()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""No_box_hr"" as ""No de Cajones/Hr"", 
                            ""Kg_fresco_hr"" as ""Kg Fresco/Hr"", 
                            ""Relacion_fr_seco"" as ""Relación Fresco-Seco"",
                            ""Kg_seco_hr"" as ""Kg Seco/Hr"",
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Deshidratado""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_des, null);

            // Configurar el DataGridView
            dgv_metas_des.Columns[0].Visible = false;
            dgv_metas_des.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void cb_Area_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex==0)
            {
                reiniciarCampos();
                //hacer visible para tunel
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;

                //habilitar controles
                cb_Turno.Enabled = true;
                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Lote";
                Txt_2.EmbeddedLabelText = "Kg Entrada (Proceso)";
                Txt_3.EmbeddedLabelText = "Merma Canica";
                Txt_4.EmbeddedLabelText = "Merma Podrido";
                Txt_5.EmbeddedLabelText = "Merma de Tina";
                Txt_6.EmbeddedLabelText = "Merma de Piso";
                Txt_7.EmbeddedLabelText = "Merma Canaletas";
                Txt_8.EmbeddedLabelText = "Merma Lavado de Bandas";
                Txt_9.EmbeddedLabelText = "Personal Operativo";
                Txt_10.EmbeddedLabelText = "Cascara y Carrete";

                Txt_Read_1.EmbeddedLabelText = "Horas Disponibles";
                Txt_Read_2.EmbeddedLabelText = "Meta Programada";
                Txt_Read_3.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_4.EmbeddedLabelText = "Kg Frescos de Entrada a secador";
                Txt_Read_5.EmbeddedLabelText = "Kg Secos Meta";

                //hacer invisibles controles
                Txt_11.Visible = false;

                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;

                //limpiar campos
                cb_Turno.SelectedIndex = -1;
                cb_OP.SelectedIndex = -1;
                dtp1.Value = DateTime.Now;
                Txt_1.Text = string.Empty;
                Txt_2.Text = string.Empty;
                Txt_3.Text = string.Empty;
                Txt_4.Text = string.Empty;
                Txt_5.Text = string.Empty;
                Txt_6.Text = string.Empty;
                Txt_7.Text = string.Empty;
                Txt_8.Text = string.Empty;
                Txt_9.Text = string.Empty;
                Txt_10.Text = string.Empty;
                Txt_meta.Text = string.Empty;

                Txt_Read_1.Text = string.Empty;
                Txt_Read_2.Text = string.Empty;
                Txt_Read_3.Text = string.Empty;
                Txt_Read_4.Text = string.Empty;
                Txt_Read_5.Text = string.Empty;

            }
            if (cb_Area.SelectedIndex == 1)
            {
                //Despegue
                reiniciarCampos();

                //hacer visible para Despegue
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Lote";
                Txt_2.EmbeddedLabelText = "Kilos Producto Seco";
                Txt_3.EmbeddedLabelText = "Merma Kg Secos";
                Txt_4.EmbeddedLabelText = "Kg Secos Fuera de Especificación";
                Txt_5.EmbeddedLabelText = "Kg para Resecar";
                Txt_6.EmbeddedLabelText = "Personal Operativo";
            }
        }

        private void reiniciarCampos()
        {
            cb_Area.SelectedIndex = -1;
            cb_Turno.SelectedIndex = -1;
            cb_OP.SelectedIndex = -1;
            dtp1.Value = DateTime.Now;
            Txt_1.Text = string.Empty;
            Txt_2.Text = string.Empty;
            Txt_3.Text = string.Empty;
            Txt_4.Text = string.Empty;
            Txt_5.Text = string.Empty;
            Txt_6.Text = string.Empty;
            Txt_7.Text = string.Empty;
            Txt_8.Text = string.Empty;
            Txt_9.Text = string.Empty;
            Txt_10.Text = string.Empty;
            Txt_11.Text = string.Empty;
            Txt_meta.Text = string.Empty;
            Txt_Read_1.Text = string.Empty;
            Txt_Read_2.Text = string.Empty;
            Txt_Read_3.Text = string.Empty;
            Txt_Read_4.Text = string.Empty;
            Txt_Read_5.Text = string.Empty;
            Txt_Read_6.Text = string.Empty;
            Txt_Read_7.Text = string.Empty;
            Txt_Read_8.Text = string.Empty;
            Txt_Read_9.Text = string.Empty;
            //hacer invisibles controles
            card_datos.Visible = false;
            card_TM.Visible = false;
            card_botones.Visible = false;
            //deshabilitar controles
            cb_Turno.Enabled = false;
            cb_OP.Enabled = false;
            dtp1.Enabled = false;
        }

        private void dgv_users_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Obtener el valor de la primera columna del renglón clicado
                id_global_users = dgv_users.Rows[e.RowIndex].Cells[0].Value.ToString();
                // Cargar los datos del usuario seleccionado en los campos de texto
                txt_no_emp.Text = dgv_users.Rows[e.RowIndex].Cells[1].Value.ToString(); // No Empleado
                txt_usuario.Text = dgv_users.Rows[e.RowIndex].Cells[2].Value.ToString(); // Nombre de usuario
                txt_contra.Text = dgv_users.Rows[e.RowIndex].Cells[3].Value.ToString(); // Contraseña
                cmb_nivel_user.SelectedItem = dgv_users.Rows[e.RowIndex].Cells[4].Value.ToString(); // Nivel

                btn_edit.Enabled = true;
                btn_delete_user.Enabled = true;
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
                    //MetroFramework.MetroMessageBox.Show(this, "Usuario guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    actualiza_grid_users(); // Actualizar el DataGridView de usuarios
                    limpiarCampos(); // Limpiar los campos de texto
                    btn_save.Enabled = false;
                    btn_cancel.Enabled = false;
                    txt_no_emp.Enabled = false;
                    txt_usuario.Enabled = false;
                    txt_contra.Enabled = false;
                    btn_edit.Enabled = false;
                    id_global_users = string.Empty;

                    cmb_nivel_user.Enabled = true;
                    cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
                    cmb_nivel_user.Enabled = false;
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
        private void limpiarCampos_meta()
        {
            cmb_area.SelectedIndex = -1;
            cmb_proceso.SelectedIndex = -1;
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;
            id_global_meta_deshidratado = string.Empty; // Limpiar la variable global
            cmb_area.Focus(); 
            cmb_proceso.Focus();
            txt_op.Focus();
        }

        private void btn_new_user_Click(object sender, EventArgs e)
        {
            // Obtener el valor máximo de la columna "No Empleado" (índice 1)
            int maxNoEmp = 0;
            foreach (DataGridViewRow row in dgv_users.Rows)
            {
                if (row.IsNewRow) continue; // Ignorar la fila para agregar nuevo
                int value;
                if (int.TryParse(row.Cells[1].Value?.ToString(), out value))
                {
                    if (value > maxNoEmp)
                        maxNoEmp = value;
                }
            }
            txt_no_emp.Text = (maxNoEmp + 1).ToString();

            btn_save.Enabled = true;
            btn_cancel.Enabled = true;
            btn_edit.Enabled = false;
            btn_delete_user.Enabled = false;

            txt_no_emp.Enabled = true;
            txt_usuario.Enabled = true;
            txt_contra.Enabled = true;
            cmb_nivel_user.Enabled = true;

            
            txt_usuario.Text = string.Empty;
            txt_contra.Text = string.Empty;
            cmb_nivel_user.SelectedIndex = -1;
            id_global_users = string.Empty; // Limpiar la variable global
            cb_Area.SelectedIndex = -1; // Limpiar el ComboBox de áreas
            cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
            txt_usuario.Focus(); // Enfocar el campo de No Empleado
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = false;
            btn_cancel.Enabled = false;
            btn_delete_user.Enabled = true;
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
            btn_delete_user.Enabled = false;
            txt_usuario.Enabled = true;
            txt_contra.Enabled = true;
            cmb_nivel_user.Enabled = true;
            cmb_nivel_user.Focus();
            txt_usuario.Focus();
        }

        private void txt_no_emp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void btn_search_user_Click(object sender, EventArgs e)
        {
            
            // Si el filtro está activo, limpiar filtro y campos de búsqueda
            if (filtroUsuariosActivo)
            {
                // Mostrar todas las filas del DataGridView
                foreach (DataGridViewRow row in dgv_users.Rows)
                {
                    row.Visible = true;
                }

                // Restablecer el estado del filtro
                filtroUsuariosActivo = false;
            }

            string noEmp = txt_search_no_emp.Text.Trim();
            string nomEmp = txt_search_nom_emp.Text.Trim();
            string nivel = cmb_serch_nivel.Text.Trim();

            // Validar que al menos un campo esté lleno
            if (string.IsNullOrEmpty(noEmp) && string.IsNullOrEmpty(nomEmp) && string.IsNullOrEmpty(nivel))
            {
                MetroFramework.MetroMessageBox.Show(this, "Favor de llenar uno o más campos para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Deseleccionar cualquier celda y fila antes de filtrar
            dgv_users.ClearSelection();
            dgv_users.CurrentCell = null;

            // Mover el CurrencyManager a una posición válida
            CurrencyManager cm = (CurrencyManager)BindingContext[dgv_users.DataSource];
            if (cm != null && cm.Count > 0)
                cm.Position = -1;

            bool hayFiltro = false;
            foreach (DataGridViewRow row in dgv_users.Rows)
            {
                if (row.IsNewRow) continue;

                bool match = true;

                if (!string.IsNullOrEmpty(noEmp))
                {
                    match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(noEmp, StringComparison.OrdinalIgnoreCase);
                }
                if (!string.IsNullOrEmpty(nomEmp))
                {
                    match &= row.Cells[2].Value != null && row.Cells[2].Value.ToString().IndexOf(nomEmp, StringComparison.OrdinalIgnoreCase) >= 0;
                }
                if (!string.IsNullOrEmpty(nivel))
                {
                    match &= row.Cells[4].Value != null && row.Cells[4].Value.ToString().Equals(nivel, StringComparison.OrdinalIgnoreCase);
                }

                row.Visible = match;
                if (match) hayFiltro = true;
            }

            filtroUsuariosActivo = hayFiltro;
        }

        private void btn_search_limpiar_Click(object sender, EventArgs e)
        {
            limpiar_filtros();
        }

        private void limpiar_filtros()
        {
            // Mostrar todas las filas del DataGridView
            foreach (DataGridViewRow row in dgv_users.Rows)
            {
                row.Visible = true;
            }

            // Limpiar los campos de búsqueda
            txt_search_no_emp.Text = string.Empty;
            txt_search_nom_emp.Text = string.Empty;
            cmb_serch_nivel.SelectedIndex = -1;

            // Restablecer el estado del filtro
            filtroUsuariosActivo = false;
        }

        private void txt_search_no_emp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void cmb_area_SelectedIndexChanged(object sender, EventArgs e)
        {
            reiniciar_controles_meta();
            if (cmb_area.SelectedIndex==0)
            {
                //habilitar controles
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_4.Enabled = true;
                txt_Meta_5.Enabled = true;

                //nombrar controles
                txt_Meta_1.Hint = "No. Cajones por Hr";
                txt_Meta_2.Hint = "Kg Fresco por Hr";
                txt_Meta_3.Hint = "Relación Fresco-Seco";
                txt_Meta_4.Hint = "Kg Seco/Hr";
                txt_Meta_5.Hint = "Personal Idóneo";

                tap_control_metas.SelectedIndex = 0;
            }
        }

        private void reiniciar_controles_meta()
        {
            txt_Meta_1.Hint = string.Empty;
            txt_Meta_2.Hint = string.Empty;
            txt_Meta_3.Hint = string.Empty;
            txt_Meta_4.Hint = string.Empty;
            txt_Meta_5.Hint = string.Empty;
            //deshabilitar controles
            txt_op.Enabled = false;
            txt_Meta_1.Enabled = false;
            txt_Meta_2.Enabled = false;
            txt_Meta_3.Enabled = false;
            txt_Meta_4.Enabled = false;
            txt_Meta_5.Enabled = false;
        }

        private void btn_new_op_Click(object sender, EventArgs e)
        {
            cmb_area.Enabled = true;
            btn_meta_cancel.Enabled = true;
            btn_meta_save.Enabled = true;
            btn_meta_edit.Enabled = false;
            btn_meta_delete.Enabled = false;

            cmb_area.SelectedIndex = -1;
            cmb_proceso.SelectedIndex = -1;
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;

            cmb_area.Focus();
            cmb_proceso.Focus();
            txt_op.Focus();
        }

        private void btn_meta_save_Click(object sender, EventArgs e)
        {
            if(cmb_area.SelectedIndex == -1) 
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if(cmb_area.Text == "Deshidratado" && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty || txt_Meta_4.Text == string.Empty || txt_Meta_5.Text == string.Empty)) 
            { 
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.Text == "Deshidratado")
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_deshidratado))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_deshidratado);

                    queryInsertUpdate = "UPDATE public.\"Deshidratado\" SET \"OP\" = @OP, \"No_box_hr\" = @no_box_hr, \"Kg_fresco_hr\" = @kg_f_h, \"Relacion_fr_seco\" = @relacion_fs, \"Kg_seco_hr\" = @kg_s_h, \"Personal_idoneo\" = @personal_i WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper()
                        },
                        new NpgsqlParameter("@no_box_hr", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = Convert.ToInt32(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@kg_f_h", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = Convert.ToInt32(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@relacion_fs", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = Convert.ToInt32(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@kg_s_h", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = Convert.ToInt32(txt_Meta_4.Text)
                        },
                        new NpgsqlParameter("@personal_i", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = Convert.ToInt32(txt_Meta_5.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {

                    // Verificar si el usuario ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Deshidratado\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text),
                    };
                    DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Deshidratado\" (\"OP\", \"No_box_hr\", \"Kg_fresco_hr\", \"Relacion_fr_seco\", \"Kg_seco_hr\", \"Personal_idoneo\") VALUES (@OP, @No_box_h, @Kg_f_h, @Relacion_f_s, @kg_s_h, @Personal_i);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper()),
                    new NpgsqlParameter("@No_box_h", Convert.ToInt32(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Kg_f_h", Convert.ToInt32(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Relacion_f_s", Convert.ToInt32(txt_Meta_3.Text)),
                    new NpgsqlParameter("@kg_s_h", Convert.ToInt32(txt_Meta_4.Text)),
                    new NpgsqlParameter("@Personal_i", Convert.ToInt32(txt_Meta_5.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_Deshitratado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    txt_Meta_4.Enabled = false;
                    txt_Meta_5.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_users = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;

                    cmb_proceso.Enabled = true;
                    cmb_proceso.Focus();
                    cmb_proceso.Enabled = false;
                }
            }
        }

        private void dgv_metas_des_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                id_global_meta_deshidratado = dgv_metas_des.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_des.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_des.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_2.Text = dgv_metas_des.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_3.Text = dgv_metas_des.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_4.Text = dgv_metas_des.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_5.Text = dgv_metas_des.Rows[e.RowIndex].Cells[6].Value.ToString();

                

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 0;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                txt_Meta_4.Enabled = false;
                txt_Meta_5.Enabled = false;
            }
        }

        private void txt_Meta_1_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void txt_Meta_1_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MessageBox.Show("Formato inválido. Solo se permiten números con un punto decimal.");
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void txt_Meta_2_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void txt_Meta_2_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MessageBox.Show("Formato inválido. Solo se permiten números con un punto decimal.");
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void txt_Meta_3_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void txt_Meta_3_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MessageBox.Show("Formato inválido. Solo se permiten números con un punto decimal.");
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void txt_Meta_4_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void txt_Meta_4_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MessageBox.Show("Formato inválido. Solo se permiten números con un punto decimal.");
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void txt_Meta_5_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void txt_Meta_5_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MessageBox.Show("Formato inválido. Solo se permiten números con un punto decimal.");
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void btn_meta_edit_Click(object sender, EventArgs e)
        {
            btn_meta_save.Enabled = true;
            btn_meta_cancel.Enabled = true;
            btn_meta_delete.Enabled = false;
            txt_op.Enabled = true;
            txt_Meta_1.Enabled = true;
            txt_Meta_2.Enabled = true;
            txt_Meta_3.Enabled = true;
            txt_Meta_4.Enabled = true;
            txt_Meta_5.Enabled = true;
            txt_op.Focus();
        }

        private void btn_meta_cancel_Click(object sender, EventArgs e)
        {
            btn_meta_save.Enabled = false;
            btn_meta_cancel.Enabled = false;
            btn_meta_delete.Enabled = true;
            limpiarCampos_meta();
            cmb_area.Enabled = false;
            txt_op.Enabled = false;
            txt_Meta_1.Enabled = false;
            txt_Meta_2.Enabled = false;
            txt_Meta_3.Enabled = false;
            txt_Meta_4.Enabled = false;
            txt_Meta_5.Enabled = false;
            btn_meta_edit.Enabled = false;
            id_global_meta_deshidratado = string.Empty;
        }

        private void btn_delete_user_Click(object sender, EventArgs e)
        {

        }
    }
}
