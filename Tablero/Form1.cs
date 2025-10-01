using MaterialSkin;
using MaterialSkin.Controls;
using MetroFramework.Controls;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telerik.WinControls.UI;
using static System.Net.Mime.MediaTypeNames;

namespace Tablero
{
    public partial class Form_principal : MaterialForm
    {
        private string id_global_users = string.Empty; // Variable para almacenar el ID del usuario seleccionado en el DataGridView
        private string id_global_meta_deshidratado = string.Empty;
        private string id_global_meta_empacado = string.Empty;
        private string id_global_meta_inspec = string.Empty;
        private string id_global_meta_evaporado = string.Empty;
        private string id_global_meta_grind = string.Empty;
        private string id_global_meta_platinum = string.Empty;
        private int id_user = 0;
        private bool filtroUsuariosActivo = false;
        private bool filtroUsuariosActivo_OP = false;
        private int filtroUsuariosActivo_OP_dgv_activo = 0;
        private string horaInicioText = string.Empty;
        private string horaFinText = string.Empty;
        private DateTime horaInicio;
        private DateTime horaFin;
        //variable para la conexión a la base de datos
        string connectionString = string.Empty;

        public Form_principal(string var_no_empledo, string var_nom_empledo, int ID_usuario, string conexionstring)
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;

            lbl_no_emp2.Text = var_no_empledo; // Mostrar el número de empleado en el label correspondiente
            lbl_nom2.Text = var_nom_empledo.ToUpper(); // Mostrar el nombre del empleado en el label correspondiente
            connectionString = conexionstring; // Asignar la cadena de conexión pasada como parámetro
            id_user = ID_usuario; // Asignar el ID del usuario pasado como parámetro


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
            //dgv_mecanico.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
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
            //dgv_operativo.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
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

            // Personalización de dgv_metas_emp
            dgv_metas_emp.EnableHeadersVisualStyles = false;
            dgv_metas_emp.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_emp.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_metas_emp.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_metas_emp.BackgroundColor = Color.White; // Fondo blanco
            dgv_metas_emp.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_metas_emp.DefaultCellStyle.ForeColor = Color.Black;
            dgv_metas_emp.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_metas_emp.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_metas_emp.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_metas_emp.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_emp.RowHeadersVisible = false;
            dgv_metas_emp.BorderStyle = BorderStyle.None;

            // Personalización de dgv_metas_insp
            dgv_metas_insp.EnableHeadersVisualStyles = false;
            dgv_metas_insp.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_insp.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_metas_insp.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_metas_insp.BackgroundColor = Color.White; // Fondo blanco
            dgv_metas_insp.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_metas_insp.DefaultCellStyle.ForeColor = Color.Black;
            dgv_metas_insp.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_metas_insp.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_metas_insp.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_metas_insp.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_insp.RowHeadersVisible = false;
            dgv_metas_insp.BorderStyle = BorderStyle.None;

            // Personalización de dgv_metas_Eva
            dgv_metas_Eva.EnableHeadersVisualStyles = false;
            dgv_metas_Eva.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_Eva.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_metas_Eva.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_metas_Eva.BackgroundColor = Color.White; // Fondo blanco
            dgv_metas_Eva.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_metas_Eva.DefaultCellStyle.ForeColor = Color.Black;
            dgv_metas_Eva.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_metas_Eva.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_metas_Eva.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_metas_Eva.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_Eva.RowHeadersVisible = false;
            dgv_metas_Eva.BorderStyle = BorderStyle.None;

            // Personalización de dgv_metas_Grind
            dgv_metas_Grind.EnableHeadersVisualStyles = false;
            dgv_metas_Grind.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_Grind.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_metas_Grind.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_metas_Grind.BackgroundColor = Color.White; // Fondo blanco
            dgv_metas_Grind.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_metas_Grind.DefaultCellStyle.ForeColor = Color.Black;
            dgv_metas_Grind.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_metas_Grind.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_metas_Grind.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_metas_Grind.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_Grind.RowHeadersVisible = false;
            dgv_metas_Grind.BorderStyle = BorderStyle.None;

            // Personalización de dgv_metas_platinum
            dgv_metas_platinum.EnableHeadersVisualStyles = false;
            dgv_metas_platinum.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_platinum.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_metas_platinum.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_metas_platinum.BackgroundColor = Color.White; // Fondo blanco
            dgv_metas_platinum.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_metas_platinum.DefaultCellStyle.ForeColor = Color.Black;
            dgv_metas_platinum.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_metas_platinum.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_metas_platinum.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_metas_platinum.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas_platinum.RowHeadersVisible = false;
            dgv_metas_platinum.BorderStyle = BorderStyle.None;
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
            actualiza_grid_Empacado();
            actualiza_grid_inspec();
            actualiza_grid_evaporado();
            actualiza_grid_grind();
            actualiza_platinum();
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
                            ""No_box_hr"" as ""Numero de Cajones/Hora"", 
                            ""Kg_fresco_hr"" as ""Kilogramos Fresco/Hora"", 
                            ""Relacion_fr_seco"" as ""Relación Fresco-Seco"",
                            ""Kg_seco_hr"" as ""Kilogramos Seco/Hora"",
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Deshidratado""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_des, null);

            // Configurar el DataGridView
            dgv_metas_des.Columns[0].Visible = false;
            dgv_metas_des.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_grid_Empacado()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Kg_person_hr"" as ""Kilogramos/Persona por Hora"", 
                            ""Meta_kg_hr_line"" as ""Meta Kilogramos/Hora por 1 linea"", 
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Empacado""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_emp, null);

            // Configurar el DataGridView
            dgv_metas_emp.Columns[0].Visible = false;
            dgv_metas_emp.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_grid_inspec()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Kg_person_hr"" as ""Kilogramos/Persona por Hora"", 
                            ""Meta_kg_hr_line"" as ""Meta Kilogramos/Hora por 1 linea"", 
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Inspeccion""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_insp, null);

            // Configurar el DataGridView
            dgv_metas_insp.Columns[0].Visible = false;
            dgv_metas_insp.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_grid_evaporado()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Capacidad_trompos"" as ""Capacidad por Trompo"", 
                            ""Cantidad_trompos"" as ""Cantidad de Trompos"",
                            ""Meta_kg_hr"" as ""Meta Kg por hora"", 
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Evaporado""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_Eva, null);

            // Configurar el DataGridView
            dgv_metas_Eva.Columns[0].Visible = false;
            dgv_metas_Eva.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_grid_grind()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Capacidad_Molino"" as ""Capacidad por Molino"", 
                            ""Cantidad_molinos"" as ""Cantidad de Molinos"",
                            ""Meta_Kg_hr"" as ""Meta Kg por hora"", 
                            ""Personal_Idoneo"" as ""Personal Idóneo""
                          FROM public.""Grind""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_Grind, null);

            // Configurar el DataGridView
            dgv_metas_Grind.Columns[0].Visible = false;
            dgv_metas_Grind.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_platinum()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Cap_Trompo_machin"" as ""Capacidad por Trompo y/o Máquinas"", 
                            ""Canti_Trompo_machin"" as ""Cantidad de Trompos y/o Máquinas"",
                            ""Meta_Kg_Hr"" as ""Meta Kg por hora"", 
                            ""Personal_Idoneo"" as ""Personal Idóneo""
                          FROM public.""Platinum""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_platinum, null);

            // Configurar el DataGridView
            dgv_metas_platinum.Columns[0].Visible = false;
            dgv_metas_platinum.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
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
                card_meal_energy.Visible = true;

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

                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta Programada";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Kg Frescos de Entrada a secador";

                //hacer invisibles controles
                Txt_11.Visible = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;

                //limpiar campos
                dtp1.Value = DateTime.Now;
                Txt_1.Text = String.Empty;
                Txt_2.Text = "0";
                Txt_3.Text = "0";
                Txt_4.Text = "0";
                Txt_5.Text = "0";
                Txt_6.Text = "0";
                Txt_7.Text = "0";
                Txt_8.Text = "0";
                Txt_9.Text = "0";
                Txt_10.Text = "0";
                Txt_meta.Text = string.Empty;

                Txt_Read_1.Text = string.Empty;
                Txt_Read_2.Text = string.Empty;
                Txt_Read_3.Text = string.Empty;
                Txt_Read_4.Text = string.Empty;

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Deshidratado\" ORDER BY \"OP\";";
                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 1)
            {
                //Despegue
                reiniciarCampos();

                //hacer visible para Despegue
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;

                //habilitar controles
                cb_Turno.Enabled = true;
                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Lote";
                Txt_2.EmbeddedLabelText = "Kilos Producto Seco";
                Txt_3.EmbeddedLabelText = "Merma Kg Secos";
                Txt_4.EmbeddedLabelText = "Kg Secos Fuera de Especificación";
                Txt_5.EmbeddedLabelText = "Kg para Resecar";
                Txt_6.EmbeddedLabelText = "Personal Operativo";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Deshidratado\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 2)
            {
                //Despegue
                reiniciarCampos();

                //hacer visible para Despegue
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;

                //habilitar controles
                cb_Turno.Enabled = true;
                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Lote";
                Txt_2.EmbeddedLabelText = "Kilos Producto Seco";
                Txt_3.EmbeddedLabelText = "Merma Kg Secos";
                Txt_4.EmbeddedLabelText = "Kg Secos Fuera de Especificación";
                Txt_5.EmbeddedLabelText = "Kg para Resecar";
                Txt_6.EmbeddedLabelText = "Personal Operativo";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Evaporado\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
        }

        // Método que suma los valores de Txt_3 a Txt_8
        private void CalcularSuma()
        {
            decimal total = 0;
            decimal Kg_entrada = 0;
            // Lista de los TextBox que quieres sumar
            RadTextBox[] cajas = { Txt_3, Txt_4, Txt_5, Txt_6, Txt_7, Txt_8 };

            foreach (RadTextBox txt in cajas)
            {
                if (decimal.TryParse(txt.Text, out decimal valor))
                {
                    total += valor;
                }
            }
            Kg_entrada = Txt_2.Text == string.Empty ? 0 : Convert.ToDecimal(Txt_2.Text);
            decimal resultado = Kg_entrada - total;
            Txt_Read_4.Text = resultado.ToString("0.00"); // puedes usar "0.##" si quieres decimales
        }

        private void reiniciarCampos()
        {
            cb_Turno.SelectedIndex = -1;
            cb_Turno.Focus();
            cb_OP.SelectedIndex = -1;
            cb_OP.Focus();
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
            txt_Tiempo_comida.Text = "30";
            //hacer invisibles controles
            card_datos.Visible = false;
            card_TM.Visible = false;
            card_botones.Visible = false;
            card_meal_energy.Visible = false;
            //deshabilitar controles
            cb_Turno.Enabled = false;
            cb_OP.Enabled = false;
            dtp1.Enabled = false;
            txt_Tiempo_energia.Text = "0";
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
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;
            cmb_area.Focus(); 
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

            if (!hayFiltro)
            {
                MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                    "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
            cmb_serch_nivel.Focus();
            txt_search_no_emp.Focus();

            // Restablecer el estado del filtro
            filtroUsuariosActivo = false;
        }

        private void limpiar_filtros_OP()
        {
            //filtroUsuariosActivo_OP_dgv_activo
            // Mostrar todas las filas del DataGridView
            // Mostrar todas las filas del DataGridView Deshidratado
            if (filtroUsuariosActivo_OP_dgv_activo==1)
            {
                foreach (DataGridViewRow row in dgv_metas_des.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Empacado
            if (filtroUsuariosActivo_OP_dgv_activo == 2)
            {
                foreach (DataGridViewRow row in dgv_metas_emp.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Inspección
            if (filtroUsuariosActivo_OP_dgv_activo == 3)
            {
                foreach (DataGridViewRow row in dgv_metas_insp.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Evaporado
            if (filtroUsuariosActivo_OP_dgv_activo == 4)
            {
                foreach (DataGridViewRow row in dgv_metas_Eva.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Grind
            if (filtroUsuariosActivo_OP_dgv_activo == 5)
            {
                foreach (DataGridViewRow row in dgv_metas_Grind.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Platinum
            if (filtroUsuariosActivo_OP_dgv_activo == 6)
            {
                foreach (DataGridViewRow row in dgv_metas_platinum.Rows)
                {
                    row.Visible = true;
                }
            }

            // Limpiar los campos de búsqueda
            txt_OP_Search.Text = string.Empty;
            cmb_area_search.SelectedIndex = -1;
            cmb_area_search.Focus();

            txt_OP_Search.Enabled = false;

            // Restablecer el estado del filtro
            filtroUsuariosActivo_OP = false;
            filtroUsuariosActivo_OP_dgv_activo = 0;
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
                txt_Meta_3.Enabled = true;
                txt_Meta_5.Enabled = true;

                txt_Meta_4.Visible = true;
                txt_Meta_5.Visible = true;

                //nombrar controles
                txt_Meta_1.Hint = "No. Cajones por Hora";
                txt_Meta_2.Hint = "Kg Fresco por Hora";
                txt_Meta_3.Hint = "Relación Fresco-Seco";
                txt_Meta_4.Hint = "Kg Seco/Hr";
                txt_Meta_5.Hint = "Personal Idóneo";

                tap_control_metas.SelectedIndex = 0;
            }
            if (cmb_area.SelectedIndex == 1)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Kg por persona por Hora";
                txt_Meta_3.Hint = "Meta Kg por hr por 1 linea";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;

                //hacer invisibles controles
                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 1;
            }
            if (cmb_area.SelectedIndex == 2)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Kg por persona por Hora";
                txt_Meta_3.Hint = "Meta Kg/hr por 1 linea";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;

                //hacer invisibles controles
                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 2;
            }
            if (cmb_area.SelectedIndex == 3)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Capacidad por trompo";
                txt_Meta_3.Hint = "Cantidad de trompos";
                txt_Meta_4.Hint = "Meta Kg por hora";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_4.Visible = true;

                //hacer invisibles controles
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 3;
            }
            if (cmb_area.SelectedIndex == 4)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Capacidad por molino";
                txt_Meta_3.Hint = "Cantidad de molinos";
                txt_Meta_4.Hint = "Meta Kg por hora";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_4.Visible = true;

                //hacer invisibles controles
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 4;
            }
            if (cmb_area.SelectedIndex == 5)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Capacidad/(Trompo y/o Máq.)";
                txt_Meta_3.Hint = "Cantidad/(Trompo y/o Máq.)";
                txt_Meta_4.Hint = "Meta Kg por hora";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_4.Visible = true;

                //hacer invisibles controles
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 5;
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
            //habilitar controles
            cmb_area.Enabled = true;
            btn_meta_cancel.Enabled = true;
            btn_meta_save.Enabled = true;
            btn_meta_edit.Enabled = false;
            btn_meta_delete.Enabled = false;

            //limpiar campos
            cmb_area.SelectedIndex = -1;
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;

            //limpiar variables globales
            id_global_meta_deshidratado = string.Empty; 
            id_global_meta_empacado = string.Empty; 
            id_global_meta_inspec = string.Empty; 
            id_global_meta_evaporado = string.Empty; 
            id_global_meta_platinum = string.Empty;

            //enfocar controles
            cmb_area.Focus();
            txt_op.Focus();

        }

        private void btn_meta_save_Click(object sender, EventArgs e)
        {
            if(cmb_area.SelectedIndex == -1) 
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            /////Deshidratado
            if (cmb_area.SelectedIndex == 0 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_3.Text == string.Empty || txt_Meta_5.Text == string.Empty)) 
            { 
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 0)
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
                        new NpgsqlParameter("@no_box_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@kg_f_h", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@relacion_fs", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@kg_s_h", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_4.Text)
                        },
                        new NpgsqlParameter("@personal_i", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_5.Text)
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
                    new NpgsqlParameter("@No_box_h", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Kg_f_h", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Relacion_f_s", Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@kg_s_h", Convert.ToDecimal(txt_Meta_4.Text)),
                    new NpgsqlParameter("@Personal_i", Convert.ToDecimal(txt_Meta_5.Text))
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
                    id_global_meta_deshidratado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            /////Empacado
            ///
            if (cmb_area.SelectedIndex == 1 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 1)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_empacado))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_empacado);

                    queryInsertUpdate = "UPDATE public.\"Empacado\" SET \"OP\" = @OP, \"Personal_idoneo\" = @Personal_I, \"Kg_person_hr\" = @Kg_p_hr, \"Meta_kg_hr_line\" = @Meta_kg_hr_line WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Kg_p_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr_line", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_3.Text)
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
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Empacado\" WHERE \"OP\"  ILIKE @OP;";
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
                    queryInsertUpdate = "INSERT INTO public.\"Empacado\" (\"OP\", \"Personal_idoneo\", \"Kg_person_hr\", \"Meta_kg_hr_line\") VALUES (@OP, @Personal_I, @Kg_person_hr, @Meta_kg_hr_line);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper()),
                    new NpgsqlParameter("@Personal_I", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Kg_person_hr", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Meta_kg_hr_line", Convert.ToDecimal(txt_Meta_3.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_Empacado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_empacado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Inspeccion
            ///
            if (cmb_area.SelectedIndex == 2 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 2)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_inspec))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_inspec);

                    queryInsertUpdate = "UPDATE public.\"Inspeccion\" SET \"OP\" = @OP, \"Personal_idoneo\" = @Personal_I, \"Kg_person_hr\" = @Kg_p_hr, \"Meta_kg_hr_line\" = @Meta_kg_hr_line WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Kg_p_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr_line", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_3.Text)
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
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Inspeccion\" WHERE \"OP\"  ILIKE @OP;";
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
                    queryInsertUpdate = "INSERT INTO public.\"Inspeccion\" (\"OP\", \"Personal_idoneo\", \"Kg_person_hr\", \"Meta_kg_hr_line\") VALUES (@OP, @Personal_I, @Kg_person_hr, @Meta_kg_hr_line);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper()),
                    new NpgsqlParameter("@Personal_I", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Kg_person_hr", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Meta_kg_hr_line", Convert.ToDecimal(txt_Meta_3.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_inspec();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_inspec = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Evaporado
            ///
            if (cmb_area.SelectedIndex == 3 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 3)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_evaporado))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_evaporado);

                    queryInsertUpdate = "UPDATE public.\"Evaporado\" SET \"OP\" = @OP, \"Personal_idoneo\" = @Personal_I, \"Capacidad_trompos\" = @Capacidad_T, \"Cantidad_trompos\" = @Cantidad_T, \"Meta_kg_hr\" = @Meta_kg_hr WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Capacidad_T", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Cantidad_T", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_4.Text)
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
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Evaporado\" WHERE \"OP\"  ILIKE @OP;";
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
                    queryInsertUpdate = "INSERT INTO public.\"Evaporado\" (\"OP\", \"Personal_idoneo\", \"Capacidad_trompos\", \"Cantidad_trompos\", \"Meta_kg_hr\") VALUES (@OP, @Personal_I, @Capacidad_trompos, @Cantidad_trompos, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper()),
                    new NpgsqlParameter("@Personal_I", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_trompos", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_trompos", Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", Convert.ToDecimal(txt_Meta_4.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_evaporado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_evaporado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Grind
            ///
            if (cmb_area.SelectedIndex == 4 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 4)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_grind))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_grind);

                    queryInsertUpdate = "UPDATE public.\"Grind\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Capacidad_Molino\" = @Capacidad_m, \"Cantidad_molinos\" = @Cantidad_m, \"Meta_Kg_hr\" = @Meta_kg_hr WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Capacidad_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Cantidad_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_4.Text)
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
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Grind\" WHERE \"OP\"  ILIKE @OP;";
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
                    queryInsertUpdate = "INSERT INTO public.\"Grind\" (\"OP\", \"Personal_Idoneo\", \"Capacidad_Molino\", \"Cantidad_molinos\", \"Meta_Kg_hr\") VALUES (@OP, @Personal_I, @Capacidad_m, @Cantidad_m, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper()),
                    new NpgsqlParameter("@Personal_I", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_m", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_m", Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", Convert.ToDecimal(txt_Meta_4.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_grind();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_grind = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Platinum
            ///
            if (cmb_area.SelectedIndex == 5 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 5)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_platinum))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_platinum);

                    queryInsertUpdate = "UPDATE public.\"Platinum\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Cap_Trompo_machin\" = @Cap_Trompo_m, \"Canti_Trompo_machin\" = @Cantidad_Trompo_m, \"Meta_Kg_Hr\" = @Meta_kg_hr WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Cap_Trompo_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Cantidad_Trompo_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_4.Text)
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
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Platinum\" WHERE \"OP\"  ILIKE @OP;";
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
                    queryInsertUpdate = "INSERT INTO public.\"Platinum\" (\"OP\", \"Personal_Idoneo\", \"Cap_Trompo_machin\", \"Canti_Trompo_machin\", \"Meta_Kg_Hr\") VALUES (@OP, @Personal_I, @Capacidad_Trompo_m, @Cantidad_Trompo_m, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper()),
                    new NpgsqlParameter("@Personal_I", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_Trompo_m", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_Trompo_m", Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", Convert.ToDecimal(txt_Meta_4.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_platinum();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_platinum = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
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

                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
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
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void btn_meta_edit_Click(object sender, EventArgs e)
        {
            btn_meta_edit.Enabled = false;
            if (cmb_area.SelectedIndex == 0)
            {
                btn_meta_save.Enabled = true;
                btn_meta_cancel.Enabled = true;
                btn_meta_delete.Enabled = false;
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_5.Enabled = true;
                txt_op.Focus();
            }
            if (cmb_area.SelectedIndex == 1 || cmb_area.SelectedIndex == 2)
            {
                btn_meta_save.Enabled = true;
                btn_meta_cancel.Enabled = true;
                btn_meta_delete.Enabled = false;
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_op.Focus();
            }
            if (cmb_area.SelectedIndex == 3 || cmb_area.SelectedIndex == 4 || cmb_area.SelectedIndex == 5)
            {
                btn_meta_save.Enabled = true;
                btn_meta_cancel.Enabled = true;
                btn_meta_delete.Enabled = false;
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_op.Focus();
            }
        }

        private void btn_meta_cancel_Click(object sender, EventArgs e)
        {
            btn_meta_save.Enabled = false;
            btn_meta_cancel.Enabled = false;
            btn_meta_delete.Enabled = false;
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
            id_global_meta_empacado = string.Empty;
            id_global_meta_inspec = string.Empty;
        }

        private void btn_delete_user_Click(object sender, EventArgs e)
        {
            if (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este usuario?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id_user = Convert.ToInt32(id_global_users);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Usuarios""
                       WHERE ""ID_User"" = @idUser";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idUser", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_user
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_users(); // Actualizar el DataGridView de usuarios
                    limpiarCampos(); // Limpiar los campos de texto
                    btn_save.Enabled = false;
                    btn_cancel.Enabled = false;
                    btn_delete_user.Enabled = false;
                    txt_no_emp.Enabled = false;
                    txt_usuario.Enabled = false;
                    txt_contra.Enabled = false;
                    btn_edit.Enabled = false;
                    id_global_users = string.Empty;

                    cmb_nivel_user.Enabled = true;
                    cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
                    cmb_nivel_user.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el usuario", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }

        private void btn_meta_delete_Click(object sender, EventArgs e)
        {
            ///deshidratado
            ///
            if (cmb_area.SelectedIndex == 0 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Deshidratado?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_deshidratado);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Deshidratado""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_Deshitratado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    txt_Meta_4.Enabled = false;
                    txt_Meta_5.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_deshidratado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }

            ///empacado
            ///
            if (cmb_area.SelectedIndex == 1 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Empacado?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_empacado);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Empacado""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_Empacado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_empacado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///inspeccion
            ///
            if (cmb_area.SelectedIndex == 2 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Inspección?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_inspec);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Inspeccion""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_inspec();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_inspec = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Evaporado
            ///
            if (cmb_area.SelectedIndex == 3 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Evaporado?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_evaporado);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Evaporado""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_evaporado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_evaporado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Grind
            ///
            if (cmb_area.SelectedIndex == 4 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Grind?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_grind);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Grind""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_grind();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_grind = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Platinum
            ///
            if (cmb_area.SelectedIndex == 5 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Platinum?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_platinum);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Platinum""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_platinum();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_platinum = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void txt_Meta_1_TextChanged(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == 0)
            {
                if (decimal.TryParse(txt_Meta_1.Text, out decimal valor))
                {
                    // Fórmula: (valor * 380) - (valor * 380 * 10%)
                    decimal resultado = (valor * 380) - (valor * 380 * 0.10m);

                    txt_Meta_2.Text = resultado.ToString("0.00"); // 2 decimales
                }
                else
                {
                    txt_Meta_2.Text = string.Empty; // Limpia si no es válido
                }
            }
            
        }

        private void txt_Meta_3_TextChanged(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == 0)
            {
                decimal valor1;
                valor1 = decimal.TryParse(txt_Meta_2.Text, out valor1) ? valor1 : 0;
                if (decimal.TryParse(txt_Meta_3.Text, out decimal valor2))
                {
                    // Fórmula: KG FRESCO POR HORA/RELACION FRESCO-SECO
                    decimal resultado = valor1 / valor2;

                    txt_Meta_4.Text = resultado.ToString("0.00"); // 2 decimales
                }
                else
                {
                    txt_Meta_4.Text = string.Empty; // Limpia si no es válido
                }
            }
            if (cmb_area.SelectedIndex == 3 || cmb_area.SelectedIndex == 4 || cmb_area.SelectedIndex == 5)
            {
                decimal valor1;
                valor1 = decimal.TryParse(txt_Meta_2.Text, out valor1) ? valor1 : 0;
                if (decimal.TryParse(txt_Meta_3.Text, out decimal valor2))
                {
                    // Fórmula: KG FRESCO POR HORA/RELACION FRESCO-SECO
                    decimal resultado = valor1 * valor2;

                    txt_Meta_4.Text = resultado.ToString("0.00"); // 2 decimales
                }
                else
                {
                    txt_Meta_4.Text = string.Empty; // Limpia si no es válido
                }
            }
        }

        private void tap_control_metas_Selected(object sender, TabControlEventArgs e)
        {
            
        }

        private void txt_Meta_2_TextChanged(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == 1 || cmb_area.SelectedIndex == 2)
            {
                decimal valor1;
                valor1 = decimal.TryParse(txt_Meta_1.Text, out valor1) ? valor1 : 0;
                if (decimal.TryParse(txt_Meta_2.Text, out decimal valor2))
                {
                    // Fórmula: KG FRESCO POR HORA/RELACION FRESCO-SECO
                    decimal resultado = valor1 * valor2;

                    txt_Meta_3.Text = resultado.ToString("0.00"); // 2 decimales
                }
                else
                {
                    txt_Meta_3.Text = string.Empty; // Limpia si no es válido
                }
            }
        }

        private void dgv_metas_emp_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                id_global_meta_empacado = dgv_metas_emp.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_2.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[3].Value.ToString();

                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 1;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_insp_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                id_global_meta_inspec = dgv_metas_insp.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_2.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[3].Value.ToString();

                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 2;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_Eva_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                id_global_meta_evaporado = dgv_metas_Eva.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 3;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_Grind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                id_global_meta_grind = dgv_metas_Grind.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 4;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_platinum_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                id_global_meta_platinum = dgv_metas_platinum.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_platinum.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_platinum.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_platinum.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_platinum.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_platinum.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 5;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void cmb_area_search_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_area_search.SelectedIndex != -1)
            {
                txt_OP_Search.Enabled = true;
                if(cmb_area_search.SelectedIndex == 0)
                {
                    tap_control_metas.SelectedIndex = 0;
                }
                if (cmb_area_search.SelectedIndex == 1)
                {
                    tap_control_metas.SelectedIndex = 1;
                }
                if (cmb_area_search.SelectedIndex == 2)
                {
                    tap_control_metas.SelectedIndex = 2;
                }
                if (cmb_area_search.SelectedIndex == 3)
                {
                    tap_control_metas.SelectedIndex = 3;
                }
                if (cmb_area_search.SelectedIndex == 4)
                {
                    tap_control_metas.SelectedIndex = 4;
                }
                if (cmb_area_search.SelectedIndex == 5)
                {
                    tap_control_metas.SelectedIndex = 5;
                }
            } 
        }

        private void materialButton7_Click(object sender, EventArgs e)
        {
            //Deshidratado
            ///
            if (cmb_area_search.SelectedIndex == 0)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 1;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_des.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();
                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_des.ClearSelection();
                dgv_metas_des.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_des.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_des.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;

                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Empacado
            ///
            if (cmb_area_search.SelectedIndex == 1)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 2;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_emp.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_emp.ClearSelection();
                dgv_metas_emp.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_emp.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_emp.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;

                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Inspeccion
            ///
            if (cmb_area_search.SelectedIndex == 2)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 3;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_insp.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_insp.ClearSelection();
                dgv_metas_insp.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_insp.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_insp.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;
                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Evaporado
            ///
            if (cmb_area_search.SelectedIndex == 3)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 4;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_Eva.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_Eva.ClearSelection();
                dgv_metas_Eva.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_Eva.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_Eva.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;
                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Grind
            ///
            if (cmb_area_search.SelectedIndex == 4)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 5;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_Grind.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno                          
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_Grind.ClearSelection();
                dgv_metas_Grind.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_Grind.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_Grind.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;
                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Platinum
            ///
            if (cmb_area_search.SelectedIndex == 5)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 6;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_platinum.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_platinum.ClearSelection();
                dgv_metas_platinum.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_platinum.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_platinum.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;
                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void materialButton5_Click(object sender, EventArgs e)
        {
            limpiar_filtros_OP();
        }

        //private void calcular_turno()
        //{
        //    // Leer las horas ingresadas
        //    string horaInicioText = Mask_txt_hr1.Text;
        //    string horaFinText = Mask_txt_hr2.Text;

        //    // Convertir a DateTime
        //    DateTime horaInicio = DateTime.ParseExact(horaInicioText, "HH:mm", null);
        //    DateTime horaFin = DateTime.ParseExact(horaFinText, "HH:mm", null);

        //    // Si la hora de fin es menor a la de inicio, significa que pasó a otro día
        //    if (horaFin < horaInicio)
        //    {
        //        horaFin = horaFin.AddDays(1);
        //    }

        //    // Calcular diferencia
        //    TimeSpan diferencia = horaFin - horaInicio;

        //    // Mostrar horas totales (con decimales si hay minutos)
        //    Txt_Read_1.Text = diferencia.TotalHours.ToString("0.##");
        //}

        private void calcular_horas_programadas()
        {
            double hr_programada, meta_programada, meta_x_hr;

            hr_programada = double.TryParse(Txt_Read_1.Text, out hr_programada) ? hr_programada : 0;
            meta_x_hr = double.TryParse(Txt_meta.Text, out meta_x_hr) ? meta_x_hr : 0;
            meta_programada = hr_programada * meta_x_hr;
            Txt_Read_2.Text = meta_programada.ToString("0.##");
        }

        private void calcular_turno()
        {
            try
            {

                if (!string.IsNullOrWhiteSpace(Mask_txt_hr1.Text) && !string.IsNullOrWhiteSpace(Mask_txt_hr2.Text) && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :")
                {
                    // Si la hora de fin es menor a la de inicio, significa que pasó a otro día
                    if (horaFin < horaInicio)
                    {
                        horaFin = horaFin.AddDays(1);
                    }

                    // Calcular diferencia inicial
                    TimeSpan diferencia = horaFin - horaInicio;

                    // Restar minutos energia
                    if (!string.IsNullOrEmpty(txt_Tiempo_energia.Text) && int.TryParse(txt_Tiempo_energia.Text, out int minutosEnergia))
                    {
                        diferencia = diferencia.Subtract(TimeSpan.FromMinutes(minutosEnergia));
                    }

                    // Restar minutos comida
                    if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && int.TryParse(txt_Tiempo_comida.Text, out int minutosComida1))
                    {
                        diferencia = diferencia.Subtract(TimeSpan.FromMinutes(minutosComida1));
                    }

                    // Mostrar horas totales (con decimales si hay minutos) - SIN CAMBIOS
                    Txt_Read_1.Text = diferencia.TotalHours.ToString("0.##");



                    // Calcular la nueva diferencia restando los minutos
                    TimeSpan diferenciaConDescuento = diferencia;

                    // Restar minutos mecánicos
                    if (!string.IsNullOrEmpty(txt_TM_mecanico.Text) && int.TryParse(txt_TM_mecanico.Text, out int minutosMecanico))
                    {
                        diferenciaConDescuento = diferenciaConDescuento.Subtract(TimeSpan.FromMinutes(minutosMecanico));
                    }

                    // Restar minutos operativos
                    if (!string.IsNullOrEmpty(txt_TM_operativo.Text) && int.TryParse(txt_TM_operativo.Text, out int minutosOperativo))
                    {
                        diferenciaConDescuento = diferenciaConDescuento.Subtract(TimeSpan.FromMinutes(minutosOperativo));
                    }
                    // Asegurar que no sea negativo
                    if (diferenciaConDescuento.TotalMinutes < 0)
                    {
                        diferenciaConDescuento = TimeSpan.Zero;
                    }
                    if (!string.IsNullOrEmpty(Txt_meta.Text))
                    {
                        calcular_horas_programadas();
                    }
                    // Mostrar el resultado con descuento en Txt_Read_3
                    Txt_Read_3.Text = diferenciaConDescuento.TotalHours.ToString("0.##");
                }
                btn_save_ficha.Enabled = true;
            }
            catch (Exception ex)
            {
                 //MessageBox.Show($"Error en cálculo: {ex.Message}");
                if(ex is FormatException)
                {
                    //MetroFramework.MetroMessageBox.Show(this, "Formato de hora inválido. Asegúrese de usar HH:mm.",
                    //                                    "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btn_save_ficha.Enabled = false;
                }
            }
        }
        private void cb_Turno_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cb_Turno.SelectedIndex != -1)
            {
                if (cb_Turno.SelectedIndex == 0)
                {
                    Mask_txt_hr1.Text = "07:00";
                    Mask_txt_hr2.Text = "15:00";

                    calcular_turno();
                }
                if (cb_Turno.SelectedIndex == 1)
                {
                    Mask_txt_hr1.Text = "15:00";
                    Mask_txt_hr2.Text = "22:30";

                    calcular_turno();
                }
                if (cb_Turno.SelectedIndex == 2)
                {
                    Mask_txt_hr1.Text = "22:30";
                    Mask_txt_hr2.Text = "07:00";

                    calcular_turno();
                }
            }
        }

        private void btn_save_ficha_Click(object sender, EventArgs e)
        {
            //if (cb_OP.SelectedValue != null)
            //{
            //    // Obtener el ValueMember (ID_OP)
            //    int idOP = (int)cb_OP.SelectedValue;

            //    // Obtener el DisplayMember (OP)
            //    string op = cb_OP.Text;

            //    MessageBox.Show($"ID_OP: {idOP}\nOP: {op}");
            //}

            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            try
            {
                // Obtener ID del usuario actual (debes implementar esto)
                int idUsuarioActual = id_user;

                // Obtener datos de los TextBox
                DateTime fecha = dtp1.Value; // Tu MetroDateTime
                int turno = Convert.ToInt32(cb_Turno.Text);
                string lote = Txt_1.Text;
                string op = cb_OP.Text;
                decimal kgEnterProceso = Convert.ToDecimal(Txt_2.Text);
                decimal kgFrescosEnterSe = Convert.ToDecimal(Txt_Read_4.Text);
                decimal mermaCanica = Convert.ToDecimal(Txt_3.Text);
                decimal mermaPodrido = Convert.ToDecimal(Txt_4.Text);
                decimal mermaTina = Convert.ToDecimal(Txt_5.Text);
                decimal mermaPiso = Convert.ToDecimal(Txt_6.Text);
                decimal mermaCanaletas = Convert.ToDecimal(Txt_7.Text);
                decimal mermaLavadoBandas = Convert.ToDecimal(Txt_8.Text);
                decimal cascaraCarrete = Convert.ToDecimal(Txt_10.Text);
                int personal_Op = Convert.ToInt32(Txt_9.Text);
                decimal hr_pro = Convert.ToDecimal(Txt_Read_1.Text);
                decimal hr_efec = Convert.ToDecimal(Txt_Read_3.Text);
                decimal meta_prog = Convert.ToDecimal(Txt_Read_2.Text);
                string area = cb_Area.Text;
                decimal meta = Convert.ToDecimal(Txt_meta.Text);

                // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);

                // Insertar en tabla Ficha y obtener el ID_Ficha generado
                int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, lote, op,
                    kgEnterProceso, kgFrescosEnterSe, mermaCanica, mermaPodrido, mermaTina, mermaPiso,
                    mermaCanaletas, mermaLavadoBandas, cascaraCarrete, hrInicio, hrFin, personal_Op, hr_pro, hr_efec, meta_prog, area, meta);

                if (idFicha > 0)
                {
                    // Insertar en tablas relacionadas
                    InsertarTiemposMuertos(dbHelper, idFicha);

                    MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                        "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cb_Area.SelectedIndex = -1;
                    reiniciarCampos();
                    cb_Area.Focus();
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Error al guardar datos",
                                                        "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al guardar: {ex.Message}",
                                                        "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int InsertarFichaYRetornarID(DatabaseHelper dbHelper, int idUsuario, DateTime fecha,
     int turno, string lote, string op, decimal kgEnterProceso, decimal kgFrescosEnterSe,
     decimal mermaCanica, decimal mermaPodrido, decimal mermaTina, decimal mermaPiso,
     decimal mermaCanaletas, decimal mermaLavadoBandas, decimal cascaraCarrete,
     TimeSpan hrInicio, TimeSpan hrFin, int personal_O, decimal hr_programadas, decimal hr_efectivas, 
     decimal meta_prog, string area_f, decimal metaHr)
        {
            string query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""Lote"", ""OP"",
                    ""Kg_enter_proceso"", ""kg_frescos_enter_se"", ""Merma_canica"",
                    ""Merma_podrido"", ""Merma_tina"", ""Merma_piso"", ""Merma_canaletas"",
                    ""Merma_lavado_bandas"", ""Cascara_carrete"", ""Hr_inicio"", ""Hr_fin"", 
                    ""Hr_programadas"", ""Personal_Operativo"", ""Hr_efectivas"", ""MetaProg"", ""Area"", ""MetaHr""
                ) VALUES (
                    @id_user, @fecha, @turno, @lote, @op,
                    @kg_enter_proceso, @kg_frescos_enter_se, @merma_canica,
                    @merma_podrido, @merma_tina, @merma_piso, @merma_canaletas,
                    @merma_lavado_bandas, @cascara_carrete, @hr_inicio, @hr_fin, @hr_prog, @Personal_Op, @Hr_efec, @meta_prog, @Area, @MetaHr
                ) RETURNING ""ID_Ficha"";";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                new NpgsqlParameter("@lote", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = lote ?? (object)DBNull.Value },
                new NpgsqlParameter("@op", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = op ?? (object)DBNull.Value },
                new NpgsqlParameter("@kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = kgEnterProceso },
                new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = kgFrescosEnterSe },
                new NpgsqlParameter("@merma_canica", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = mermaCanica },
                new NpgsqlParameter("@merma_podrido", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = mermaPodrido },
                new NpgsqlParameter("@merma_tina", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = mermaTina },
                new NpgsqlParameter("@merma_piso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = mermaPiso },
                new NpgsqlParameter("@merma_canaletas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = mermaCanaletas },
                new NpgsqlParameter("@merma_lavado_bandas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = mermaLavadoBandas },
                new NpgsqlParameter("@cascara_carrete", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = cascaraCarrete },
                new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                new NpgsqlParameter("@hr_prog", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                new NpgsqlParameter("@Hr_efec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                new NpgsqlParameter("@Personal_Op", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                new NpgsqlParameter("@meta_prog", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_prog },
                new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f ?? (object)DBNull.Value },
                new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr }
            };

            return dbHelper.ExecuteScalarInt(query, parameters);
        }

        private void InsertarTiemposMuertos(DatabaseHelper dbHelper, int idFicha)
        {
            // Insertar tiempos muertos mecánicos desde DataGridView
            InsertarTiemposMuertosMecanicos(dbHelper, idFicha);

            // Insertar tiempos muertos operativos desde DataGridView
            InsertarTiemposMuertosOperativos(dbHelper, idFicha);

            // Insertar tiempo muerto comida desde TextBox
            InsertarTiempoMuertoComida(dbHelper, idFicha);

            // Insertar tiempo muerto energía desde TextBox
            InsertarTiempoMuertoEnergia(dbHelper, idFicha);
        }

        // Insertar en Tiempo_muerto_Mecanico desde DataGridView
        private void InsertarTiemposMuertosMecanicos(DatabaseHelper dbHelper, int idFicha)
        {
            foreach (DataGridViewRow row in dgv_mecanico.Rows)
            {
                if (!row.IsNewRow && row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    decimal minDetenidos = Convert.ToDecimal(row.Cells[0].Value);
                    string motivos = row.Cells[1].Value.ToString();

                    string query = @"INSERT INTO public.""Tiempo_muerto_Mecanico"" (
                            ""ID_Ficha"", ""Min_Detenidos"", ""Motivos""
                        ) VALUES (@id_ficha, @min_detenidos, @motivos);";

                    NpgsqlParameter[] parameters = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
                        new NpgsqlParameter("@min_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minDetenidos },
                        new NpgsqlParameter("@motivos", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = motivos ?? (object)DBNull.Value }
                    };

                    dbHelper.ExecuteNonQuery(query, parameters);
                }
            }
        }

        // Insertar en Tiempo_Muerto_Operativo desde DataGridView
        private void InsertarTiemposMuertosOperativos(DatabaseHelper dbHelper, int idFicha)
        {
            foreach (DataGridViewRow row in dgv_operativo.Rows)
            {
                if (!row.IsNewRow && row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    decimal minDetenidos = Convert.ToDecimal(row.Cells[0].Value);
                    string motivos = row.Cells[1].Value.ToString();

                    string query = @"INSERT INTO public.""Tiempo_Muerto_Operativo"" (
                            ""ID_Ficha"", ""Min_Detenidos"", ""Motivos""
                        ) VALUES (@id_ficha, @min_detenidos, @motivos);";

                    NpgsqlParameter[] parameters = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
                        new NpgsqlParameter("@min_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minDetenidos },
                        new NpgsqlParameter("@motivos", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = motivos ?? (object)DBNull.Value }
                    };

                    dbHelper.ExecuteNonQuery(query, parameters);
                }
            }
        }

        // Insertar en Tiempo_Muerto_Comida desde TextBox
        private void InsertarTiempoMuertoComida(DatabaseHelper dbHelper, int idFicha)
        {
            if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text))
            {
                decimal minutosDetenidos = Convert.ToDecimal(txt_Tiempo_comida.Text);

                string query = @"INSERT INTO public.""Tiempo_Muerto_Comida"" (
                        ""ID_Ficha"", ""Minutos_Detenidos""
                    ) VALUES (@id_ficha, @minutos_detenidos);";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
                    new NpgsqlParameter("@minutos_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minutosDetenidos }
                };

                dbHelper.ExecuteNonQuery(query, parameters);
            }
        }

        // Insertar en Tiempo_Muerto_Energia desde TextBox
        private void InsertarTiempoMuertoEnergia(DatabaseHelper dbHelper, int idFicha)
        {
            if (!string.IsNullOrEmpty(txt_Tiempo_energia.Text))
            {
                decimal minutosDetenidos = Convert.ToDecimal(txt_Tiempo_energia.Text);

                string query = @"INSERT INTO public.""Tiempo_Muerto_Energia"" (
                        ""ID_Ficha"", ""Minutos_Detenidos""
                    ) VALUES (@id_ficha, @minutos_detenidos);";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
                    new NpgsqlParameter("@minutos_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minutosDetenidos }
                };

                dbHelper.ExecuteNonQuery(query, parameters);
            }
        }

        private void CalcularTotalMecanico()
        {
            decimal total = 0;

            foreach (DataGridViewRow row in dgv_mecanico.Rows)
            {
                if (row.IsNewRow) continue; // ignora la fila en blanco para agregar

                if (row.Cells[0].Value != null &&
                    decimal.TryParse(row.Cells[0].Value.ToString(), out decimal valor))
                {
                    total += valor;
                }
            }
            txt_TM_mecanico.Text = total.ToString("0"); // formato con 2 decimales
        }

        private void CalcularTotaloperativo()
        {
            decimal total = 0;

            foreach (DataGridViewRow row in dgv_operativo.Rows)
            {
                if (row.IsNewRow) continue; // ignora la fila en blanco para agregar

                if (row.Cells[0].Value != null &&
                    decimal.TryParse(row.Cells[0].Value.ToString(), out decimal valor))
                {
                    total += valor;
                }
            }
            txt_TM_operativo.Text = total.ToString("0"); // formato con 2 decimales
        }

        private void dgv_mecanico_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            CalcularTotalMecanico();
        }

        private void dgv_mecanico_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            CalcularTotalMecanico();
        }

        private void dgv_mecanico_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            CalcularTotalMecanico();
        }

        private void dgv_operativo_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            CalcularTotaloperativo();
        }

        private void dgv_operativo_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            CalcularTotaloperativo();
        }

        private void dgv_operativo_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            CalcularTotaloperativo();
        }

        private void Txt_3_TextChanged(object sender, EventArgs e)
        {
            CalcularSuma();
        }

        private void Txt_4_TextChanged(object sender, EventArgs e)
        {
            CalcularSuma();
        }

        private void Txt_5_TextChanged(object sender, EventArgs e)
        {
            CalcularSuma();
        }

        private void Txt_6_TextChanged(object sender, EventArgs e)
        {
            CalcularSuma();
        }

        private void Txt_7_TextChanged(object sender, EventArgs e)
        {
            CalcularSuma();
        }

        private void Txt_8_TextChanged(object sender, EventArgs e)
        {
            CalcularSuma();
        }

        private void cb_OP_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void txt_TM_mecanico_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
        }

        private void txt_TM_operativo_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
        }

        private void Txt_2_TextChanged(object sender, EventArgs e)
        {
            CalcularSuma();
        }

        private void cb_OP_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //MessageBox.Show("entró");
            string valorBuscado = cb_OP.Text;

            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            if (cb_Area.SelectedIndex == 0 && cb_OP.SelectedIndex != -1)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"No_box_hr\" FROM public.\"Deshidratado\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;
             
                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["No_box_hr"].ToString();
                    Txt_meta.Text = resultado;

                    if (!string.IsNullOrEmpty(Txt_Read_1.Text))
                    {
                        calcular_horas_programadas();
                    }
                }
            }
        }

        private void txt_Tiempo_comida_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
        }

        private void Mask_txt_hr1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                horaInicioText = Mask_txt_hr1.Text;
                // Convertir a DateTime
                horaInicio = DateTime.ParseExact(horaInicioText, "HH:mm", null);
                errorProvider1.SetError(Mask_txt_hr1, ""); // Borra el error si es válido
                btn_save_ficha.Enabled = true;
                calcular_turno();
            }
            catch (Exception ex)
            {
                errorProvider1.SetError(Mask_txt_hr1, "Formato de hora inválido. Asegúrese de usar HH:mm.");
                btn_save_ficha.Enabled = false;
            }
        }

        private void Mask_txt_hr2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                horaFinText = Mask_txt_hr2.Text;
                // Convertir a DateTime
                horaFin = DateTime.ParseExact(horaFinText, "HH:mm", null);
                errorProvider2.SetError(Mask_txt_hr2, ""); // Borra el error si es válido
                btn_save_ficha.Enabled = true;
                calcular_turno();
            }
            catch (Exception ex)
            {
                errorProvider2.SetError(Mask_txt_hr2, "Formato de hora inválido. Asegúrese de usar HH:mm.");
                btn_save_ficha.Enabled = false;
            }
        }

        private void txt_Tiempo_energia_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
        }
    }
}
