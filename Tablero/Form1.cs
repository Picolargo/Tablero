using MaterialSkin;
using MaterialSkin.Controls;
using Npgsql;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.Export;
using Telerik.WinControls.UI;
using Telerik.WinControls.UI.Export;
using Excel = Microsoft.Office.Interop.Excel;

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
        private string id_global_meta_revolturas = string.Empty;
        private string id_global_meta_polvos = string.Empty;
        private string id_global_meta_maquinas = string.Empty;
        private string id_global_ficha = string.Empty;
        private string id_global_polvos_calidad = string.Empty;
        private string id_global_tunel_calidad = string.Empty;
        private string id_global_detalles_OP = string.Empty;
        private int id_user = 0;
        private string nivel_user = string.Empty;
        private bool filtroUsuariosActivo = false;
        private bool filtroUsuariosActivo_OP = false;
        private bool filtrodetallesOP = false;
        private int filtroUsuariosActivo_OP_dgv_activo = 0;
        private string horaInicioText = string.Empty;
        private string horaFinText = string.Empty;
        private DateTime horaInicio;
        private DateTime horaFin;
        private bool editar = false;
        //variable para la conexión a la base de datos
        string connectionString = string.Empty;
        //
        public Form_principal(string var_no_empledo, string var_nom_empledo, int ID_usuario, string nivel, string conexionstring)
        {
            InitializeComponent();
            // Configurar propiedades similares al ejemplo
            rgv_reporte_consolidado.EnableGrouping = false;
            rgv_reporte_consolidado.EnableHotTracking = true;
            rgv_reporte_consolidado.ShowFilteringRow = false;
            rgv_reporte_consolidado.EnableFiltering = true;
            rgv_reporte_consolidado.EnableCustomFiltering = true;

            this.WindowState = FormWindowState.Maximized;

            lbl_no_emp2.Text = var_no_empledo; // Mostrar el número de empleado en el label correspondiente
            lbl_nom2.Text = var_nom_empledo.ToUpper(); // Mostrar el nombre del empleado en el label correspondiente
            connectionString = conexionstring; // Asignar la cadena de conexión pasada como parámetro
            id_user = ID_usuario; // Asignar el ID del usuario pasado como parámetro
            nivel_user = nivel; // Asignar el nivel del usuario pasado como parámetro

            // Initialize MaterialSkinManager and set the theme and color scheme  
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Orange600, Primary.Orange600, Primary.BlueGrey800, Accent.Blue700, TextShade.WHITE);


            // Personaliza todos los DataGridView de una sola línea cada uno
            PersonalizarDataGridView(dgv_mecanico, true);
            PersonalizarDataGridView(dgv_operativo, true);
            PersonalizarDataGridView(dgv_users);
            PersonalizarDataGridView(dgv_metas_des);
            PersonalizarDataGridView(dgv_metas_emp);
            PersonalizarDataGridView(dgv_metas_insp);
            PersonalizarDataGridView(dgv_metas_Eva);
            PersonalizarDataGridView(dgv_metas_Grind);
            PersonalizarDataGridView(dgv_metas_revolturas);
            PersonalizarDataGridView(dgv_metas_polvos);
            PersonalizarDataGridView(dgv_metas_maquinas);
            PersonalizarDataGridView(dgv_polvos_calidad);
            PersonalizarDataGridView(dgv_detalles_op);
            PersonalizarDataGridView(dgv_Tunel_calidad);
            PersonalizarDataGridView(dgv_reporte_merma);
        }

        private void PersonalizarDataGridView(DataGridView dgv, bool editable = false)
        {
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold);

            dgv.BackgroundColor = Color.White;
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.Black;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Regular);

            dgv.GridColor = Color.FromArgb(255, 152, 0);
            dgv.RowHeadersVisible = false;
            dgv.BorderStyle = BorderStyle.None;

            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = false;

            // Opcional según si quieres permitir edición
            dgv.AllowUserToAddRows = editable;
            dgv.AllowUserToDeleteRows = editable;
            dgv.AllowUserToResizeRows = false;
            dgv.ReadOnly = !editable;
        }

        private void dgv_mecanico_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_mecanico.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
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
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
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
            if (nivel_user == "Administrador")
            {
                lbl_user_no_emp.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
                lbl_Nom.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
                dtp_polvos.Value = DateTime.Now;
                dtp_tunel.Value = DateTime.Now;

                // Obtener y mostrar el número de semana inicial
                ActualizarNumeroSemana();

                actualiza_grid_users(); // Llamar al método para actualizar el DataGridView de usuarios
                actualiza_grid_Deshitratado();
                actualiza_grid_Empacado();
                actualiza_grid_inspec();
                actualiza_grid_evaporado();
                actualiza_grid_grind();
                actualiza_revolturas();
                actualiza_maquinas();
                actualiza_polvos();
                actualiza_polvos_calidad();
                actualiza_detalles_OP();
                actualiza_tunel_calidad();
                configurar_limpieza();
            }
            if (nivel_user == "Supervisor")
            {
                materialTabControl1.TabPages.Remove(tabPage2);
                materialTabControl1.TabPages.Remove(tabPage3);
                materialTabControl1.TabPages.Remove(tabPage4);
                materialTabControl1.TabPages.Remove(tabPage10);
                materialTabControl1.TabPages.Remove(tabPage11);
                tabControl_detallesOP.TabPages.Remove(tabPage12);
                actualiza_detalles_OP();

                lbl_user_no_emp.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
                lbl_Nom.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
            }
            if (nivel_user == "Calidad")
            {
                materialTabControl1.TabPages.Remove(tabPage1);
                materialTabControl1.TabPages.Remove(tabPage2);
                materialTabControl1.TabPages.Remove(tabPage3);
                materialTabControl1.TabPages.Remove(tabPage4);
                materialTabControl1.TabPages.Remove(tabPage9);
                materialTabControl1.TabPages.Remove(tabPage10);

                dtp_polvos.Value = DateTime.Now;
                dtp_tunel.Value = DateTime.Now;
                // Obtener y mostrar el número de semana inicial
                ActualizarNumeroSemana();
                actualiza_polvos_calidad();
                configurar_limpieza();
            }
        }

        private void configurar_limpieza()
        {
            tabPage14.BackColor = Color.White;
            tabPage15.BackColor = Color.White;
            lbl_no_semana_polvos.BackColor = Color.White;
            lbl_no_semana_tunel.BackColor = Color.White;
            txt_no_semana_polvos.BackColor = Color.White;
            txt_no_semana_tunel.BackColor = Color.White;
            tabControl1.SelectedIndex = 0;
            tab_limpiezas.SelectedIndex = 0;
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

        private void actualiza_revolturas()
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
                          FROM public.""Revolturas""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_revolturas, null);

            // Configurar el DataGridView
            dgv_metas_revolturas.Columns[0].Visible = false;
            dgv_metas_revolturas.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_maquinas()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Cap_Trompo_machin"" as ""Capacidad por Trompo y/o Máquinas"", 
                            ""Canti_Trompo_machin"" as ""Cantidad de Trompos y/o Máquinas"",
                            ""Meta_Kg_Hr"" as ""Meta Kg por hora"", 
                            ""Personal_Idoneo"" as ""Personal Idóneo"",
                            ""Kg_pz"" as ""Kg por pieza""
                          FROM public.""Maquinas""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_maquinas, null);

            // Configurar el DataGridView
            dgv_metas_maquinas.Columns[0].Visible = false;
            dgv_metas_maquinas.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_polvos()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Personal_Idoneo"" as ""Personal Idóneo"", 
                            ""Meta_kg_hr_hum"" as ""Meta kg/hr Temporada Humeda"",
                            ""Meta_kg_hr_idon"" as ""Meta kg/hr Temporada Idónea"" 
                          FROM public.""Polvos""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_polvos, null);

            // Configurar el DataGridView
            dgv_metas_polvos.Columns[0].Visible = false;
            dgv_metas_polvos.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_detalles_OP()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_Dt_OP"" as ""ID"", 
                            ""Orden_Produccion"" as ""Orden de Producción"", 
                            ""Producto"", 
                            ""Medida"",
                            ""Descripcion"" as ""Descripción"", 
                            ""Especificacion"" as ""Especificación"",
                            ""Ingredientes"",
                            ""Humedad"",
                            ""Comercio"",
                            ""Manzana"",
                            ""Analisis"",
                            ""Area_Proceso"" as ""Área de Proceso"",
                            ""OP_Origen"" as ""OP de Origen"",
                            ""Destino"",
                            ""Clasificacion"" as ""Clasificación""
                          FROM public.""Detalles_OP""
                          ORDER BY ""Orden_Produccion"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_detalles_op, null);

            // Configurar el DataGridView
            dgv_detalles_op.Columns[0].Visible = false;
            dgv_detalles_op.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_polvos_calidad()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta para la tabla Ficha filtrando por Área = 'Polvos'
            string querySimple = @"SELECT 
                    ""ID_Limpieza"" as ""ID"", 
                    ""Fecha"", 
                    EXTRACT(WEEK FROM ""Fecha"") as ""No Semana"",
                    ""Kg_merma"" as ""Kg de merma por Limpieza""
                  FROM public.""Limpieza_polvos""
                  ORDER BY ""Fecha"" DESC;";

            // Cargar los datos de la tabla Ficha en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_polvos_calidad, null);

            // Configurar el DataGridView
            dgv_polvos_calidad.Columns[0].Visible = false;

            // Configurar formato para la columna de fecha
            if (dgv_polvos_calidad.Columns["Fecha"] != null)
            {
                dgv_polvos_calidad.Columns["Fecha"].DefaultCellStyle.Format = "dd/MM/yyyy";
            }
        }
        private void actualiza_tunel_calidad()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta para la tabla Ficha filtrando por Área = 'Polvos'
            string querySimple = @"SELECT 
                    ""ID_Tunel"" as ""ID"", 
                    ""Fecha"", 
                    EXTRACT(WEEK FROM ""Fecha"") as ""No Semana"",
                    ""Kg_merma"" as ""Kg de merma de Tunel""
                  FROM public.""Limpieza_tunel""
                  ORDER BY ""Fecha"" DESC;";

            // Cargar los datos de la tabla Ficha en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_Tunel_calidad, null);

            // Configurar el DataGridView
            dgv_Tunel_calidad.Columns[0].Visible = false;

            // Configurar formato para la columna de fecha
            if (dgv_Tunel_calidad.Columns["Fecha"] != null)
            {
                dgv_Tunel_calidad.Columns["Fecha"].DefaultCellStyle.Format = "dd/MM/yyyy";
            }
        }
        private void actualiza_reporte_merma()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta para la tabla Ficha filtrando por Área = 'Polvos'
            string querySimple = @"WITH semanas AS (SELECT ""Area"", EXTRACT(WEEK FROM ""Fecha"") AS numero_semana, EXTRACT(YEAR FROM ""Fecha"") AS año,
                                    -- Cálculo para Tunel/Sumergidor
                                    CASE 
                                        WHEN ""Area"" = 'Tunel/Sumergidor' THEN
                                            SUM(""Merma_podrido"" + ""Merma_tina"" + ""Merma_piso"" + ""Merma_canaletas"" + ""Merma_lavado_bandas"") / 
                                            NULLIF(SUM(""Kg_enter_proceso"" - ""Merma_canica""), 0)
                                    END AS merma_tunel,
        
                                    -- Cálculo para Despegue (se unirá con Limpieza_tunel)
                                    CASE 
                                        WHEN ""Area"" = 'Despegue' THEN
                                            SUM(""Merma_kg"") / NULLIF(SUM(""Kg_prod_seco""), 0)
                                    END AS merma_despegue_base,
        
                                    -- Cálculo para otras áreas
                                    CASE 
                                        WHEN ""Area"" IN ('Evaporado', 'Grind', 'Inspección', 'Empacado', 'Revolturas', 'Máquinas') THEN
                                            SUM(""Merma_kg"") / NULLIF(SUM(""Kg_enter_proceso""), 0)
                                    END AS merma_otras_areas,
        
                                    -- Cálculo para Polvos (se unirá con Limpieza_polvos)
                                    CASE 
                                        WHEN ""Area"" = 'Polvos' THEN
                                            SUM(""Merma_kg"") / NULLIF(SUM(""Kg_prod_seco""), 0)
                                    END AS merma_polvos_base,
        
                                    -- Sumas para Despegue
                                    SUM(CASE WHEN ""Area"" = 'Despegue' THEN ""Merma_kg"" ELSE 0 END) AS suma_merma_despegue,
                                    SUM(CASE WHEN ""Area"" = 'Despegue' THEN ""Kg_prod_seco"" ELSE 0 END) AS suma_kg_seco_despegue,
        
                                    -- Sumas para Polvos
                                    SUM(CASE WHEN ""Area"" = 'Polvos' THEN ""Merma_kg"" ELSE 0 END) AS suma_merma_polvos,
                                    SUM(CASE WHEN ""Area"" = 'Polvos' THEN ""Kg_prod_seco"" ELSE 0 END) AS suma_kg_seco_polvos
        
                                FROM public.""Ficha""
                                GROUP BY ""Area"", EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
                            ),

                            limpieza_tunel_semanal AS (
                                SELECT 
                                    EXTRACT(WEEK FROM ""Fecha"") AS numero_semana,
                                    EXTRACT(YEAR FROM ""Fecha"") AS año,
                                    SUM(""Kg_merma"") AS kg_merma_tunel_semana
                                FROM public.""Limpieza_tunel""
                                GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
                            ),

                            limpieza_polvos_semanal AS (
                                SELECT 
                                    EXTRACT(WEEK FROM ""Fecha"") AS numero_semana,
                                    EXTRACT(YEAR FROM ""Fecha"") AS año,
                                    SUM(""Kg_merma"") AS kg_merma_polvos_semana
                                FROM public.""Limpieza_polvos""
                                GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
                            )

                            SELECT 
                                s.""Area"",
                                s.numero_semana as ""Numero de Semana"",
                                s.año as ""Año"",
                                CASE 
                                    WHEN s.""Area"" = 'Tunel/Sumergidor' THEN
                                        ROUND(COALESCE(s.merma_tunel, 0) * 100, 2)
            
                                    WHEN s.""Area"" = 'Despegue' THEN
                                        ROUND(
                                            COALESCE(
                                                (s.suma_merma_despegue + COALESCE(lt.kg_merma_tunel_semana, 0)) / 
                                                NULLIF(s.suma_kg_seco_despegue, 0), 0) * 100, 2)
            
                                    WHEN s.""Area"" IN ('Evaporado', 'Grind', 'Inspección', 'Empacado', 'Revolturas', 'Máquinas') THEN
                                        ROUND(COALESCE(s.merma_otras_areas, 0) * 100, 2)
            
                                    WHEN s.""Area"" = 'Polvos' THEN
                                        ROUND(COALESCE((s.suma_merma_polvos + COALESCE(lp.kg_merma_polvos_semana, 0)) / NULLIF(s.suma_kg_seco_polvos, 0), 0) * 100, 2)
                                    ELSE 0
                                END AS ""Porcentaje de Merma""
    
                            FROM semanas s
                            LEFT JOIN limpieza_tunel_semanal lt ON s.numero_semana = lt.numero_semana AND s.año = lt.año
                            LEFT JOIN limpieza_polvos_semanal lp ON s.numero_semana = lp.numero_semana AND s.año = lp.año
                            WHERE s.""Area"" IS NOT NULL
                            ORDER BY s.año, s.numero_semana, s.""Area"";";

            // Cargar los datos de la tabla Ficha en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_reporte_merma, null);
        }
        private void cb_Area_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                //TUNEL/ SUMERGIDOR
                if (!editar) { reiniciarCampos(); }

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
                Txt_3.EmbeddedLabelText = "Canica";
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

                //hacer visibles controles
                Txt_1.Visible = true;
                Txt_6.Visible = true;
                Txt_7.Visible = true;
                Txt_8.Visible = true;
                Txt_9.Visible = true;
                Txt_10.Visible = true;

                //hacer invisibles controles

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                Txt_11.Visible = false;
                cb_lote.Visible = false;
                cb_proceso.Visible = false;

                //Inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Txt_6.Enabled = false;
                Txt_7.Enabled = false;
                Txt_8.Enabled = false;
                Txt_9.Enabled = false;
                Txt_10.Enabled = false;
                Txt_11.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                Txt_Read_1.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Deshidratado\" ORDER BY \"OP\";";
                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 1)
            {
                //Despegue
                if (!editar) { reiniciarCampos(); }

                //hacer visible para Despegue
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;

                Txt_1.Visible = false;

                //habilitar controles
                cb_Turno.Enabled = false;
                cb_OP.Enabled = false;
                dtp1.Enabled = true;

                cb_lote.Visible = true;

                //nombrar controles
                Txt_2.EmbeddedLabelText = "Kilos Producto Seco";
                Txt_3.EmbeddedLabelText = "Merma Kg Secos";
                Txt_4.EmbeddedLabelText = "Kg Secos Fuera de Espec.";
                Txt_5.EmbeddedLabelText = "Kg para Resecar";
                Txt_6.EmbeddedLabelText = "Personal Operativo";

                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta Programada";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Kg Frescos Entrada a secador";
                Txt_Read_5.EmbeddedLabelText = "Cumplimiento de Metas (%)";
                //Txt_Read_6.EmbeddedLabelText = "Kg Secos Meta";
                Txt_Read_7.EmbeddedLabelText = "Relación Fresco-Seco";
                Txt_Read_8.EmbeddedLabelText = "FTT";

                Txt_Read_5.Visible = true;
                Txt_Read_7.Visible = true;
                Txt_Read_8.Visible = true;

                Txt_Read_9.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;
                cb_proceso.Visible = false;
                Txt_Read_6.Visible = false;

                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Txt_6.Enabled = false;

                //Inabilitar controles
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;
                Txt_Read_1.Enabled = false;
                Txt_Read_2.Enabled = false;
                Txt_Read_3.Enabled = false;
                Txt_Read_4.Enabled = false;
                Txt_Read_5.Enabled = false;
                Txt_Read_6.Enabled = false;

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Deshidratado\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");

                // Cargar combobox lote

                query = "SELECT \"ID_Ficha\", \"Lote\" FROM public.\"Ficha\" WHERE \"Terminado_Tunel\" = false and \"Area\" = 'Tunel/Sumergidor' ORDER BY \"Lote\";";
                dbHelper.LoadDataIntoComboBox(query, cb_lote, "Lote", "ID_Ficha");
            }
            if (cb_Area.SelectedIndex == 2)
            {
                //Evaporado
                reiniciarCampos();

                cb_lote.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;
                Txt_Read_5.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";
                Txt_Read_5.EmbeddedLabelText = "Aumento por humedad (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Evaporado\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 3)
            {
                //Grind
                reiniciarCampos();

                cb_lote.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Grind\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 4)
            {
                //Inspeccion
                reiniciarCampos();

                cb_lote.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;
                cb_proceso.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;
                cb_proceso.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Inspeccion\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 5)
            {
                //Empacado
                reiniciarCampos();

                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;
                cb_proceso.Visible = false;
                cb_lote.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Inspeccion\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 6)
            {
                //Polvos
                reiniciarCampos();

                cb_lote.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;
                Txt_6.Visible = true;
                Txt_7.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma PT";
                Txt_5.EmbeddedLabelText = "Personal Operativo";
                Txt_6.EmbeddedLabelText = "Polvo de colector";
                Txt_7.EmbeddedLabelText = "Granulo";

                //hacer invisibles controles
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Txt_6.Enabled = false;
                Txt_7.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Polvos\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 7)
            {
                //Revolturas
                reiniciarCampos();

                cb_lote.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Revolturas\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 8)
            {
                //Máquinas
                reiniciarCampos();

                cb_lote.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;
                Txt_6.Visible = true;
                Txt_7.Visible = true;
                Txt_8.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Piezas Producidas";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";
                Txt_6.EmbeddedLabelText = "Bobina Kg Entrada";
                Txt_7.EmbeddedLabelText = "Bobina Utilizada";
                Txt_8.EmbeddedLabelText = "Bobina Merma";

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";
                Txt_Read_5.EmbeddedLabelText = "Kg entrada (proceso)";

                //hacer invisibles controles
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Txt_6.Enabled = false;
                Txt_7.Enabled = false;
                Txt_8.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;
                cb_proceso.Enabled = false;

                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Maquinas\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
        }

        //// Método que suma los valores de Txt_3 a Txt_8
        //private void CalcularSuma()
        //{
        //    decimal total = 0;
        //    decimal Kg_entrada = 0;
        //    // Lista de los TextBox que quieres sumar
        //    RadTextBox[] cajas = { Txt_3, Txt_4, Txt_5, Txt_6, Txt_7, Txt_8 };

        //    foreach (RadTextBox txt in cajas)
        //    {
        //        if (decimal.TryParse(txt.Text, out decimal valor))
        //        {
        //            total += valor;
        //        }
        //    }
        //    Kg_entrada = Txt_2.Text == string.Empty ? 0 : Convert.ToDecimal(Txt_2.Text);
        //    decimal resultado = Kg_entrada - total;
        //    Txt_Read_4.Text = resultado.ToString("0.00"); // puedes usar "0.##" si quieres decimales
        //}

        // 3) CalcularSuma: no usar Convert.ToDecimal directo — usar TryParse
        private void CalcularSuma()
        {
            decimal total = 0;
            decimal Kg_entrada = 0;
            // Lista de los TextBox que quieres sumar
            RadTextBox[] cajas = { Txt_3, Txt_4, Txt_5, Txt_6, Txt_7, Txt_8 };

            foreach (RadTextBox txt in cajas)
            {
                if (decimal.TryParse(txt.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal valor) ||
                    decimal.TryParse(txt.Text, NumberStyles.Number, CultureInfo.CurrentCulture, out valor))
                {
                    total += valor;
                }
            }

            // Para Txt_2 usar TryParse (evita excepciones mientras el usuario escribe)
            if (!decimal.TryParse(Txt_2.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out Kg_entrada) &&
                !decimal.TryParse(Txt_2.Text, NumberStyles.Number, CultureInfo.CurrentCulture, out Kg_entrada))
            {
                Kg_entrada = 0m;
            }

            decimal resultado = Kg_entrada - total;
            Txt_Read_4.Text = resultado.ToString("0.00", CultureInfo.CurrentCulture);
        }

        private void reiniciarCampos()
        {
            cb_Area.Enabled = true;
            cb_Turno.SelectedIndex = -1;
            cb_Turno.Focus();
            cb_OP.Enabled = true;
            cb_OP.SelectedIndex = -1;
            cb_OP.Focus();
            cb_proceso.SelectedIndex = -1;
            cb_proceso.Focus();
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
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckUser, parametersCheck);
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
            cmb_area.Enabled = true;
            cmb_area.SelectedIndex = -1;
            cmb_area.Focus();
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;
            txt_op.Focus();
        }
        private void limpiarCampos_detalles_OP()
        {
            txt_orden_produc.Text = string.Empty;
            txt_producto.Text = string.Empty;
            txt_medida.Text = string.Empty;
            txt_descripcion.Text = string.Empty;
            txt_especificacion.Text = string.Empty;
            txt_ingredientes.Text = string.Empty;
            txt_humedad.Text = string.Empty;
            cb_Comercio.SelectedIndex = -1;
            cb_Manzana.SelectedIndex = -1;
            txt_analisis.Text = string.Empty;
            txt_area_proceso.Text = string.Empty;
            txt_op_origen.Text = string.Empty;
            txt_destino.Text = string.Empty;
            txt_clacificacion.Text = string.Empty;
            txt_orden_produc.Focus();
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
            // Mostrar todas las filas del DataGridView
            // Mostrar todas las filas del DataGridView Deshidratado
            if (filtroUsuariosActivo_OP_dgv_activo == 1)
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
            // Mostrar todas las filas del DataGridView Revolturas
            if (filtroUsuariosActivo_OP_dgv_activo == 6)
            {
                foreach (DataGridViewRow row in dgv_metas_revolturas.Rows)
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
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;
            txt_op.Focus();
            reiniciar_controles_meta();
            if (cmb_area.SelectedIndex == 0)
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
            if (cmb_area.SelectedIndex == 6)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Meta kg/hr temporada humenda";
                txt_Meta_3.Hint = "Meta kg/hr temporada idónea";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;

                //hacer invisibles controles
                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 6;
            }
            if (cmb_area.SelectedIndex == 7)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Capacidad/(Trompo y/o Máq.)";
                txt_Meta_3.Hint = "Cantidad/(Trompo y/o Máq.)";
                txt_Meta_4.Hint = "Meta Kg por hora";
                txt_Meta_5.Hint = "Kg por pieza";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_5.Enabled = true;

                //hacer visible controles
                txt_Meta_4.Visible = true;
                txt_Meta_5.Visible = true;

                tap_control_metas.SelectedIndex = 7;
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
            id_global_meta_revolturas = string.Empty;
            id_global_meta_polvos = string.Empty;

            //enfocar controles
            cmb_area.Focus();
            txt_op.Focus();
        }

        private void btn_meta_save_Click(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == -1)
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
                            Value = txt_op.Text.ToUpper().Trim()
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

                    // Verificar si ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Deshidratado\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Deshidratado\" (\"OP\", \"No_box_hr\", \"Kg_fresco_hr\", \"Relacion_fr_seco\", \"Kg_seco_hr\", \"Personal_idoneo\") VALUES (@OP, @No_box_h, @Kg_f_h, @Relacion_f_s, @kg_s_h, @Personal_i);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
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
                            Value = txt_op.Text.ToUpper().Trim()
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
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
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
                            Value = txt_op.Text.ToUpper().Trim()
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
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Inspeccion\" (\"OP\", \"Personal_idoneo\", \"Kg_person_hr\", \"Meta_kg_hr_line\") VALUES (@OP, @Personal_I, @Kg_person_hr, @Meta_kg_hr_line);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
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
                            Value = txt_op.Text.ToUpper().Trim()
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
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Evaporado\" (\"OP\", \"Personal_idoneo\", \"Capacidad_trompos\", \"Cantidad_trompos\", \"Meta_kg_hr\") VALUES (@OP, @Personal_I, @Capacidad_trompos, @Cantidad_trompos, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
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
                            Value = txt_op.Text.ToUpper().Trim()
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
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Grind\" (\"OP\", \"Personal_Idoneo\", \"Capacidad_Molino\", \"Cantidad_molinos\", \"Meta_Kg_hr\") VALUES (@OP, @Personal_I, @Capacidad_m, @Cantidad_m, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
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
            ///Revolturas
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
                if (!string.IsNullOrEmpty(id_global_meta_revolturas))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_revolturas);

                    queryInsertUpdate = "UPDATE public.\"Revolturas\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Cap_Trompo_machin\" = @Cap_Trompo_m, \"Canti_Trompo_machin\" = @Cantidad_Trompo_m, \"Meta_Kg_Hr\" = @Meta_kg_hr WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
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
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Revolturas\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Revolturas\" (\"OP\", \"Personal_Idoneo\", \"Cap_Trompo_machin\", \"Canti_Trompo_machin\", \"Meta_Kg_Hr\") VALUES (@OP, @Personal_I, @Capacidad_Trompo_m, @Cantidad_Trompo_m, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_Trompo_m", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_Trompo_m", Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", Convert.ToDecimal(txt_Meta_4.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_revolturas();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_revolturas = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Polvos
            ///
            if (cmb_area.SelectedIndex == 6 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 6)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_polvos))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_polvos);

                    queryInsertUpdate = "UPDATE public.\"Polvos\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Meta_kg_hr_hum\" = @Meta_kg_hr_hum, \"Meta_kg_hr_idon\" = @Meta_kg_hr_idon WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr_hum", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr_idon", NpgsqlTypes.NpgsqlDbType.Numeric)
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
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Polvos\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Polvos\" (\"OP\", \"Personal_Idoneo\", \"Meta_kg_hr_hum\", \"Meta_kg_hr_idon\") VALUES (@OP, @Personal_I, @Meta_kg_hr_hum, @Meta_kg_hr_idon);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Meta_kg_hr_hum", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Meta_kg_hr_idon", Convert.ToDecimal(txt_Meta_3.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                if (result > 0)
                {
                    actualiza_polvos();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_polvos = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Maquinas
            ///
            if (cmb_area.SelectedIndex == 7 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty || txt_Meta_5.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 7)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_maquinas))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = Convert.ToInt32(id_global_meta_maquinas);

                    queryInsertUpdate = "UPDATE public.\"Maquinas\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Cap_Trompo_machin\" = @Cap_Trompo_m, \"Canti_Trompo_machin\" = @Cantidad_Trompo_m, \"Meta_Kg_Hr\" = @Meta_kg_hr, \"Kg_pz\" = @Kg_pz WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
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
                        new NpgsqlParameter("@Kg_pz", NpgsqlTypes.NpgsqlDbType.Numeric)
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
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Maquinas\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Maquinas\" (\"OP\", \"Personal_Idoneo\", \"Cap_Trompo_machin\", \"Canti_Trompo_machin\", \"Meta_Kg_Hr\", \"Kg_pz\") VALUES (@OP, @Personal_I, @Capacidad_Trompo_m, @Cantidad_Trompo_m, @Meta_kg_hr, @Kg_pz);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_Trompo_m", Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_Trompo_m", Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", Convert.ToDecimal(txt_Meta_4.Text)),
                    new NpgsqlParameter("@Kg_pz", Convert.ToDecimal(txt_Meta_5.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_maquinas();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    txt_Meta_5.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_maquinas = string.Empty;

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
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 0;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_deshidratado = dgv_metas_des.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_des.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_des.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_2.Text = dgv_metas_des.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_3.Text = dgv_metas_des.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_4.Text = dgv_metas_des.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_5.Text = dgv_metas_des.Rows[e.RowIndex].Cells[6].Value.ToString();



                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

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
            if (cmb_area.SelectedIndex == 3 || cmb_area.SelectedIndex == 4 || cmb_area.SelectedIndex == 5 || cmb_area.SelectedIndex == 6 || cmb_area.SelectedIndex == 7)
            {
                btn_meta_save.Enabled = true;
                btn_meta_cancel.Enabled = true;
                btn_meta_delete.Enabled = false;
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                if (cmb_area.SelectedIndex == 7) { txt_Meta_5.Enabled = true; }
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
            ///Revolturas
            ///
            if (cmb_area.SelectedIndex == 5 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Revolturas?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_revolturas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Revolturas""
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
                    actualiza_revolturas();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_revolturas = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Polvos
            ///
            if (cmb_area.SelectedIndex == 5 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Revolturas?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_revolturas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Polvos""
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
                    actualiza_polvos();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_polvos = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Maquinas
            ///
            if (cmb_area.SelectedIndex == 7 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Maquinas?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = Convert.ToInt32(id_global_meta_maquinas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Maquinas""
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
                    actualiza_maquinas();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    txt_Meta_5.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_maquinas = string.Empty;

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
                kg_seco_hr();
            }
            if (cmb_area.SelectedIndex == 1 || cmb_area.SelectedIndex == 2)
            {
                meta_hr_linea();
            }
        }
        private void kg_seco_hr()
        {
            if (decimal.TryParse(txt_Meta_3.Text, out decimal valor2) && decimal.TryParse(txt_Meta_2.Text, out decimal valor1))
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

        private void txt_Meta_3_TextChanged(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == 0)
            {
                kg_seco_hr();
            }
            if (cmb_area.SelectedIndex == 3 || cmb_area.SelectedIndex == 4 || cmb_area.SelectedIndex == 5 || cmb_area.SelectedIndex == 7)
            {
                meta_hr();
            }
        }
        private void meta_hr()
        {
            if (decimal.TryParse(txt_Meta_3.Text, out decimal valor2) && decimal.TryParse(txt_Meta_2.Text, out decimal valor1))
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
        private void txt_Meta_2_TextChanged(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == 1 || cmb_area.SelectedIndex == 2)
            {
                meta_hr_linea();
            }
            if (cmb_area.SelectedIndex == 3 || cmb_area.SelectedIndex == 4 || cmb_area.SelectedIndex == 5 || cmb_area.SelectedIndex == 7)
            {
                meta_hr();
            }
        }

        private void meta_hr_linea()
        {
            if (decimal.TryParse(txt_Meta_2.Text, out decimal valor2) && decimal.TryParse(txt_Meta_1.Text, out decimal valor1))
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

        private void dgv_metas_emp_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 1;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_empacado = dgv_metas_emp.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_2.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[3].Value.ToString();

                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

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
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 2;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_inspec = dgv_metas_insp.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_2.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[3].Value.ToString();

                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

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
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 3;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_evaporado = dgv_metas_Eva.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

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
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 4;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_grind = dgv_metas_Grind.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

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
                if (cmb_area_search.SelectedIndex == 0)
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
            ///Revolturas
            ///
            if (cmb_area_search.SelectedIndex == 5)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 6;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_revolturas.Rows)
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
                dgv_metas_revolturas.ClearSelection();
                dgv_metas_revolturas.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_revolturas.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_revolturas.Rows)
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

        private void calcular_meta_programada()
        {
            double hr_programada, meta_programada, meta_x_hr;

            hr_programada = double.TryParse(Txt_Read_1.Text, out hr_programada) ? hr_programada : 0;
            meta_x_hr = double.TryParse(Txt_meta.Text, out meta_x_hr) ? meta_x_hr : 0;
            meta_programada = hr_programada * meta_x_hr;
            Txt_Read_2.Text = meta_programada.ToString("0.##");
        }
        private void calcular_meta_programada_inspec_emp()
        {
            double hr_programada, meta_programada, meta_x_hr, personal_o;

            hr_programada = double.TryParse(Txt_Read_1.Text, out hr_programada) ? hr_programada : 0;
            meta_x_hr = double.TryParse(Txt_meta.Text, out meta_x_hr) ? meta_x_hr : 0;
            personal_o = double.TryParse(Txt_5.Text, out personal_o) ? personal_o : 0;
            meta_programada = meta_x_hr * personal_o * hr_programada;
            Txt_Read_2.Text = meta_programada.ToString("0.##");
        }

        private void calcular_turno()
        {
            if (cb_Area.SelectedIndex != -1)
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

                        //if (cb_Area.SelectedIndex == 1)
                        //{
                        //    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                        //    string valorBuscado2 = cb_OP.Text;
                        //    // Consulta para buscar donde OP = valor_buscado
                        //    string query2 = "SELECT \"Kg_seco_hr\" FROM public.\"Deshidratado\" WHERE \"OP\" = @valorBuscado;";

                        //    // Crear parámetro
                        //    NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                        //    {
                        //        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado2 }
                        //    };

                        //    // Ejecutar consulta
                        //    System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                        //    decimal horas_prog = Convert.ToDecimal(Txt_Read_1.Text);
                        //    // Variable string donde guardar el resultado
                        //    decimal kg_seco_hr, kg_secos_meta;

                        //    // Verificar si se encontraron resultados
                        //    if (dt2 != null && dt2.Rows.Count > 0)
                        //    {
                        //        kg_seco_hr = Convert.ToDecimal(dt2.Rows[0]["Kg_seco_hr"].ToString());
                        //        kg_secos_meta = kg_seco_hr * horas_prog;
                        //        Txt_Read_6.Text = kg_secos_meta.ToString("0.##");
                        //    }
                        //}

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
                        if (!string.IsNullOrEmpty(Txt_meta.Text) && cb_Area.SelectedIndex != 4)
                        {
                            calcular_meta_programada();
                        }
                        // Mostrar el resultado con descuento en Txt_Read_3
                        Txt_Read_3.Text = diferenciaConDescuento.TotalHours.ToString("0.##");
                    }
                    btn_save_ficha.Enabled = true;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show($"Error en cálculo: {ex.Message}");
                    if (ex is FormatException)
                    {
                        //MetroFramework.MetroMessageBox.Show(this, "Formato de hora inválido. Asegúrese de usar HH:mm.",
                        //                                    "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        btn_save_ficha.Enabled = false;
                    }
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
                Mask_txt_hr1.Enabled = true;
                Mask_txt_hr2.Enabled = true;
                txt_Tiempo_energia.Enabled = true;
                txt_Tiempo_comida.Enabled = true;
                dgv_mecanico.Enabled = true;
                dgv_operativo.Enabled = true;
                if (cb_Area.SelectedIndex == 0)
                {
                    if (!editar) { Txt_1.Enabled = true; } else { Txt_1.Enabled = false; }
                    Txt_2.Enabled = true;
                    Txt_3.Enabled = true;
                    Txt_4.Enabled = true;
                    Txt_5.Enabled = true;
                    Txt_6.Enabled = true;
                    Txt_7.Enabled = true;
                    Txt_8.Enabled = true;
                    Txt_9.Enabled = true;
                    Txt_10.Enabled = true;
                    Txt_11.Enabled = true;

                }
                if (cb_Area.SelectedIndex == 1)
                {
                    Txt_2.Enabled = true;
                    Txt_3.Enabled = true;
                    Txt_4.Enabled = true;
                    Txt_5.Enabled = true;
                    Txt_6.Enabled = true;
                }
                if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7)
                {
                    Txt_1.Enabled = true;
                    Txt_2.Enabled = true;
                    Txt_3.Enabled = true;
                    Txt_4.Enabled = true;
                    Txt_5.Enabled = true;
                    if (cb_Area.SelectedIndex == 4)
                    {
                        cb_proceso.Enabled = true;
                    }
                    if (cb_Area.SelectedIndex == 6)
                    {
                        Txt_6.Enabled = true;
                        Txt_7.Enabled = true;
                    }
                }
                if (cb_Area.SelectedIndex == 8)
                {
                    Txt_1.Enabled = true;
                    Txt_2.Enabled = true;
                    Txt_3.Enabled = true;
                    Txt_4.Enabled = true;
                    Txt_5.Enabled = true;
                    Txt_6.Enabled = true;
                    Txt_7.Enabled = true;
                    Txt_8.Enabled = true;
                }
            }
        }

        private void btn_save_ficha_Click(object sender, EventArgs e)
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            ////TUNEL/ SUMERGIDOR
            if (cb_Area.SelectedIndex == 0)
            {
                try
                {
                    int idUsuarioActual = id_user;

                    if (cb_OP.SelectedIndex != -1 && !string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) && !string.IsNullOrEmpty(Txt_6.Text) && !string.IsNullOrEmpty(Txt_7.Text) && !string.IsNullOrEmpty(Txt_8.Text) && !string.IsNullOrEmpty(Txt_9.Text) && !string.IsNullOrEmpty(Txt_10.Text))
                    {
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
                        decimal meta_kg = Convert.ToDecimal(Txt_Read_2.Text);
                        string area = cb_Area.Text;
                        decimal meta = Convert.ToDecimal(Txt_meta.Text);

                        // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                        TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                        TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);

                        if (editar)
                        {
                            updateFicha(dbHelper, idUsuarioActual, fecha, turno, null, op,
                            kgEnterProceso, kgFrescosEnterSe, mermaCanica, mermaPodrido, mermaTina, mermaPiso,
                            mermaCanaletas, mermaLavadoBandas, cascaraCarrete, hrInicio, hrFin, personal_Op, hr_pro, hr_efec, meta_kg, null, meta, 0);

                        }
                        else
                        {
                            // Verificar si ya existe
                            string queryChecklote = "SELECT COUNT(*) FROM public.\"Ficha\" WHERE \"Lote\"  ILIKE @Lote;";
                            NpgsqlParameter[] parametersLote = new NpgsqlParameter[]
                            {
                            new NpgsqlParameter("@Lote", lote),
                            };
                            System.Data.DataTable dtLote = dbHelper.ExecuteSelectQuery(queryChecklote, parametersLote);
                            if (dtLote != null && dtLote.Rows.Count > 0 && Convert.ToInt32(dtLote.Rows[0][0]) > 0)
                            {
                                MetroFramework.MetroMessageBox.Show(this, "El Lote ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            // Insertar en tabla Ficha y obtener el ID_Ficha generado
                            int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, lote, op,
                                kgEnterProceso, kgFrescosEnterSe, mermaCanica, mermaPodrido, mermaTina, mermaPiso,
                                mermaCanaletas, mermaLavadoBandas, cascaraCarrete, hrInicio, hrFin, personal_Op, hr_pro, hr_efec, meta_kg, area, meta, 0);

                            

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
                    }
                    else
                    {
                        MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                            "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, $"Error al guardar: {ex.Message}",
                                                            "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (cb_Area.SelectedIndex == 1)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && cb_lote.SelectedIndex != -1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) && !string.IsNullOrEmpty(Txt_6.Text))
                {
                    //Despegue
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal KgProdSeco = Convert.ToDecimal(Txt_2.Text);
                    decimal MermaKgSeco = Convert.ToDecimal(Txt_3.Text);
                    decimal KgFueraSpec = Convert.ToDecimal(Txt_4.Text);
                    decimal KgResecar = Convert.ToDecimal(Txt_5.Text);
                    int PersonalOpe = Convert.ToInt32(Txt_6.Text);
                    decimal hr_programadas = Convert.ToDecimal(Txt_Read_1.Text);
                    decimal meta_kg = Convert.ToDecimal(Txt_Read_2.Text);
                    decimal hr_efec = Convert.ToDecimal(Txt_Read_3.Text);
                    string textoPorcentCumplimiento = Txt_Read_5.Text.Replace("%", "").Trim();
                    decimal PorcentCumplimiento = Convert.ToDecimal(textoPorcentCumplimiento) / 100m;
                    //decimal Kg_secos_meta = Convert.ToDecimal(Txt_Read_6.Text);
                    decimal Relacion_Fresco_seco = Convert.ToDecimal(Txt_Read_7.Text);
                    string textoFTT = Txt_Read_8.Text.Replace("%", "").Trim();
                    decimal FTT = Convert.ToDecimal(textoFTT) / 100m;
                    string area = cb_Area.Text;
                    decimal meta = Convert.ToDecimal(Txt_meta.Text);
                    string lote = cb_lote.Text;
                    decimal kgFrescosEnterSe = Convert.ToDecimal(Txt_Read_4.Text);

                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);

                    if (editar)
                    {
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                            KgFueraSpec, KgResecar, PorcentCumplimiento, 0, Relacion_Fresco_seco, FTT,
                            KgProdSeco, MermaKgSeco, kgFrescosEnterSe, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, 0);
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, lote,
                            KgFueraSpec, KgResecar, PorcentCumplimiento, 0, Relacion_Fresco_seco, FTT,
                            KgProdSeco, MermaKgSeco, kgFrescosEnterSe, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0);

                        if (idFicha > 0)
                        {
                            // Ahora hacer el UPDATE usando el parámetro @loteterminado
                            string updateQuery = @"UPDATE public.""Ficha"" 
                              SET ""Terminado_Tunel"" = true 
                              WHERE ""Lote"" = @loteterminado;";

                            NpgsqlParameter[] updateParameters = new NpgsqlParameter[]
                            {
                                new NpgsqlParameter("@loteterminado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = lote ?? (object)DBNull.Value }
                            };

                            int updateResult = dbHelper.ExecuteNonQuery(updateQuery, updateParameters);

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
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 2)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text))
                {
                    //Evaporado
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal KgEntrada = Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = Convert.ToDecimal(Txt_Read_3.Text);
                    string textoPorcent_Aumento_Hume = Txt_Read_5.Text.Replace("%", "").Trim();
                    string area = cb_Area.Text;
                    decimal Porcent_Aumento_Hume = Convert.ToDecimal(textoPorcent_Aumento_Hume) / 100m;
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    decimal meta = Convert.ToDecimal(Txt_meta.Text);

                    if (editar)
                    {
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        Porcent_Aumento_Hume, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, 0);
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        Porcent_Aumento_Hume, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0);

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
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 7)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text))
                {
                    //Grind
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal KgEntrada = Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = Convert.ToDecimal(Txt_Read_3.Text);
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    string area = cb_Area.Text;
                    decimal meta = Convert.ToDecimal(Txt_meta.Text);


                    if (editar)
                    {
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        0, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, 0);
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        0, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0);
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
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 4)
            {
                if (cb_proceso.SelectedIndex != -1 && !string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text))
                {
                    //Inspeccion
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal KgEntrada = Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = Convert.ToDecimal(Txt_Read_3.Text);
                    string proceso = cb_proceso.Text;
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    string area = cb_Area.Text;
                    decimal meta = Convert.ToDecimal(Txt_meta.Text);

                    if (editar)
                    {
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, proceso,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        0, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, 0);
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, proceso,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        0, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0);
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
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 6)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) && !string.IsNullOrEmpty(Txt_6.Text) && !string.IsNullOrEmpty(Txt_7.Text))
                {
                    //Polvos
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal KgEntrada = Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = Convert.ToDecimal(Txt_Read_3.Text);
                    decimal polvo_colector = Convert.ToDecimal(Txt_6.Text);
                    decimal Granulo = Convert.ToDecimal(Txt_7.Text);
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    string area = cb_Area.Text;
                    decimal meta = Convert.ToDecimal(Txt_meta.Text);

                    if (editar)
                    {
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        polvo_colector, Granulo, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0);
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, null,
                            KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                            polvo_colector, Granulo, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0);

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
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 8)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) && !string.IsNullOrEmpty(Txt_6.Text) && !string.IsNullOrEmpty(Txt_7.Text) && !string.IsNullOrEmpty(Txt_8.Text))
                {
                    //Revolturas
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal Pz_producidas = Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = Convert.ToDecimal(Txt_Read_3.Text);
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    string area = cb_Area.Text;
                    decimal kg_entrada = Convert.ToDecimal(Txt_Read_5.Text);
                    decimal bobina_entrada = Convert.ToDecimal(Txt_6.Text);
                    decimal bobina_utilizada = Convert.ToDecimal(Txt_7.Text);
                    decimal bobina_merma = Convert.ToDecimal(Txt_8.Text);
                    decimal meta = Convert.ToDecimal(Txt_meta.Text);

                    if (editar)
                    {
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        Pz_producidas, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        kg_entrada, bobina_entrada, bobina_utilizada, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, bobina_merma);
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, null,
                            Pz_producidas, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                            kg_entrada, bobina_entrada, bobina_utilizada, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, bobina_merma);
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
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void updateFicha(DatabaseHelper dbHelper, int idUsuario, DateTime fecha,
        int turno, string var1, string var2, decimal var3, decimal var4,
        decimal var5, decimal var6, decimal var7, decimal var8,
        decimal var9, decimal var10, decimal var11,
        TimeSpan hrInicio, TimeSpan hrFin, int personal_O, decimal hr_programadas, decimal hr_efectivas,
        decimal meta_kg, string area_f, decimal metaHr, decimal var12)
        {
            DatabaseHelper dbHelper2 = new DatabaseHelper(connectionString);
            NpgsqlParameter[] parameters = new NpgsqlParameter[] { };
            string queryInsertUpdate = string.Empty;
            int result = 0;
            // Convertir el ID a entero ANTES de crear el parámetro
            int idFicha = Convert.ToInt32(id_global_ficha);
            if (cb_Area.SelectedIndex == 0)
            {
                // actualizar

                queryInsertUpdate = "UPDATE public.\"Ficha\" SET \"ID_user\" = @id_user, \"Fecha\" = @Fecha, \"Turno\" = @Turno, \"OP\" = @OP," +
                    " \"Kg_enter_proceso\" = @Kg_enter_proceso, \"kg_frescos_enter_se\" = @kg_frescos_enter_se, \"Merma_canica\" = @Merma_canica, \"Merma_podrido\" = @Merma_podrido," +
                    " \"Merma_tina\" = @Merma_tina, \"Merma_piso\" = @Merma_piso, \"Merma_canaletas\" = @Merma_canaletas, \"Merma_lavado_bandas\" = @Merma_lavado_bandas," +
                    " \"Cascara_carrete\" = @Cascara_carrete, \"Hr_inicio\" = @Hr_inicio, \"Hr_fin\" = @Hr_fin, \"Hr_programadas\" = @Hr_programadas, \"Personal_Operativo\" = @Personal_Operativo," +
                    " \"Hr_efectivas\" = @Hr_efectivas, \"Kg_meta\" = @Kg_meta, \"MetaHr\" = @MetaHr WHERE \"ID_Ficha\" = @ID_Ficha;";

                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@Turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Merma_canica", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_podrido", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@merma_tina", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var7 },
                    new NpgsqlParameter("@Merma_piso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Merma_canaletas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Merma_lavado_bandas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Cascara_carrete", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@Hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha}
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 1)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @op, ""Kg_prod_seco"" = @Kg_prod_seco,
                                    ""Merma_kg"" = @Merma_kg_seco, ""Kg_fuera_espec"" = @Kg_fuera_spec, ""Kg_resecar"" = @Kg_resecar, ""Personal_Operativo"" = @Personal_Operativo, 
                                    ""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, ""porcent_cump_meta"" = @Porcent_cumplimiento, ""Kg_meta"" = @Kg_meta,
                                    ""Relacion_Fr_seco"" = @Relacion_fresco_seco, ""FTT"" = @FTT, ""Hr_inicio"" = @hr_inicio, ""Hr_fin"" = @hr_fin, ""MetaHr"" = @MetaHr, ""kg_frescos_enter_se"" = @kg_frescos_enter_se
                                    WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@op", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_prod_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Kg_fuera_spec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Merma_kg_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Kg_resecar", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Porcent_cumplimiento", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Relacion_fresco_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var7 },
                    new NpgsqlParameter("@FTT", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha},
                    new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 }
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 2)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @OP, ""Kg_meta"" = @Kg_meta, 
                    ""porcent_cump_meta"" = @porcent_cump_meta, ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Kg_prod_term"" = @Kg_prod_term, ""Kg_fuera_espec"" = @Kg_fuera_espec, 
                    ""Merma_kg"" = @Merma_kg,""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, ""Personal_Operativo"" = @Personal_Operativo, 
                    ""porcent_aumento_hum"" = @porcent_aumento_hum, ""Hr_inicio"" = @Hr_inicio, ""Hr_fin"" = @Hr_fin, ""MetaHr"" = @MetaHr WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@porcent_aumento_hum", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha}
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 7)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @OP, ""Kg_meta"" = @Kg_meta, 
                                    ""porcent_cump_meta"" = @porcent_cump_meta, ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Kg_prod_term"" = @Kg_prod_term, 
                                    ""Kg_fuera_espec"" = @Kg_fuera_espec, ""Merma_kg"" = @Merma_kg,""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, 
                                    ""Personal_Operativo"" = @Personal_Operativo, ""Hr_inicio"" = @Hr_inicio, ""Hr_fin"" = @Hr_fin, ""MetaHr"" = @MetaHr WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha}
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 4)
            {

                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Proceso"" = @Proceso, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @op, ""Kg_meta"" = @Kg_meta, 
                                    ""porcent_cump_meta"" = @porcent_cump_meta, ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Kg_prod_term"" = @Kg_prod_term, 
                                    ""Kg_fuera_espec"" = @Kg_fuera_espec, ""Merma_kg"" = @Merma_kg,""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, 
                                    ""Personal_Operativo"" = @Personal_Operativo, ""Hr_inicio"" = @Hr_inicio, ""Hr_fin"" = @Hr_fin, ""MetaHr"" = @MetaHr WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@Proceso", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha}
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 6)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @OP, ""Kg_meta"" = @Kg_meta, 
                                    ""porcent_cump_meta"" = @porcent_cump_meta, ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Kg_prod_term"" = @Kg_prod_term, 
                                    ""Kg_fuera_espec"" = @Kg_fuera_espec, ""Merma_kg"" = @Merma_kg, ""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, 
                                    ""Personal_Operativo"" = @Personal_Operativo, ""Hr_inicio"" = @Hr_inicio, ""Hr_fin"" = @Hr_fin, ""Polvo_colector"" = @Polvo_Colector, 
                                    ""Granulo"" = @Granulo, ""MetaHr"" = @MetaHr WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Polvo_Colector", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Granulo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha}
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 8)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @OP, ""Pz_prod"" = @Pz_prod, 
                    ""Kg_prod_term"" = @Kg_prod_term, ""Kg_fuera_espec"" = @Kg_fuera_espec, ""Merma_kg"" = @Merma_kg, ""Kg_meta"" = @Kg_meta, ""porcent_cump_meta"" = @porcent_cump_meta, 
                    ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Bobina_kg_enter"" = @Bobina_kg_enter, ""Bobina_utilizada"" = @Bobina_utilizada, ""Hr_inicio"" = @Hr_inicio, 
                    ""Hr_fin"" = @Hr_fin, ""Personal_Operativo"" = @Personal_Operativo, ""Hr_programadas"" = @Hr_programadas, ""Bobina_merma"" = @Bobina_merma, 
                    ""Hr_efectivas"" = @Hr_efectivas, ""MetaHr"" = @MetaHr WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Pz_prod", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Bobina_kg_enter", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Bobina_utilizada", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Bobina_merma", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var12 },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha}
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }

            // Actualizar los tiempos muertos asociados
            UpdateTiemposMuertos(dbHelper);
            if (result > 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                            "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cb_Area.SelectedIndex = -1;
                reiniciarCampos();
                cb_Area.Focus();
                editar = false; // Reiniciar el estado de edición
                id_global_ficha = string.Empty; // Limpiar el ID global
            }
        }

        private int InsertarFichaYRetornarID(DatabaseHelper dbHelper, int idUsuario, DateTime fecha,
        int turno, string var1, string var2, decimal var3, decimal var4,
        decimal var5, decimal var6, decimal var7, decimal var8,
        decimal var9, decimal var10, decimal var11,
        TimeSpan hrInicio, TimeSpan hrFin, int personal_O, decimal hr_programadas, decimal hr_efectivas,
        decimal meta_kg, string area_f, decimal metaHr, decimal var12)
        {
            string query = String.Empty;
            NpgsqlParameter[] parameters = new NpgsqlParameter[] { };
            if (cb_Area.SelectedIndex == 0)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""Lote"", ""OP"",
                    ""Kg_enter_proceso"", ""kg_frescos_enter_se"", ""Merma_canica"",
                    ""Merma_podrido"", ""Merma_tina"", ""Merma_piso"", ""Merma_canaletas"",
                    ""Merma_lavado_bandas"", ""Cascara_carrete"", ""Hr_inicio"", ""Hr_fin"", 
                    ""Hr_programadas"", ""Personal_Operativo"", ""Hr_efectivas"", ""Kg_meta"", ""Area"", ""MetaHr""
                ) VALUES (
                    @id_user, @fecha, @turno, @lote, @op,
                    @kg_enter_proceso, @kg_frescos_enter_se, @merma_canica,
                    @merma_podrido, @merma_tina, @merma_piso, @merma_canaletas,
                    @merma_lavado_bandas, @cascara_carrete, @hr_inicio, @hr_fin, @hr_prog, @Personal_Op, @Hr_efec, @Kg_meta, @Area, @MetaHr
                ) RETURNING ""ID_Ficha"";";

                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@lote", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@op", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@merma_canica", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@merma_podrido", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@merma_tina", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var7 },
                    new NpgsqlParameter("@merma_piso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@merma_canaletas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@merma_lavado_bandas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@cascara_carrete", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@hr_prog", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Personal_Op", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                    new NpgsqlParameter("@Hr_efec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f ?? (object)DBNull.Value },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr }
                };
            }
            if (cb_Area.SelectedIndex == 1)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Lote"", ""Kg_prod_seco"", ""Merma_kg"", ""Kg_fuera_espec"", ""Kg_resecar"", 
                    ""Personal_Operativo"", ""Hr_programadas"", ""Hr_efectivas"", ""porcent_cump_meta"", ""Kg_meta"",
                    ""Relacion_Fr_seco"", ""FTT"", ""Hr_inicio"", ""Hr_fin"", ""Area"", ""MetaHr"", ""kg_frescos_enter_se""
                ) VALUES (
                    @id_user, @fecha, @turno, @op, @Lote, @Kg_prod_seco, @Merma_kg_seco, @Kg_fuera_spec,
                    @Kg_resecar, @Personal_Operativo, @Hr_programadas,
                    @Hr_efectivas, @Porcent_cumplimiento, @Kg_meta,
                    @Relacion_fresco_seco, @FTT, @hr_inicio, @hr_fin, @Area, @MetaHr, @kg_frescos_enter_se
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@op", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Lote", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_prod_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Kg_fuera_spec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Merma_kg_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_resecar", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Porcent_cumplimiento", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Relacion_fresco_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var7 },
                    new NpgsqlParameter("@FTT", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 }
                };
            }
            if (cb_Area.SelectedIndex == 2)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""porcent_cump_meta"", ""Kg_enter_proceso"", 
                    ""Kg_prod_term"", ""Kg_fuera_espec"", ""Merma_kg"",""Hr_programadas"", ""Hr_efectivas"", ""Personal_Operativo"", 
                    ""porcent_aumento_hum"", ""Hr_inicio"", ""Hr_fin"", ""Area"", ""MetaHr""
                ) VALUES (
                    @id_user, @fecha, @turno, @OP, @Kg_meta, @porcent_cump_meta,
                    @Kg_enter_proceso, @Kg_prod_term, @Kg_fuera_espec,
                    @Merma_kg, @Hr_programadas, @Hr_efectivas, @Personal_Operativo,
                    @porcent_aumento_hum, @Hr_inicio, @Hr_fin, @Area, @MetaHr
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@porcent_aumento_hum", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr }
                };
            }
            if (cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 7 || cb_Area.SelectedIndex == 5)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""porcent_cump_meta"", ""Kg_enter_proceso"", ""Kg_prod_term"", 
                    ""Kg_fuera_espec"", ""Merma_kg"",""Hr_programadas"", ""Hr_efectivas"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", ""Area"", ""MetaHr""
                ) VALUES (
                    @id_user, @fecha, @turno, @OP, @Kg_meta, @porcent_cump_meta, @Kg_enter_proceso, @Kg_prod_term, 
                    @Kg_fuera_espec,@Merma_kg, @Hr_programadas, @Hr_efectivas, @Personal_Operativo, @Hr_inicio, @Hr_fin, @Area, @MetaHr
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr }
                };
            }
            if (cb_Area.SelectedIndex == 4)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Proceso"", ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""porcent_cump_meta"", ""Kg_enter_proceso"", ""Kg_prod_term"", 
                    ""Kg_fuera_espec"", ""Merma_kg"",""Hr_programadas"", ""Hr_efectivas"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", ""Area"", ""MetaHr""
                ) VALUES (
                    @id_user, @Proceso, @fecha, @turno, @OP, @Kg_meta, @porcent_cump_meta, @Kg_enter_proceso, @Kg_prod_term, 
                    @Kg_fuera_espec,@Merma_kg, @Hr_programadas, @Hr_efectivas, @Personal_Operativo, @Hr_inicio, @Hr_fin, @Area, @MetaHr
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@Proceso", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr }
                };
            }
            if (cb_Area.SelectedIndex == 6)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""porcent_cump_meta"", ""Kg_enter_proceso"", ""Kg_prod_term"", ""Kg_fuera_espec"", 
                    ""Merma_kg"",""Hr_programadas"", ""Hr_efectivas"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", ""Polvo_colector"", ""Granulo"", ""Area"", ""MetaHr""
                ) VALUES (
                    @id_user, @fecha, @turno, @OP, @Kg_meta, @porcent_cump_meta, @Kg_enter_proceso, @Kg_prod_term, 
                    @Kg_fuera_espec,@Merma_kg, @Hr_programadas, @Hr_efectivas, @Personal_Operativo, @Hr_inicio, @Hr_fin, @Polvo_Colector, @Granulo, @Area, @MetaHr
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Polvo_Colector", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Granulo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr }
                };
            }
            if (cb_Area.SelectedIndex == 8)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Pz_prod"", ""Kg_prod_term"", ""Kg_fuera_espec"", ""Merma_kg"", ""Kg_meta"", ""porcent_cump_meta"", 
                    ""Kg_enter_proceso"", ""Bobina_kg_enter"", ""Bobina_utilizada"", ""Hr_inicio"", ""Hr_fin"", ""Personal_Operativo"", ""Hr_programadas"", ""Bobina_merma"", 
                    ""Hr_efectivas"", ""Area"", ""MetaHr""
                ) VALUES (
                    @id_user, @fecha, @turno, @OP, @Pz_prod, @Kg_prod_term, @Kg_fuera_espec, @Merma_kg, @Kg_meta, @porcent_cump_meta, @Kg_enter_proceso, @Bobina_kg_enter,  
                    @Bobina_utilizada, @Hr_inicio, @Hr_fin, @Personal_Operativo, @Hr_programadas, @Bobina_merma, @Hr_efectivas, @Area, @MetaHr
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Pz_prod", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Bobina_kg_enter", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Bobina_utilizada", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Bobina_merma", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var12 },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr }
                };
            }
            return dbHelper.ExecuteScalarInt(query, parameters);
        }

        private void UpdateTiemposMuertos(DatabaseHelper dbHelper)
        {
            // Insertar tiempos muertos mecánicos desde DataGridView
            UpdateTiemposMuertosMecanicos(dbHelper);

            // Insertar tiempos muertos operativos desde DataGridView
            UpdateTiemposMuertosOperativos(dbHelper);

            // Insertar tiempo muerto comida desde TextBox
            UpdateTiempoMuertoComida(dbHelper);

            // Insertar tiempo muerto energía desde TextBox
            UpdateTiempoMuertoEnergia(dbHelper);
        }
        private void UpdateTiemposMuertosMecanicos(DatabaseHelper dbHelper)
        {
            int id_ficha_int = Convert.ToInt32(id_global_ficha);

            // Primero eliminamos todos los registros existentes para esta ficha
            string deleteQuery = @"DELETE FROM public.""Tiempo_muerto_Mecanico""
                         WHERE ""ID_Ficha"" = @id_ficha;";

            NpgsqlParameter[] deleteParams = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = id_ficha_int }
            };

            dbHelper.ExecuteNonQuery(deleteQuery, deleteParams);

            InsertarTiemposMuertosMecanicos(dbHelper, id_ficha_int);
        }
        private void UpdateTiemposMuertosOperativos(DatabaseHelper dbHelper)
        {
            int id_ficha_int = Convert.ToInt32(id_global_ficha);

            // Primero eliminamos todos los registros existentes para esta ficha
            string deleteQuery = @"DELETE FROM public.""Tiempo_Muerto_Operativo"" 
                         WHERE ""ID_Ficha"" = @id_ficha;";

            NpgsqlParameter[] deleteParams = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = id_ficha_int }
            };

            dbHelper.ExecuteNonQuery(deleteQuery, deleteParams);

            // Luego insertamos todos los registros del DataGridView
            InsertarTiemposMuertosOperativos(dbHelper, id_ficha_int);
        }
        private void UpdateTiempoMuertoComida(DatabaseHelper dbHelper)
        {
            int id_ficha_int = Convert.ToInt32(id_global_ficha);
            if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text))
            {
                decimal minutosDetenidos = Convert.ToDecimal(txt_Tiempo_comida.Text);

                string query = @"UPDATE public.""Tiempo_Muerto_Comida"" SET ""Minutos_Detenidos"" = @minutos_detenidos WHERE ""ID_Ficha"" = @id_ficha;";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = id_ficha_int },
                    new NpgsqlParameter("@minutos_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minutosDetenidos }
                };

                dbHelper.ExecuteNonQuery(query, parameters);
            }
        }
        private void UpdateTiempoMuertoEnergia(DatabaseHelper dbHelper)
        {
            int id_ficha_int = Convert.ToInt32(id_global_ficha);
            if (!string.IsNullOrEmpty(txt_Tiempo_energia.Text))
            {
                decimal minutosDetenidos = Convert.ToDecimal(txt_Tiempo_energia.Text);

                string query = @"UPDATE public.""Tiempo_Muerto_Energia"" SET ""Minutos_Detenidos"" = @minutos_detenidos WHERE ""ID_Ficha"" = @id_ficha;";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = id_ficha_int },
                    new NpgsqlParameter("@minutos_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minutosDetenidos }
                };

                dbHelper.ExecuteNonQuery(query, parameters);
            }
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
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
            if ((cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                porcentaje_logrado_planeacion();
            }
        }

        private void Txt_4_TextChanged(object sender, EventArgs e)
        {
            var tb = (RadTextBox)sender;
            string original = tb.Text;
            string saneado = SanitizeNumericText(original);

            if (saneado != original)
            {
                int sel = tb.SelectionStart;
                tb.Text = saneado;
                // Ajusta la posición del cursor (no exceder la longitud)
                tb.SelectionStart = Math.Min(sel, tb.Text.Length);
            }

            if (cb_Area.SelectedIndex == 0)
            {
                CalcularSuma();
            }
            if (cb_Area.SelectedIndex == 1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_cumplimiento_metas();
                Ftt_metodo();
                Relacion_Fresco_seco();
            }
        }

        private void Txt_5_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
            if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
            {
                calcular_meta_programada_inspec_emp();
                porcentaje_logrado_planeacion();
            }
        }

        private void Txt_6_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
        }


        private void Txt_7_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
        }

        private void Txt_8_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
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
            var tb = (RadTextBox)sender;
            string original = tb.Text;
            string saneado = SanitizeNumericText(original);

            if (saneado != original)
            {
                int sel = tb.SelectionStart;
                tb.Text = saneado;
                // Ajusta la posición del cursor (no exceder la longitud)
                tb.SelectionStart = Math.Min(sel, tb.Text.Length);
            }

            if (cb_Area.SelectedIndex == 0)
            {
                CalcularSuma();
            }
            if (cb_Area.SelectedIndex == 1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_cumplimiento_metas();
                Ftt_metodo();
                Relacion_Fresco_seco();
            }
            if ((cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 6) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_logrado_planeacion();
            }
            if (cb_Area.SelectedIndex == 2 && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text))
            {
                porcentaje_aumento_humedad();
            }
            if ((cb_Area.SelectedIndex == 8 || cb_Area.SelectedIndex == 7) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_logrado_planeacion_platinum();
            }
        }

        // Función que deja sólo dígitos y un solo punto (mantiene el primer punto)
        private string SanitizeNumericText(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            // Elimina todo excepto dígitos y puntos
            string onlyDigitsAndDots = Regex.Replace(input, @"[^0-9.]", "");

            int firstDot = onlyDigitsAndDots.IndexOf('.');
            if (firstDot >= 0)
            {
                // deja sólo el primer punto
                string before = onlyDotsAndTrim(onlyDigitsAndDots.Substring(0, firstDot + 1));
                string after = onlyDigitsAndDots.Substring(firstDot + 1).Replace(".", ""); // eliminar puntos siguientes
                return before + after;
            }
            else
            {
                return onlyDotsAndTrim(onlyDigitsAndDots);
            }

            // local helper
            string onlyDotsAndTrim(string s) => s; // ya está limpio en este flujo
        }
        private void cb_OP_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                cb_Turno.Enabled = true;
            }
            if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7 || cb_Area.SelectedIndex == 8)
            {
                cb_Turno.Enabled = true;
            }

            buscar_Meta_hr();
        }

        private void buscar_Meta_hr()
        {
            //MessageBox.Show("entró");
            string valorBuscado = cb_OP.Text;

            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            if (cb_Area.SelectedIndex == 0 || cb_Area.SelectedIndex == 1)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"No_box_hr\" FROM public.\"Deshidratado\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["No_box_hr"].ToString();
                    Txt_meta.Text = resultado;

                    if (!string.IsNullOrEmpty(Txt_Read_1.Text))
                    {
                        calcular_meta_programada();
                    }
                }
            }
            if (cb_Area.SelectedIndex == 1)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Kg_seco_hr\" FROM public.\"Deshidratado\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Kg_seco_hr"].ToString();
                    Txt_meta.Text = resultado;

                    if (!string.IsNullOrEmpty(Txt_Read_1.Text))
                    {
                        calcular_meta_programada();
                    }
                }
            }
            if (cb_Area.SelectedIndex == 2)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_kg_hr\" FROM public.\"Evaporado\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_kg_hr"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 3)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_Kg_hr\" FROM public.\"Grind\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_Kg_hr"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Kg_person_hr\" FROM public.\"Inspeccion\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Kg_person_hr"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 6)
            {
                // Consulta para buscar donde OP = valor_buscado
                // Obtiene la fecha actual
                DateTime hoy = DateTime.Now;

                // Define las fechas de inicio y fin del rango (solo cambia el año actual)
                DateTime inicio = new DateTime(hoy.Year, 5, 1);   // 1 de mayo
                DateTime fin = new DateTime(hoy.Year, 9, 30);     // 30 de septiembre

                string query;

                // Valida si la fecha actual está dentro del rango
                if (hoy >= inicio && hoy <= fin)
                {
                    query = "SELECT \"Meta_kg_hr_hum\" FROM public.\"Polvos\" WHERE \"OP\" = @valorBuscado;";
                }
                else
                {
                    query = "SELECT \"Meta_kg_hr_idon\" FROM public.\"Polvos\" WHERE \"OP\" = @valorBuscado;";
                }


                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {

                    if (hoy >= inicio && hoy <= fin)
                    {
                        resultado = dt.Rows[0]["Meta_kg_hr_hum"].ToString();
                    }
                    else
                    {
                        resultado = dt.Rows[0]["Meta_kg_hr_idon"].ToString();
                    }
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 7)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_Kg_Hr\" FROM public.\"Revolturas\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_Kg_Hr"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 8)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_Kg_Hr\" FROM public.\"Maquinas\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_Kg_Hr"].ToString();
                    Txt_meta.Text = resultado;
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
            catch
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
            catch
            {
                errorProvider2.SetError(Mask_txt_hr2, "Formato de hora inválido. Asegúrese de usar HH:mm.");
                btn_save_ficha.Enabled = false;
            }
        }

        private void txt_Tiempo_energia_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
        }

        private void Txt_2_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_2_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (RadTextBox)sender;
            char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void Txt_3_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_4_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_5_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_6_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_7_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_8_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_9_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir control de borrado (Backspace)
            if (e.KeyChar == (char)Keys.Back)
            {
                return;
            }

            // Permitir solo un punto decimal
            if (e.KeyChar == '.')
            {
                if ((sender as RadTextBox).Text.Contains("."))
                {
                    e.Handled = true; // Ya hay un punto → se bloquea
                }
                return;
            }

            // Permitir solo números
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Bloquear todo lo que no sea número
            }
        }

        private void Txt_10_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir control de borrado (Backspace)
            if (e.KeyChar == (char)Keys.Back)
            {
                return;
            }

            // Permitir solo un punto decimal
            if (e.KeyChar == '.')
            {
                if ((sender as RadTextBox).Text.Contains("."))
                {
                    e.Handled = true; // Ya hay un punto → se bloquea
                }
                return;
            }

            // Permitir solo números
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Bloquear todo lo que no sea número
            }
        }

        private void porcentaje_logrado_planeacion()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            //double meta_kg = Convert.ToDouble(Txt_Read_2.Text);
            //double Kg_Prod_Terminado = Convert.ToDouble(Txt_2.Text);
            //double kg_fuera_espec = Convert.ToDouble(Txt_3.Text);
            if (double.TryParse(Txt_Read_2.Text, out double meta_kg) && double.TryParse(Txt_2.Text, out double Kg_Prod_Terminado) && double.TryParse(Txt_3.Text, out double kg_fuera_espec))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (meta_kg == 0)
                {
                    Txt_Read_4.Text = "";
                    return;
                }

                // Calcular la fórmula: ((Q2 - S2) / P2)
                double resultado = (Kg_Prod_Terminado - kg_fuera_espec) / meta_kg;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_4.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }
        private void porcentaje_logrado_planeacion_platinum()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            //double meta_kg = Convert.ToDouble(Txt_Read_2.Text);
            //double Kg_Prod_Terminado = Convert.ToDouble(Txt_2.Text);
            if (double.TryParse(Txt_Read_2.Text, out double meta_kg) && double.TryParse(Txt_2.Text, out double Kg_Prod_Terminado))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (meta_kg == 0)
                {
                    Txt_Read_4.Text = "";
                    return;
                }

                double resultado = Kg_Prod_Terminado / meta_kg;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_4.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }
        private void calcular_kg_entrada_proceso()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta para buscar donde OP = valor_buscado
            string query = "SELECT \"Kg_pz\" FROM public.\"Maquinas\" WHERE \"OP\" = @valorBuscado;";
            string valorBuscado = cb_OP.Text;
            string kg_pz_string;
            decimal kg_pz, pz_producidas, resultado;
            kg_pz = 0;

            // Crear parámetro
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
            };

            // Ejecutar consulta
            System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

            // Verificar si se encontraron resultados
            if (dt != null && dt.Rows.Count > 0)
            {
                kg_pz_string = dt.Rows[0]["Kg_pz"].ToString();
                kg_pz = Convert.ToDecimal(kg_pz_string);
            }
            pz_producidas = Convert.ToDecimal(Txt_1.Text);
            resultado = pz_producidas * kg_pz;

            Txt_Read_5.Text = resultado.ToString();
        }

        private void porcentaje_aumento_humedad()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            //double Kg_entrada = Convert.ToDouble(Txt_1.Text);
            //double Kg_Prod_Terminado = Convert.ToDouble(Txt_2.Text);
            if (double.TryParse(Txt_1.Text, out double Kg_entrada) && double.TryParse(Txt_2.Text, out double Kg_Prod_Terminado))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (Kg_entrada == 0)
                {
                    Txt_Read_5.Text = "";
                    return;
                }

                // Calcular la fórmula: ((Q2 - S2) / P2)
                double resultado = (Kg_Prod_Terminado / Kg_entrada) - 1;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_5.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }

        private void porcentaje_cumplimiento_metas()
        {
            if (double.TryParse(Txt_2.Text, out double kg_prod_seco) && double.TryParse(Txt_4.Text, out double kg_fuera_espec) && double.TryParse(Txt_Read_2.Text, out double kg_secos_meta))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (kg_secos_meta == 0)
                {
                    Txt_Read_5.Text = "";
                    return;
                }

                // Calcular la fórmula: ((Q2 - S2) / P2)
                double resultado = (kg_prod_seco - kg_fuera_espec) / kg_secos_meta;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_5.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }

        private void Ftt_metodo()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            //double kg_prod_seco = Convert.ToDouble(Txt_2.Text);
            //double kg_fuera_espec = Convert.ToDouble(Txt_4.Text);
            if (double.TryParse(Txt_2.Text, out double kg_prod_seco) && double.TryParse(Txt_4.Text, out double kg_fuera_espec))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (kg_prod_seco == 0)
                {
                    Txt_Read_8.Text = "";
                    return;
                }

                double resultado = (kg_prod_seco - kg_fuera_espec) / kg_prod_seco;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_8.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }

        private void Relacion_Fresco_seco()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            //double kg_frescos_enter_sec = Convert.ToDouble(Txt_Read_4.Text);
            //double kg_prod_seco = Convert.ToDouble(Txt_2.Text);
            if (double.TryParse(Txt_Read_4.Text, out double kg_frescos_enter_sec) && double.TryParse(Txt_2.Text, out double kg_prod_seco))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (kg_prod_seco == 0)
                {
                    Txt_Read_7.Text = "";
                    return;
                }

                double resultado = kg_frescos_enter_sec / kg_prod_seco;

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_7.Text = resultado.ToString("0.##"); // Formato de porcentaje con 2 decimales
            }
        }
        private void cb_lote_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string valorBuscado = cb_lote.Text;

            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            if (cb_lote.SelectedIndex != -1)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"OP\", \"kg_frescos_enter_se\" FROM public.\"Ficha\" WHERE \"Lote\" = @valorBuscado and \"Area\" = 'Tunel/Sumergidor';";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string OP = string.Empty;
                string kgFrescoEnterSecador = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    OP = dt.Rows[0]["OP"].ToString();
                    kgFrescoEnterSecador = dt.Rows[0]["kg_frescos_enter_se"].ToString();
                    cb_OP.Enabled = true;
                    cb_OP.Text = OP;
                    cb_OP.Enabled = false;

                    cb_Turno.Enabled = true;

                    Txt_Read_4.Text = kgFrescoEnterSecador;
                    buscar_Meta_hr();
                }
            }
            else
            {
                cb_Turno.SelectedIndex = -1;
                cb_Turno.Enabled = false;
            }
        }

        private void Txt_11_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir control de borrado (Backspace)
            if (e.KeyChar == (char)Keys.Back)
            {
                return;
            }

            // Permitir solo un punto decimal
            if (e.KeyChar == '.')
            {
                if ((sender as System.Windows.Forms.TextBox).Text.Contains("."))
                {
                    e.Handled = true; // Ya hay un punto → se bloquea
                }
                return;
            }

            // Permitir solo números
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Bloquear todo lo que no sea número
            }
        }

        private void Txt_1_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7 || cb_Area.SelectedIndex == 8)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                if (!string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) && cb_Area.SelectedIndex == 2)
                {
                    porcentaje_aumento_humedad();
                }
                if (!string.IsNullOrEmpty(Txt_1.Text) && cb_Area.SelectedIndex == 8)
                {
                    calcular_kg_entrada_proceso();
                }
            }
        }

        private void Txt_1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (cb_Area.SelectedIndex == 2)
            {
                var tb = (RadTextBox)sender;
                char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

                // Permitir backspace
                if (e.KeyChar == '\b') return;

                // Permitir un solo punto decimal
                if (e.KeyChar == decimalSep)
                {
                    if (tb.Text.Contains(decimalSep)) e.Handled = true;
                    return;
                }

                // Permitir solo dígitos
                if (!char.IsDigit(e.KeyChar))
                    e.Handled = true;
            }
        }

        private void Txt_1_Validating(object sender, CancelEventArgs e)
        {
            if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7)
            {
                RadTextBox textBox = (RadTextBox)sender;
                string text = textBox.Text;

                // Regex que acepta "" o dígitos con un punto decimal opcional
                if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true; // mantiene el foco en el control
                }
            }
        }

        private void txt_TM_mecanico_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (MaterialTextBox)sender;
            char decimalSep = '.';

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void txt_Tiempo_energia_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (MaterialTextBox)sender;
            char decimalSep = '.';

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void dgv_metas_polvos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 6;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_polvos = dgv_metas_polvos.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_polvos.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_polvos.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_2.Text = dgv_metas_polvos.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_3.Text = dgv_metas_polvos.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_revolturas_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 5;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_revolturas = dgv_metas_revolturas.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_maquinas_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 7;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_maquinas = dgv_metas_maquinas.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_5.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[6].Value.ToString();

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                txt_Meta_5.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void btn_cancelar_ficha_Click(object sender, EventArgs e)
        {
            cb_Area.SelectedIndex = -1;
            reiniciarCampos();
            cb_Area.Focus();
        }

        private void editarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            editar = true;
            Editar Editar_ficha = new Editar(connectionString, editar);
            // Suscripción al evento con los dos parámetros
            Editar_ficha.FichaSeleccionada += (id_global, area) =>
            {
                if (id_global == null || area == null)
                {
                    editar = false;
                    return;
                }
                //MessageBox.Show($"ID seleccionado: {id_global}\nÁrea: {area}");
                id_global_ficha = id_global;
                if (area == "Tunel/ Sumergidor")
                {
                    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                    // Consulta para buscar donde OP = valor_buscado
                    string query2 = "SELECT \"OP\", \"Turno\", \"Fecha\", \"MetaHr\", \"Hr_inicio\", \"Hr_fin\", \"Lote\", \"Kg_enter_proceso\", \"Merma_canica\", " +
                    "\"Merma_podrido\", \"Merma_tina\", \"Merma_piso\", \"Merma_canaletas\", \"Merma_lavado_bandas\", \"Personal_Operativo\", \"Cascara_carrete\", " +
                    "\"Merma_Tunel\" FROM public.\"Ficha\" WHERE \"ID_Ficha\" = @valorBuscado;";

                    // Crear parámetro
                    NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id_global) }
                    };

                    // Ejecutar consulta
                    System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                    // Variables
                    string OP = string.Empty;
                    string Turno = string.Empty;
                    DateTime Fecha = DateTime.MinValue;
                    string Hr_inicio = string.Empty;
                    string Hr_fin = string.Empty;
                    string MetaHr = string.Empty;
                    string Lote = string.Empty;
                    string Kg_enter_proceso = string.Empty;
                    string Merma_canica = string.Empty;
                    string Merma_podrido = string.Empty;
                    string Merma_tina = string.Empty;
                    string Merma_piso = string.Empty;
                    string Merma_canaletas = string.Empty;
                    string Merma_lavado_bandas = string.Empty;
                    string Personal_Operativo = string.Empty;
                    string Cascara_carrete = string.Empty;
                    string Merma_Tunel = string.Empty;

                    // Verificar si se encontraron resultados
                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        OP = dt2.Rows[0]["OP"].ToString();
                        Turno = dt2.Rows[0]["Turno"].ToString();
                        Fecha = Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                        Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                        Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                        MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                        Lote = dt2.Rows[0]["Lote"].ToString();
                        Kg_enter_proceso = dt2.Rows[0]["Kg_enter_proceso"].ToString();
                        Merma_canica = dt2.Rows[0]["Merma_canica"].ToString();
                        Merma_podrido = dt2.Rows[0]["Merma_podrido"].ToString();
                        Merma_tina = dt2.Rows[0]["Merma_tina"].ToString();
                        Merma_piso = dt2.Rows[0]["Merma_piso"].ToString();
                        Merma_canaletas = dt2.Rows[0]["Merma_canaletas"].ToString();
                        Merma_lavado_bandas = dt2.Rows[0]["Merma_lavado_bandas"].ToString();
                        Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                        Cascara_carrete = dt2.Rows[0]["Cascara_carrete"].ToString();
                        Merma_Tunel = dt2.Rows[0]["Merma_Tunel"].ToString();
                    }

                    cb_Area.Text = area;
                    cb_OP.Text = OP;
                    // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                    cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                    cb_OP.Focus();

                    cb_Turno.Text = Turno;
                    cb_Turno.Focus();
                    dtp1.Value = Fecha;
                    Mask_txt_hr1.Text = Hr_inicio;
                    Mask_txt_hr2.Text = Hr_fin;
                    Txt_meta.Text = MetaHr;
                    Txt_1.Text = Lote;
                    Txt_2.Text = Kg_enter_proceso;
                    Txt_3.Text = Merma_canica;
                    Txt_4.Text = Merma_podrido;
                    Txt_5.Text = Merma_tina;
                    Txt_6.Text = Merma_piso;
                    Txt_7.Text = Merma_canaletas;
                    Txt_8.Text = Merma_lavado_bandas;
                    Txt_9.Text = Personal_Operativo;
                    Txt_10.Text = Cascara_carrete;
                    Txt_11.Text = Merma_Tunel;
                    // Deshabilitar controles no editables
                    cb_Area.Enabled = false;
                    Txt_1.Enabled = false;

                    actualiza_tiempos(id_global);
                }
                if (area == "Despegue")
                {
                    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                    // Consulta para buscar donde OP = valor_buscado
                    string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""Lote"", ""Kg_prod_seco"", ""Merma_kg"", ""Kg_fuera_espec"", ""Kg_resecar"", 
                                    ""Personal_Operativo"", ""Hr_programadas"", ""Hr_efectivas"", ""porcent_cump_meta"", ""Kg_meta"", ""Relacion_Fr_seco"", 
                                    ""FTT"", ""Hr_inicio"", ""Hr_fin"", ""MetaHr"" FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";

                    // Crear parámetro
                    NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id_global) }
                    };

                    // Ejecutar consulta
                    System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);

                    // Variables
                    string OP = string.Empty;
                    string Turno = string.Empty;
                    DateTime Fecha = DateTime.MinValue;
                    string Hr_inicio = string.Empty;
                    string Hr_fin = string.Empty;
                    string MetaHr = string.Empty;
                    string Lote = string.Empty;
                    string Kg_prod_seco = string.Empty;
                    string Merma_kg = string.Empty;
                    string Kg_fuera_espec = string.Empty;
                    string Kg_resecar = string.Empty;
                    string Personal_Operativo = string.Empty;

                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        OP = dt2.Rows[0]["OP"].ToString();
                        Turno = dt2.Rows[0]["Turno"].ToString();
                        Fecha = Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                        Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                        Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                        MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                        Lote = dt2.Rows[0]["Lote"].ToString();
                        Kg_prod_seco = dt2.Rows[0]["Kg_prod_seco"].ToString();
                        Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                        Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                        Kg_resecar = dt2.Rows[0]["Kg_resecar"].ToString();
                        Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                    }

                    cb_Area.Text = area;
                    cb_OP.Text = OP;
                    // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                    cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                    cb_OP.Focus();
                    cb_Turno.Text = Turno;
                    cb_Turno.Focus();
                    dtp1.Value = Fecha;
                    Mask_txt_hr1.Text = Hr_inicio;
                    Mask_txt_hr2.Text = Hr_fin;
                    Txt_meta.Text = MetaHr;
                    cb_lote.Text = Lote;
                    cb_lote.Focus();
                    cb_lote_SelectionChangeCommitted(cb_lote, EventArgs.Empty);
                    Txt_2.Text = Kg_prod_seco;
                    Txt_3.Text = Merma_kg;
                    Txt_4.Text = Kg_fuera_espec;
                    Txt_5.Text = Kg_resecar;
                    Txt_6.Text = Personal_Operativo;

                    // Deshabilitar controles no editables
                    cb_Area.Enabled = false;
                    cb_lote.Enabled = false;

                    actualiza_tiempos(id_global);
                }
                if (area == "Evaporado" || area == "Grind" || area == "Empacado" || area == "Revolturas")
                {
                    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                    string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""Kg_enter_proceso"", ""Kg_prod_term"", 
                    ""Kg_fuera_espec"", ""Merma_kg"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", 
                    ""MetaHr"" FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";
                    // Crear parámetro
                    NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id_global) }
                    };

                    // Ejecutar consulta
                    System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);

                    // Variables
                    string OP = string.Empty;
                    string Turno = string.Empty;
                    DateTime Fecha = DateTime.MinValue;
                    string Hr_inicio = string.Empty;
                    string Hr_fin = string.Empty;
                    string MetaHr = string.Empty;
                    string Kg_meta = string.Empty;
                    string Kg_enter_proceso = string.Empty;
                    string Kg_prod_term = string.Empty;
                    string Kg_fuera_espec = string.Empty;
                    string Merma_kg = string.Empty;
                    string Personal_Operativo = string.Empty;

                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        OP = dt2.Rows[0]["OP"].ToString();
                        Turno = dt2.Rows[0]["Turno"].ToString();
                        Fecha = Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                        Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                        Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                        MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                        Kg_meta = dt2.Rows[0]["Kg_meta"].ToString();
                        Kg_enter_proceso = dt2.Rows[0]["Kg_enter_proceso"].ToString();
                        Kg_prod_term = dt2.Rows[0]["Kg_prod_term"].ToString();
                        Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                        Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                        Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                    }

                    cb_Area.Text = area;
                    cb_OP.Text = OP;
                    // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                    cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                    cb_OP.Focus();
                    cb_Turno.Text = Turno;
                    cb_Turno.Focus();
                    dtp1.Value = Fecha;
                    Mask_txt_hr1.Text = Hr_inicio;
                    Mask_txt_hr2.Text = Hr_fin;
                    Txt_meta.Text = MetaHr;
                    Txt_1.Text = Kg_enter_proceso;
                    Txt_2.Text = Kg_prod_term;
                    Txt_3.Text = Kg_fuera_espec;
                    Txt_4.Text = Merma_kg;
                    Txt_5.Text = Personal_Operativo;
                    // Deshabilitar controles no editables
                    cb_Area.Enabled = false;

                    actualiza_tiempos(id_global);
                }
                if (area == "Inspeccion")
                {
                    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                    string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""Proceso"", ""Kg_prod_term"", ""Kg_enter_proceso"", 
                    ""Kg_fuera_espec"", ""Merma_kg"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", 
                    ""MetaHr"" FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";
                    // Crear parámetro
                    NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id_global) }
                    };
                    // Ejecutar consulta
                    System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                    // Variables
                    string OP = string.Empty;
                    string Turno = string.Empty;
                    DateTime Fecha = DateTime.MinValue;
                    string Hr_inicio = string.Empty;
                    string Hr_fin = string.Empty;
                    string MetaHr = string.Empty;
                    string Proceso = string.Empty;
                    string Kg_prod_term = string.Empty;
                    string Kg_enter_proceso = string.Empty;
                    string Kg_fuera_espec = string.Empty;
                    string Merma_kg = string.Empty;
                    string Personal_Operativo = string.Empty;
                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        OP = dt2.Rows[0]["OP"].ToString();
                        Turno = dt2.Rows[0]["Turno"].ToString();
                        Fecha = Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                        Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                        Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                        MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                        Proceso = dt2.Rows[0]["Proceso"].ToString();
                        Kg_prod_term = dt2.Rows[0]["Kg_prod_term"].ToString();
                        Kg_enter_proceso = dt2.Rows[0]["Kg_enter_proceso"].ToString();
                        Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                        Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                        Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                    }
                    cb_Area.Text = area;
                    cb_OP.Text = OP;
                    // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                    cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                    cb_OP.Focus();
                    cb_Turno.Text = Turno;
                    cb_Turno.Focus();
                    dtp1.Value = Fecha;
                    Mask_txt_hr1.Text = Hr_inicio;
                    Mask_txt_hr2.Text = Hr_fin;
                    Txt_meta.Text = MetaHr;
                    Txt_1.Text = Kg_enter_proceso;
                    Txt_2.Text = Kg_prod_term;
                    Txt_3.Text = Kg_fuera_espec;
                    Txt_4.Text = Merma_kg;
                    Txt_5.Text = Personal_Operativo;
                    cb_proceso.Text = Proceso;
                    cb_proceso.Focus();
                    // Deshabilitar controles no editables
                    cb_Area.Enabled = false;
                    actualiza_tiempos(id_global);
                }
                if (area == "Polvos")
                {
                    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                    string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""Kg_enter_proceso"", ""Kg_prod_term"",  
                    ""Kg_fuera_espec"", ""Merma_kg"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", 
                    ""MetaHr"", ""Polvo_colector"", ""Granulo"" FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";
                    // Crear parámetro
                    NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id_global) }
                    };
                    // Ejecutar consulta
                    System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                    // Variables
                    string OP = string.Empty;
                    string Turno = string.Empty;
                    DateTime Fecha = DateTime.MinValue;
                    string Hr_inicio = string.Empty;
                    string Hr_fin = string.Empty;
                    string MetaHr = string.Empty;
                    string Kg_prod_term = string.Empty;
                    string Kg_enter_proceso = string.Empty;
                    string Kg_fuera_espec = string.Empty;
                    string Merma_kg = string.Empty;
                    string Personal_Operativo = string.Empty;
                    string Polvo_colector = string.Empty;
                    string Granulo = string.Empty;
                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        OP = dt2.Rows[0]["OP"].ToString();
                        Turno = dt2.Rows[0]["Turno"].ToString();
                        Fecha = Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                        Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                        Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                        MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                        Kg_prod_term = dt2.Rows[0]["Kg_prod_term"].ToString();
                        Kg_enter_proceso = dt2.Rows[0]["Kg_enter_proceso"].ToString();
                        Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                        Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                        Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                        Polvo_colector = dt2.Rows[0]["Polvo_colector"].ToString();
                        Granulo = dt2.Rows[0]["Granulo"].ToString();
                    }
                    cb_Area.Text = area;
                    cb_OP.Text = OP;
                    // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                    cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                    cb_OP.Focus();
                    cb_Turno.Text = Turno;
                    cb_Turno.Focus();
                    dtp1.Value = Fecha;
                    Mask_txt_hr1.Text = Hr_inicio;
                    Mask_txt_hr2.Text = Hr_fin;
                    Txt_meta.Text = MetaHr;
                    Txt_1.Text = Kg_enter_proceso;
                    Txt_2.Text = Kg_prod_term;
                    Txt_3.Text = Kg_fuera_espec;
                    Txt_4.Text = Merma_kg;
                    Txt_5.Text = Personal_Operativo;
                    cb_proceso.Focus();
                    Txt_6.Text = Polvo_colector;
                    Txt_7.Text = Granulo;
                    // Deshabilitar controles no editables
                    cb_Area.Enabled = false;
                    actualiza_tiempos(id_global);
                }
                if (area == "Máquinas")
                {
                    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                    string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""MetaHr"", ""Hr_inicio"", ""Hr_fin"", ""Pz_prod"", ""Kg_meta"", ""Kg_prod_term"", 
                    ""Kg_fuera_espec"", ""Merma_kg"", ""Personal_Operativo"", ""Bobina_kg_enter"", ""Bobina_utilizada"", ""Bobina_merma"" 
                    FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";
                    // Crear parámetro
                    NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id_global) }
                    };

                    // Ejecutar consulta
                    System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);

                    // Variables
                    string OP = string.Empty;
                    string Turno = string.Empty;
                    DateTime Fecha = DateTime.MinValue;
                    string Hr_inicio = string.Empty;
                    string Hr_fin = string.Empty;
                    string MetaHr = string.Empty;
                    string Kg_meta = string.Empty;
                    string Kg_prod_term = string.Empty;
                    string Kg_fuera_espec = string.Empty;
                    string Merma_kg = string.Empty;
                    string Personal_Operativo = string.Empty;
                    string Pz_prod = string.Empty;
                    string Bobina_kg_enter = string.Empty;
                    string Bobina_utilizada = string.Empty;
                    string Bobina_merma = string.Empty;

                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        Fecha = Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                        Turno = dt2.Rows[0]["Turno"].ToString();
                        OP = dt2.Rows[0]["OP"].ToString();
                        MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                        Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                        Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                        Pz_prod = dt2.Rows[0]["Pz_prod"].ToString();
                        Kg_meta = dt2.Rows[0]["Kg_meta"].ToString();
                        Kg_prod_term = dt2.Rows[0]["Kg_prod_term"].ToString();
                        Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                        Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                        Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                        Bobina_kg_enter = dt2.Rows[0]["Bobina_kg_enter"].ToString();
                        Bobina_utilizada = dt2.Rows[0]["Bobina_utilizada"].ToString();
                        Bobina_merma = dt2.Rows[0]["Bobina_merma"].ToString();
                    }

                    cb_Area.Text = area;
                    cb_OP.Text = OP;
                    // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                    cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                    cb_OP.Focus();
                    cb_Turno.Text = Turno;
                    cb_Turno.Focus();
                    dtp1.Value = Fecha;
                    Mask_txt_hr1.Text = Hr_inicio;
                    Mask_txt_hr2.Text = Hr_fin;
                    Txt_meta.Text = MetaHr;
                    Txt_1.Text = Pz_prod;
                    Txt_2.Text = Kg_prod_term;
                    Txt_3.Text = Kg_fuera_espec;
                    Txt_4.Text = Merma_kg;
                    Txt_5.Text = Personal_Operativo;
                    Txt_6.Text = Bobina_kg_enter;
                    Txt_7.Text = Bobina_utilizada;
                    Txt_8.Text = Bobina_merma;
                    // Deshabilitar controles no editables
                    cb_Area.Enabled = false;

                    actualiza_tiempos(id_global);
                }

            };

            Editar_ficha.Show();
        }

        private void actualiza_tiempos(string id)
        {
            try
            {
                // Cargar tiempos muertos mecánicos
                actualiza_tiempos_mecanicos(id);

                // Cargar tiempos muertos operativos
                actualiza_tiempos_operativos(id);

                // Cargar tiempos muertos comida (si es necesario)
                actualiza_tiempo_comida(id);

                // Cargar tiempos muertos energía (si es necesario)
                actualiza_tiempo_energia(id);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar tiempos muertos: {ex.Message}");
            }
        }

        private void actualiza_tiempos_mecanicos(string id)
        {
            dgv_mecanico.DataSource = null;
            dgv_mecanico.Columns.Clear();

            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT ""Min_Detenidos"" as ""Minutos Detenidos"", 
                            ""Motivos"" as ""Motivos Mecánicos"" 
                     FROM public.""Tiempo_muerto_Mecanico"" 
                     WHERE ""ID_Ficha"" = @id_ficha 
                     ORDER BY ""ID_T_Mecanico"" ASC;";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id) }
            };

            dbHelper.LoadDataIntoDataGridView(query, dgv_mecanico, parameters);
            // Configurar columnas DESPUÉS de cargar los datos
            ConfigurarColumnasDataGridView(dgv_mecanico);
        }

        private void actualiza_tiempos_operativos(string id)
        {
            dgv_operativo.DataSource = null;
            dgv_operativo.Columns.Clear();

            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT ""Min_Detenidos"" as ""Minutos Detenidos"", 
                            ""Motivos"" as ""Motivos Operativos"" 
                     FROM public.""Tiempo_Muerto_Operativo"" 
                     WHERE ""ID_Ficha"" = @id_ficha 
                     ORDER BY ""ID_T_Operativo"" ASC;";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id) }
            };

            dbHelper.LoadDataIntoDataGridView(query, dgv_operativo, parameters);
            // Configurar columnas DESPUÉS de cargar los datos
            ConfigurarColumnasDataGridView(dgv_operativo);
        }

        private void actualiza_tiempo_comida(string id)
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT ""Minutos_Detenidos"" 
                     FROM public.""Tiempo_Muerto_Comida"" 
                     WHERE ""ID_Ficha"" = @id_ficha;";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id) }
            };

            System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

            if (dt != null && dt.Rows.Count > 0)
            {
                txt_Tiempo_comida.Text = dt.Rows[0]["Minutos_Detenidos"].ToString();
            }
            else
            {
                txt_Tiempo_comida.Text = "0";
            }
        }

        private void actualiza_tiempo_energia(string id)
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT ""Minutos_Detenidos"" 
                     FROM public.""Tiempo_Muerto_Energia"" 
                     WHERE ""ID_Ficha"" = @id_ficha;";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = Convert.ToInt32(id) }
            };

            System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

            if (dt != null && dt.Rows.Count > 0)
            {
                txt_Tiempo_energia.Text = dt.Rows[0]["Minutos_Detenidos"].ToString();
            }
            else
            {
                txt_Tiempo_energia.Text = "0";
            }
        }

        private void ConfigurarColumnasDataGridView(DataGridView dataGridView)
        {
            if (dataGridView.Columns.Count >= 2)
            {
                // Primera columna: AutoSizeMode = AllCells
                dataGridView.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                //dataGridView.Columns[0].HeaderText = "Minutos Detenidos";

                // Segunda columna: AutoSizeMode = Fill
                dataGridView.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                //dataGridView.Columns[1].HeaderText = "Motivos";

                // Opcional: Configuración adicional para mejor apariencia
                dataGridView.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dataGridView.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                dataGridView.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void dtp_calidad_ValueChanged(object sender, EventArgs e)
        {
            ActualizarNumeroSemana();
        }
        // Método para actualizar el número de semana
        private void ActualizarNumeroSemana()
        {
            try
            {
                DateTime fechaSeleccionada = dtp_polvos.Value;

                // Obtener el número de semana según el calendario gregoriano
                CultureInfo cultura = CultureInfo.CurrentCulture;
                Calendar calendario = cultura.Calendar;

                // Calcular el número de semana
                int numeroSemana = calendario.GetWeekOfYear(
                    fechaSeleccionada,
                    cultura.DateTimeFormat.CalendarWeekRule,
                    cultura.DateTimeFormat.FirstDayOfWeek
                );

                // Mostrar en el TextBox
                txt_no_semana_polvos.Text = numeroSemana.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al calcular el número de semana: {ex.Message}");
            }
        }

        private void txt_merma_calidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                btn_cancel_polvos.Enabled = true;
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                btn_cancel_polvos.Enabled = true;
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
                btn_cancel_polvos.Enabled = true;
                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void btn_save_calidad_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_merma_polvos.Text))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                DateTime fecha = dtp_polvos.Value;
                if (!string.IsNullOrEmpty(id_global_polvos_calidad))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idCalidad = Convert.ToInt32(id_global_polvos_calidad);

                    queryInsertUpdate = "UPDATE public.\"Limpieza_polvos\" SET \"Fecha\" = @Fecha, \"Kg_merma\" = @Kg_merma WHERE \"ID_Limpieza\" = @ID_Limpieza;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date)
                        {
                            Value = fecha
                        },
                        new NpgsqlParameter("@Kg_merma", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_merma_polvos.Text)
                        },
                        new NpgsqlParameter("@ID_Limpieza", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idCalidad  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Limpieza_polvos\" (\"Fecha\", \"Kg_merma\") VALUES (@Fecha, @Kg_merma);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                        new NpgsqlParameter("@Kg_merma", Convert.ToDecimal(txt_merma_polvos.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_polvos_calidad();
                    txt_merma_polvos.Text = string.Empty;
                    txt_no_semana_polvos.Text = string.Empty;
                    btn_cancel_polvos.Enabled = false;
                    btn_edit_polvos.Enabled = false;
                    id_global_polvos_calidad = string.Empty;
                    btn_save_polvos.Enabled = false;
                    txt_merma_polvos.Enabled = false;
                    dtp_polvos.Enabled = false;
                }
            }
        }

        private void dgv_metas_polvos_calidad_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DateTime Fecha = DateTime.MinValue;
            if (e.RowIndex >= 0)
            {
                id_global_polvos_calidad = dgv_polvos_calidad.Rows[e.RowIndex].Cells[0].Value.ToString();
                Fecha = Convert.ToDateTime(dgv_polvos_calidad.Rows[e.RowIndex].Cells[1].Value);
                txt_merma_polvos.Text = dgv_polvos_calidad.Rows[e.RowIndex].Cells[3].Value.ToString();

                dtp_polvos.Value = Fecha;
                btn_edit_polvos.Enabled = true;
                btn_delete_polvos.Enabled = true;
                txt_merma_polvos.Enabled = false;
                dtp_polvos.Enabled = false;
                btn_save_polvos.Enabled = false;
                btn_cancel_polvos.Enabled = false;
            }
        }

        private void btn_new_calidad_Click(object sender, EventArgs e)
        {
            //habilitar controles
            dtp_polvos.Enabled = true;
            txt_merma_polvos.Enabled = true;
            btn_cancel_polvos.Enabled = true;
            btn_save_polvos.Enabled = true;
            btn_edit_polvos.Enabled = false;
            btn_delete_polvos.Enabled = false;

            //limpiar campos
            txt_merma_polvos.Text = string.Empty;
            dtp_polvos.Value = DateTime.Now;

            //limpiar variables globales
            id_global_polvos_calidad = string.Empty;

            //enfocar
            txt_merma_polvos.Focus();
        }

        private void btn_edit_calidad_Click(object sender, EventArgs e)
        {
            btn_edit_polvos.Enabled = false;
            btn_save_polvos.Enabled = true;
            btn_cancel_polvos.Enabled = true;
            btn_delete_polvos.Enabled = false;
            txt_merma_polvos.Enabled = true;
            txt_merma_polvos.Focus();
        }

        private void btn_cancel_calidad_Click(object sender, EventArgs e)
        {
            btn_save_polvos.Enabled = false;
            btn_cancel_polvos.Enabled = false;
            btn_delete_polvos.Enabled = false;
            dtp_polvos.Value = DateTime.Now;
            txt_merma_polvos.Text = string.Empty;
            btn_edit_polvos.Enabled = false;
            id_global_polvos_calidad = string.Empty;
            txt_merma_polvos.Enabled = false;
            dtp_polvos.Enabled = false;
        }

        private void btn_delete_calidad_Click(object sender, EventArgs e)
        {
            if (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar esta Merma de la tabla?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id_calidad = Convert.ToInt32(id_global_polvos_calidad);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Limpieza_polvos""
                       WHERE ""ID_Limpieza"" = @id";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_calidad
                    }
                };
                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {

                    actualiza_polvos_calidad();
                    txt_merma_polvos.Text = string.Empty;
                    txt_no_semana_polvos.Text = string.Empty;
                    btn_cancel_polvos.Enabled = false;
                    btn_edit_polvos.Enabled = false;
                    id_global_polvos_calidad = string.Empty;
                    btn_save_polvos.Enabled = false;
                    btn_delete_polvos.Enabled = false;
                    txt_merma_polvos.Enabled = false;
                    dtp_polvos.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar este registro", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btn_new_dt_op_Click(object sender, EventArgs e)
        {
            //habilitar controles
            txt_orden_produc.Enabled = true;
            txt_producto.Enabled = true;
            txt_medida.Enabled = true;
            txt_descripcion.Enabled = true;
            txt_especificacion.Enabled = true;
            txt_ingredientes.Enabled = true;
            txt_humedad.Enabled = true;
            cb_Comercio.Enabled = true;
            cb_Manzana.Enabled = true;
            txt_analisis.Enabled = true;
            txt_area_proceso.Enabled = true;
            txt_op_origen.Enabled = true;
            txt_destino.Enabled = true;
            txt_clacificacion.Enabled = true;
            btn_cancel_dt_op.Enabled = true;
            btn_save_dt_op.Enabled = true;
            btn_edit_dt_op.Enabled = false;
            btn_delete_dt_op.Enabled = false;

            //limpiar campos
            cb_Comercio.SelectedIndex = -1;
            cb_Manzana.SelectedIndex = -1;
            txt_orden_produc.Text = string.Empty;
            txt_producto.Text = string.Empty;
            txt_medida.Text = string.Empty;
            txt_descripcion.Text = string.Empty;
            txt_especificacion.Text = string.Empty;
            txt_ingredientes.Text = string.Empty;
            txt_humedad.Text = string.Empty;
            txt_analisis.Text = string.Empty;
            txt_area_proceso.Text = string.Empty;
            txt_op_origen.Text = string.Empty;
            txt_destino.Text = string.Empty;
            txt_clacificacion.Text = string.Empty;

            //limpiar variables globales
            id_global_detalles_OP = string.Empty;

            //enfocar controles
            cb_Comercio.Focus();
            cb_Manzana.Focus();
            txt_orden_produc.Focus();
        }

        private void btn_save_dt_op_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_orden_produc.Text) || string.IsNullOrEmpty(txt_producto.Text) || string.IsNullOrEmpty(txt_medida.Text) ||
                string.IsNullOrEmpty(txt_descripcion.Text) || cb_Comercio.SelectedIndex == -1 || string.IsNullOrEmpty(txt_clacificacion.Text) || 
                string.IsNullOrEmpty(txt_destino.Text))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                // Verificar si ya existe
                DatabaseHelper dbHelperop = new DatabaseHelper(connectionString);
                string queryChecop = "SELECT COUNT(*) FROM public.\"Detalles_OP\" WHERE \"Orden_Produccion\"  ILIKE @OP;";
                NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@OP", txt_orden_produc.Text.ToUpper().Trim()),
                };
                System.Data.DataTable dtCheckop = dbHelperop.ExecuteSelectQuery(queryChecop, parametersCheck);
                if (dtCheckop != null && dtCheckop.Rows.Count > 0 && Convert.ToInt32(dtCheckop.Rows[0][0]) > 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_detalles_OP))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int id = Convert.ToInt32(id_global_detalles_OP);

                    queryInsertUpdate = "UPDATE public.\"Detalles_OP\" SET \"Orden_Produccion\" = @Orden_Produccion, \"Producto\" = @Producto, \"Medida\" = @Medida, \"Descripcion\" = @Descripcion, \"Especificacion\" = @Especificacion, \"Ingredientes\" = @Ingredientes, \"Humedad\" = @Humedad, \"Comercio\" = @Comercio, \"Manzana\" = @Manzana, \"Analisis\" = @Analisis, \"Area_Proceso\" = @Area_Proceso, \"OP_Origen\" = @OP_Origen, \"Destino\" = @Destino, \"Clasificacion\" = @Clasificacion WHERE \"ID_Dt_OP\" = @ID;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Orden_Produccion", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_orden_produc.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Producto", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_producto.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Medida", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_medida.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Descripcion", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_descripcion.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Especificacion", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_especificacion.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Ingredientes", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_ingredientes.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Humedad", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_humedad.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Comercio", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = cb_Comercio.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Manzana", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = cb_Manzana.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Analisis", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_analisis.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Area_Proceso", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_area_proceso.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@OP_Origen", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op_origen.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Destino", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_destino.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Clasificacion", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_clacificacion.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@ID", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = id  // variable convertida a int
                        }
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Detalles_OP\" (\"Orden_Produccion\", \"Producto\", \"Medida\", \"Descripcion\", \"Especificacion\", \"Ingredientes\", \"Humedad\", \"Comercio\", \"Manzana\", \"Analisis\", \"Area_Proceso\", \"OP_Origen\", \"Destino\", \"Clasificacion\") VALUES (@Orden_Produccion, @Producto, @Medida, @Descripcion, @Especificacion, @Ingredientes, @Humedad, @Comercio, @Manzana, @Analisis, @Area_Proceso, @OP_Origen, @Destino, @Clasificacion);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Orden_Produccion", txt_orden_produc.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Producto", txt_producto.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Medida", txt_medida.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Descripcion", txt_descripcion.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Especificacion", txt_especificacion.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Ingredientes", txt_ingredientes.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Humedad", txt_humedad.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Comercio", cb_Comercio.Text),
                        new NpgsqlParameter("@Manzana", cb_Manzana.Text),
                        new NpgsqlParameter("@Analisis", txt_analisis.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Area_Proceso", txt_area_proceso.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@OP_Origen", txt_op_origen.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Destino", txt_destino.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Clasificacion", txt_clacificacion.Text.ToUpper().Trim())
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    // Actualizar DataGridView
                    actualiza_detalles_OP();
                    limpiarCampos_detalles_OP();

                    // Deshabilitar controles después de guardar
                    btn_save_dt_op.Enabled = false;
                    btn_cancel_dt_op.Enabled = false;
                    txt_orden_produc.Enabled = false;
                    txt_producto.Enabled = false;
                    txt_medida.Enabled = false;
                    txt_descripcion.Enabled = false;
                    txt_especificacion.Enabled = false;
                    txt_ingredientes.Enabled = false;
                    txt_humedad.Enabled = false;
                    txt_analisis.Enabled = false;
                    txt_area_proceso.Enabled = false;
                    txt_op_origen.Enabled = false;
                    txt_destino.Enabled = false;
                    txt_clacificacion.Enabled = false;
                    btn_edit_dt_op.Enabled = false;
                    id_global_detalles_OP = string.Empty;

                    // Deshabilitar y enfocar comboboxes
                    cb_Comercio.Enabled = true;
                    cb_Comercio.Focus();
                    cb_Comercio.Enabled = false;
                    cb_Manzana.Enabled = true;
                    cb_Manzana.Focus();
                    cb_Manzana.Enabled = false;
                }
            }
        }

        private void dgv_detalles_op_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (nivel_user == "Administrador") 
            {
                if (e.RowIndex >= 0)
                {
                    id_global_detalles_OP = dgv_detalles_op.Rows[e.RowIndex].Cells[0].Value.ToString();
                    txt_orden_produc.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[1].Value.ToString();
                    txt_producto.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[2].Value.ToString();
                    txt_medida.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[3].Value.ToString();
                    txt_descripcion.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[4].Value.ToString();
                    txt_especificacion.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[5].Value.ToString();
                    txt_ingredientes.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[6].Value.ToString();
                    txt_humedad.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[7].Value.ToString();
                    cb_Comercio.Enabled = true;
                    cb_Comercio.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[8].Value.ToString();
                    cb_Comercio.Enabled = false;
                    cb_Manzana.Enabled = true;
                    cb_Manzana.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[9].Value.ToString();
                    cb_Manzana.Enabled = false;
                    txt_analisis.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[10].Value.ToString();
                    txt_area_proceso.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[11].Value.ToString();
                    txt_op_origen.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[12].Value.ToString();
                    txt_destino.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[13].Value.ToString();
                    txt_clacificacion.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[14].Value.ToString();

                    btn_edit_dt_op.Enabled = true;
                    btn_delete_dt_op.Enabled = true;

                    txt_orden_produc.Enabled = false;
                    txt_producto.Enabled = false;
                    txt_medida.Enabled = false;
                    txt_descripcion.Enabled = false;
                    txt_especificacion.Enabled = false;
                    txt_ingredientes.Enabled = false;
                    txt_humedad.Enabled = false;
                    cb_Comercio.Enabled = false;
                    cb_Manzana.Enabled = false;
                    txt_analisis.Enabled = false;
                    txt_area_proceso.Enabled = false;
                    txt_op_origen.Enabled = false;
                    txt_destino.Enabled = false;
                    txt_clacificacion.Enabled = false;

                    btn_save_dt_op.Enabled = false;
                    btn_cancel_dt_op.Enabled = false;
                }
            } 
        }

        private void btn_cancel_dt_op_Click(object sender, EventArgs e)
        {
            btn_save_dt_op.Enabled = false;
            btn_cancel_dt_op.Enabled = false;
            btn_delete_dt_op.Enabled = false;
            limpiarCampos_detalles_OP();
            txt_orden_produc.Enabled = false;
            txt_producto.Enabled = false;
            txt_medida.Enabled = false;
            txt_descripcion.Enabled = false;
            txt_especificacion.Enabled = false;
            txt_ingredientes.Enabled = false;
            txt_humedad.Enabled = false;
            cb_Comercio.Enabled = false;
            cb_Manzana.Enabled = false;
            txt_analisis.Enabled = false;
            txt_area_proceso.Enabled = false;
            txt_op_origen.Enabled = false;
            txt_destino.Enabled = false;
            txt_clacificacion.Enabled = false;
            btn_edit_dt_op.Enabled = false;
            id_global_detalles_OP = string.Empty;
        }

        private void btn_edit_dt_op_Click(object sender, EventArgs e)
        {
            btn_save_dt_op.Enabled = true;
            btn_cancel_dt_op.Enabled = true;
            btn_delete_dt_op.Enabled = false;
            txt_orden_produc.Enabled = true;
            txt_producto.Enabled = true;
            txt_medida.Enabled = true;
            txt_descripcion.Enabled = true;
            txt_especificacion.Enabled = true;
            txt_ingredientes.Enabled = true;
            txt_humedad.Enabled = true;
            cb_Comercio.Enabled = true;
            cb_Manzana.Enabled = true;
            txt_analisis.Enabled = true;
            txt_area_proceso.Enabled = true;
            txt_op_origen.Enabled = true;
            txt_destino.Enabled = true;
            txt_clacificacion.Enabled = true;
            txt_orden_produc.Focus();
        }

        private void btn_delete_dt_op_Click(object sender, EventArgs e)
        {
            if (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar esta Orden de Producción?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id_OP = Convert.ToInt32(id_global_detalles_OP);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Detalles_OP""
                       WHERE ""ID_Dt_OP"" = @idOP";

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
                    actualiza_detalles_OP();
                    limpiarCampos_detalles_OP();
                    btn_save_dt_op.Enabled = false;
                    btn_cancel_dt_op.Enabled = false;
                    btn_delete_dt_op.Enabled = false;
                    txt_orden_produc.Enabled = false;
                    txt_producto.Enabled = false;
                    txt_medida.Enabled = false;
                    txt_descripcion.Enabled = false;
                    txt_especificacion.Enabled = false;
                    txt_ingredientes.Enabled = false;
                    btn_edit_dt_op.Enabled = false;
                    txt_analisis.Enabled = false;
                    txt_area_proceso.Enabled = false;
                    txt_op_origen.Enabled = false;
                    txt_destino.Enabled = false;
                    txt_clacificacion.Enabled = false;
                    id_global_detalles_OP = string.Empty;

                    cb_Comercio.Enabled = true;
                    cb_Comercio.Focus();
                    cb_Comercio.Enabled = false;
                    cb_Manzana.Enabled = true;
                    cb_Manzana.Focus();
                    cb_Manzana.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar la Orden de Producción", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btn_buscar_detalles_Click(object sender, EventArgs e)
        {
            // Si el filtro está activo, limpiar filtro y campos de búsqueda
            if (filtrodetallesOP)
            {
                // Mostrar todas las filas del DataGridView
                foreach (DataGridViewRow row in dgv_detalles_op.Rows)
                {
                    row.Visible = true;
                }

                // Restablecer el estado del filtro
                filtrodetallesOP = false;
            }

            string OP = txt_buscar_dt_op.Text.Trim().ToUpper();
            //validar que el campo OP este lleno
            if (string.IsNullOrEmpty(OP))
            {
                MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Deseleccionar cualquier celda y fila antes de filtrar
            dgv_detalles_op.ClearSelection();
            dgv_detalles_op.CurrentCell = null;

            // Mover el CurrencyManager a una posición válida
            CurrencyManager cm = (CurrencyManager)BindingContext[dgv_detalles_op.DataSource];
            if (cm != null && cm.Count > 0)
                cm.Position = -1;

            bool hayFiltro = false;
            foreach (DataGridViewRow row in dgv_detalles_op.Rows)
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
                MetroFramework.MetroMessageBox.Show(this, "No se encontró OP",
                                                    "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_limpiar_detalles_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgv_detalles_op.Rows)
            {
                row.Visible = true;
            }

            // Limpiar los campos de búsqueda
            txt_buscar_dt_op.Text = string.Empty;
            txt_buscar_dt_op.Focus();

            // Restablecer el estado del filtro
            filtrodetallesOP = false;
        }

        private void Txt_3_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (RadTextBox)sender;
            char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void Txt_4_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (RadTextBox)sender;
            char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void Txt_5_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (RadTextBox)sender;
            char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void dtp_tunel_ValueChanged(object sender, EventArgs e)
        {
            ActualizarNumeroSemana_tunel();
        }
        private void ActualizarNumeroSemana_tunel()
        {
            try
            {
                DateTime fechaSeleccionada = dtp_polvos.Value;

                // Obtener el número de semana según el calendario gregoriano
                CultureInfo cultura = CultureInfo.CurrentCulture;
                Calendar calendario = cultura.Calendar;

                // Calcular el número de semana
                int numeroSemana = calendario.GetWeekOfYear(
                    fechaSeleccionada,
                    cultura.DateTimeFormat.CalendarWeekRule,
                    cultura.DateTimeFormat.FirstDayOfWeek
                );

                // Mostrar en el TextBox
                txt_no_semana_tunel.Text = numeroSemana.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al calcular el número de semana: {ex.Message}");
            }
        }

        private void txt_merma_tunel_KeyPress(object sender, KeyPressEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                btn_cancel_polvos.Enabled = true;
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                btn_cancel_polvos.Enabled = true;
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
                btn_cancel_polvos.Enabled = true;
                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void btn_new_tunel_Click(object sender, EventArgs e)
        {
            //habilitar controles
            dtp_tunel.Enabled = true;
            txt_merma_tunel.Enabled = true;
            btn_cancel_tunel.Enabled = true;
            btn_save_tunel.Enabled = true;
            btn_edit_tunel.Enabled = false;
            btn_delete_tunel.Enabled = false;

            //limpiar campos
            txt_merma_tunel.Text = string.Empty;
            dtp_tunel.Value = DateTime.Now;

            //limpiar variables globales
            id_global_tunel_calidad = string.Empty;

            //enfocar
            txt_merma_tunel.Focus();
        }

        private void btn_save_tunel_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_merma_tunel.Text))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                DateTime fecha = dtp_tunel.Value;
                if (!string.IsNullOrEmpty(id_global_tunel_calidad))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idCalidad = Convert.ToInt32(id_global_tunel_calidad);

                    queryInsertUpdate = "UPDATE public.\"Limpieza_tunel\" SET \"Fecha\" = @Fecha, \"Kg_merma\" = @Kg_merma WHERE \"ID_Tunel\" = @ID_Limpieza;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date)
                        {
                            Value = fecha
                        },
                        new NpgsqlParameter("@Kg_merma", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = Convert.ToDecimal(txt_merma_tunel.Text)
                        },
                        new NpgsqlParameter("@ID_Limpieza", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idCalidad  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Limpieza_tunel\" (\"Fecha\", \"Kg_merma\") VALUES (@Fecha, @Kg_merma);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                        new NpgsqlParameter("@Kg_merma", Convert.ToDecimal(txt_merma_tunel.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_tunel_calidad();
                    txt_merma_tunel.Text = string.Empty;
                    txt_no_semana_tunel.Text = string.Empty;
                    btn_cancel_tunel.Enabled = false;
                    btn_edit_tunel.Enabled = false;
                    id_global_tunel_calidad = string.Empty;
                    btn_save_tunel.Enabled = false;
                    txt_merma_tunel.Enabled = false;
                    dtp_tunel.Enabled = false;
                }
            }
        }

        private void btn_edit_tunel_Click(object sender, EventArgs e)
        {
            btn_edit_tunel.Enabled = false;
            btn_save_tunel.Enabled = true;
            btn_cancel_tunel.Enabled = true;
            btn_delete_tunel.Enabled = false;
            txt_merma_tunel.Enabled = true;
            txt_merma_tunel.Focus();
        }

        private void btn_delete_tunel_Click(object sender, EventArgs e)
        {
            if (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar esta Merma de la tabla?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id_calidad = Convert.ToInt32(id_global_tunel_calidad);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Limpieza_tunel""
                       WHERE ""ID_Tunel"" = @id";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_calidad
                    }
                };
                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {

                    actualiza_tunel_calidad();
                    txt_merma_tunel.Text = string.Empty;
                    txt_no_semana_tunel.Text = string.Empty;
                    btn_cancel_tunel.Enabled = false;
                    btn_edit_tunel.Enabled = false;
                    id_global_tunel_calidad = string.Empty;
                    btn_save_tunel.Enabled = false;
                    btn_delete_tunel.Enabled = false;
                    txt_merma_tunel.Enabled = false;
                    dtp_tunel.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar este registro", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btn_cancel_tunel_Click(object sender, EventArgs e)
        {
            btn_save_tunel.Enabled = false;
            btn_cancel_tunel.Enabled = false;
            btn_delete_tunel.Enabled = false;
            dtp_tunel.Value = DateTime.Now;
            txt_merma_tunel.Text = string.Empty;
            btn_edit_tunel.Enabled = false;
            id_global_tunel_calidad = string.Empty;
            txt_merma_tunel.Enabled = false;
            dtp_tunel.Enabled = false;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(tabControl1.SelectedIndex == 0) 
            {
                tab_limpiezas.SelectedIndex = 0;
            }
            if (tabControl1.SelectedIndex == 1)
            {
                tab_limpiezas.SelectedIndex = 1;
            }
        }

        private void tab_limpiezas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tab_limpiezas.SelectedIndex == 0)
            {
                tabControl1.SelectedIndex = 0;
            }
            if (tab_limpiezas.SelectedIndex == 1)
            {
                tabControl1.SelectedIndex = 1;
            }
        }

        private void dgv_Tunel_calidad_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DateTime Fecha = DateTime.MinValue;
            if (e.RowIndex >= 0)
            {
                id_global_tunel_calidad = dgv_Tunel_calidad.Rows[e.RowIndex].Cells[0].Value.ToString();
                Fecha = Convert.ToDateTime(dgv_Tunel_calidad.Rows[e.RowIndex].Cells[1].Value);
                txt_merma_tunel.Text = dgv_Tunel_calidad.Rows[e.RowIndex].Cells[3].Value.ToString();

                dtp_tunel.Value = Fecha;
                btn_edit_tunel.Enabled = true;
                btn_delete_tunel.Enabled = true;
                txt_merma_tunel.Enabled = false;
                dtp_tunel.Enabled = false;
                btn_save_tunel.Enabled = false;
                btn_cancel_tunel.Enabled = false;
            }
        }

        private void btn_new_report_merma_Click(object sender, EventArgs e)
        {
            actualiza_reporte_merma();
            btn_clean_merma.Enabled = true;
            btn_export_excel_merma.Enabled = true;
        }

        private void btn_clean_merma_Click(object sender, EventArgs e)
        {
            dgv_reporte_merma.DataSource = null;   // Desvincula cualquier origen de datos
            dgv_reporte_merma.Rows.Clear();        // Limpia todas las filas (por si no hay DataSource)
            dgv_reporte_merma.Columns.Clear();     // Limpia todas las columnas
            btn_export_excel_merma.Enabled = false;
            btn_clean_merma.Enabled = false;
        }

        private async void btn_export_excel_merma_Click(object sender, EventArgs e)
        {
            LoadingScreen.ShowLoading();
            try
            {
                // Mostrar el diálogo de guardar archivo en el hilo principal
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls",
                    Title = "Guardar archivo de Excel"
                };

                string filePath = null;

                // Mostrar el diálogo de guardado en el hilo principal
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = saveFileDialog.FileName;

                    // Ejecutar la tarea pesada (exportar a Excel) en un hilo separado usando Task.Run
                    await Task.Run(() =>
                    {
                        ExportarDataGridViewFiltradoAExcel(dgv_reporte_merma, filePath);
                    });

                    MetroFramework.MetroMessageBox.Show(this, "La exportación fue completada con éxito.", "Exportación a Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, "Error durante la exportación: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Ocultar la pantalla de cargando
                LoadingScreen.HideLoading();
            }
        }

        private void ExportarDataGridViewFiltradoAExcel(DataGridView dgv, string filePath)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            int colIndex = 1;

            // Exportar encabezados solo para las columnas visibles
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                if (dgv.Columns[i].Visible)
                {
                    worksheet.Cells[1, colIndex] = dgv.Columns[i].HeaderText;

                    Excel.Range headerCell = worksheet.Cells[1, colIndex];
                    headerCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    headerCell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    headerCell.Font.Bold = true;
                    headerCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    colIndex++;
                }
            }

            int rowIndex = 2;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Visible)
                {
                    colIndex = 1;
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        if (dgv.Columns[j].Visible)
                        {
                            worksheet.Cells[rowIndex, colIndex] = row.Cells[j].Value?.ToString();
                            colIndex++;
                        }
                    }
                    rowIndex++;
                }
            }

            Excel.Range usedRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowIndex - 1, colIndex - 1]];
            usedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            worksheet.Columns.AutoFit();

            if (filePath.EndsWith(".xlsx"))
            {
                workbook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            else if (filePath.EndsWith(".xls"))
            {
                workbook.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
            }

            workbook.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
       

        private void btn_new_report_consolidado_Click(object sender, EventArgs e)
        {
            reporte_consolidado();
            btn_export_excel_consolidado.Enabled = true;
            btn_clean_consolidado.Enabled = true;
        }
        private void reporte_consolidado()
        {
            string var1 = cb_area_reporte.Text;
            string querySimple= string.Empty;
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            if (var1 == "Tunel/Sumergidor")
            {
                querySimple = @"WITH turnos_trabajados AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") AS semana,
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        COUNT(*) AS total_turnos_trabajados
    FROM public.""Ficha""
    WHERE ""Area"" = 'Tunel/Sumergidor'
    GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
),
merma_semanal AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") AS semana,
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        SUM(""Kg_merma"") AS merma_total_semanal
    FROM public.""Limpieza_tunel""
    GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
)
SELECT 
    f.""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    f.""Turno"",
    f.""Lote"",
    f.""OP"",
    -- ... resto de columnas igual ...
    CASE 
        WHEN tt.total_turnos_trabajados > 0 AND ms.merma_total_semanal IS NOT NULL
        THEN ROUND(ms.merma_total_semanal / tt.total_turnos_trabajados, 2)
        ELSE 0 
    END AS ""Limpieza Túnel""
FROM public.""Ficha"" f
LEFT JOIN turnos_trabajados tt 
    ON EXTRACT(WEEK FROM f.""Fecha"") = tt.semana 
    AND EXTRACT(YEAR FROM f.""Fecha"") = tt.año
LEFT JOIN merma_semanal ms 
    ON EXTRACT(WEEK FROM f.""Fecha"") = ms.semana 
    AND EXTRACT(YEAR FROM f.""Fecha"") = ms.año
WHERE f.""Area"" = 'Tunel/Sumergidor'
ORDER BY f.""OP"" ASC;";
            }


            if (var1 == "Despegue")
            {
                querySimple = @"
                SELECT 
                    ""Fecha"",
                    EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
                    ""Turno"",
                    ""Lote"",
                    ""OP"",
                    ""kg_frescos_enter_se"" AS ""Kg Frescos Entrada a Secador"",
                    ""porcent_cump_meta"" AS ""% Cumplimiento a Metas"",
                    ""Kg_prod_seco"" AS ""Kilos Producto Seco"",
                    ""Merma_kg"" AS ""Merma(Kg)"",
                    ""Kg_fuera_espec"" AS ""Kg Fuera de Especificación"",
                    ""Kg_resecar"" AS ""Kg para Resecar"",
                    ""Relacion_Fr_seco"" AS ""Relación Fresco-Seco"",
                    ""FTT""
                FROM public.""Ficha""
                WHERE ""Area"" = @Area
                ORDER BY ""OP"" ASC;";
            }

            if (var1 == "Evaporado")
            {
                querySimple = @"
                SELECT 
                    ""Fecha"",
                    EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
                    ""Turno"",
                    ""OP"",
                    ""Kg_meta"" as ""Meta(Kg)"",
                    ""porcent_cump_meta"" AS ""% Cumplimiento a Metas"",
                    ""Kg_enter_proceso"" AS ""Kg Entrada(Proceso)"",
                    ""Kg_prod_term"" as ""Kg Producto Terminado"",
                    ""Kg_fuera_espec"" AS ""Kg Fuera de Especificación"",
                    ""Merma_kg"" AS ""Merma(Kg)"",
                    ""porcent_aumento_hum"" as ""% Aumento de Humedad""
                FROM public.""Ficha""
                WHERE ""Area"" = @Area
                ORDER BY ""OP"" ASC;";
            }

            if (var1 == "Grind" || var1 == "Inspeccion" || var1 == "Empacado")
            {
                querySimple = @"
                SELECT 
                    ""Fecha"",
                    EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
                    ""Turno"",
                    ""OP"",
                    ""Kg_meta"" as ""Meta(Kg)"",
                    ""porcent_cump_meta"" AS ""% Cumplimiento a Metas"",
                    ""Kg_enter_proceso"" AS ""Kg Entrada(Proceso)"",
                    ""Kg_prod_term"" as ""Kg Producto Terminado"",
                    ""Kg_fuera_espec"" AS ""Kg Fuera de Especificación"",
                    ""Merma_kg"" AS ""Merma(Kg)""
                FROM public.""Ficha""
                WHERE ""Area"" = @Area
                ORDER BY ""OP"" ASC;";
            }

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value }
                    };
            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridViewTelerik(querySimple, rgv_reporte_consolidado, parameters);

            // Configurar el DataGridView
            rgv_reporte_consolidado.Columns["Fecha"].FormatString = "{0:dd/MM/yyyy}";

            // 🔹 Formato de porcentaje para las columnas de tipo decimal
            if (rgv_reporte_consolidado.Columns.Contains("% Cumplimiento a Metas"))
            {
                var colMeta = rgv_reporte_consolidado.Columns["% Cumplimiento a Metas"];
                colMeta.DataType = typeof(decimal);
                colMeta.FormatString = "{0:P0}"; // ejemplo: 0.85 → 85%
                colMeta.TextAlignment = ContentAlignment.MiddleRight;
            }

            if (rgv_reporte_consolidado.Columns.Contains("FTT"))
            {
                var colFTT = rgv_reporte_consolidado.Columns["FTT"];
                colFTT.DataType = typeof(decimal);
                colFTT.FormatString = "{0:P0}";
                colFTT.TextAlignment = ContentAlignment.MiddleRight;
            }

            if (rgv_reporte_consolidado.Columns.Contains("% Aumento de Humedad"))
            {
                var colFTT = rgv_reporte_consolidado.Columns["% Aumento de Humedad"];
                colFTT.DataType = typeof(decimal);
                colFTT.FormatString = "{0:P0}";
                colFTT.TextAlignment = ContentAlignment.MiddleRight;
            }

            // Limpiar el filtro al cargar nuevos datos
            txt_filtro_report_consolidado.Clear();
        }

        private async void btn_export_excel_consolidado_Click(object sender, EventArgs e)
        {
            LoadingScreen.ShowLoading();
            try
            {
                // Mostrar el diálogo de guardar archivo en el hilo principal
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls",
                    Title = "Guardar archivo de Excel"
                };

                string filePath = null;

                // Mostrar el diálogo de guardado en el hilo principal
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = saveFileDialog.FileName;

                    // Ejecutar la tarea pesada (exportar a Excel) en un hilo separado usando Task.Run
                    await Task.Run(() =>
                    {
                       ExportarRadGridViewFiltradoAExcel(rgv_reporte_consolidado, filePath);
                    });

                    MetroFramework.MetroMessageBox.Show(this, "La exportación fue completada con éxito.", "Exportación a Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, "Error durante la exportación: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Ocultar la pantalla de cargando
                LoadingScreen.HideLoading();
            }
        }
        private void ExportarRadGridViewFiltradoAExcel(RadGridView radGridView, string filePath)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            int colIndex = 1;

            // Exportar encabezados solo para las columnas visibles
            foreach (GridViewDataColumn column in radGridView.Columns)
            {
                if (column.IsVisible)
                {
                    worksheet.Cells[1, colIndex] = column.HeaderText;

                    Excel.Range headerCell = worksheet.Cells[1, colIndex];
                    headerCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                    headerCell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                    headerCell.Font.Bold = true;
                    headerCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    colIndex++;
                }
            }

            int rowIndex = 2;

            // Exportar datos de las filas visibles
            foreach (GridViewRowInfo row in radGridView.Rows)
            {
                if (row.IsVisible)
                {
                    colIndex = 1;
                    foreach (GridViewDataColumn column in radGridView.Columns)
                    {
                        if (column.IsVisible)
                        {
                            var cellValue = row.Cells[column.Name].Value;
                            worksheet.Cells[rowIndex, colIndex] = cellValue?.ToString();
                            colIndex++;
                        }
                    }
                    rowIndex++;
                }
            }

            Excel.Range usedRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowIndex - 1, colIndex - 1]];
            usedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            worksheet.Columns.AutoFit();

            if (filePath.EndsWith(".xlsx"))
            {
                workbook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            else if (filePath.EndsWith(".xls"))
            {
                workbook.SaveAs(filePath, Excel.XlFileFormat.xlExcel8);
            }

            workbook.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        private void btn_filtro_consolidado_Click(object sender, EventArgs e)
        {
            txt_filtro_report_consolidado.Clear();
        }

        private void rgv_reporte_consolidado_CustomFiltering(object sender, GridViewCustomFilteringEventArgs e)
        {
            string textoFiltro = txt_filtro_report_consolidado.Text.Trim();

            // Si no hay texto de filtro, mostrar todas las filas
            if (string.IsNullOrEmpty(textoFiltro))
            {
                e.Visible = true;
                ResetearEstiloCeldas(e);
                return;
            }

            // Iniciar actualización para mejor rendimiento
            rgv_reporte_consolidado.BeginUpdate();

            // Por defecto ocultar la fila
            e.Visible = false;

            // Buscar en todas las celdas de la fila
            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                GridViewCellInfo celda = e.Row.Cells[i];

                // Verificar si la celda tiene valor
                if (celda.Value != null)
                {
                    string textoCelda = celda.Value.ToString();

                    // Buscar coincidencia (case-insensitive)
                    if (textoCelda.IndexOf(textoFiltro, 0, StringComparison.InvariantCultureIgnoreCase) >= 0)
                    {
                        e.Visible = true; // Mostrar la fila si hay coincidencia

                        // Resaltar la celda que coincide
                        celda.Style.CustomizeFill = true;
                        celda.Style.DrawFill = true;
                        celda.Style.BackColor = Color.FromArgb(201, 252, 254); // Color azul claro
                    }
                    else
                    {
                        // Resetear estilo si no coincide
                        celda.Style.Reset();
                    }
                }
                else
                {
                    // Resetear estilo si la celda es nula
                    celda.Style.Reset();
                }
            }

            rgv_reporte_consolidado.EndUpdate(false);
        }
        // Método auxiliar para resetear el estilo de todas las celdas
        private void ResetearEstiloCeldas(GridViewCustomFilteringEventArgs e)
        {
            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                e.Row.Cells[i].Style.Reset();
            }
        }

        private void txt_filtro_report_consolidado_TextChanged(object sender, EventArgs e)
        {
            // Refrescar el grid para aplicar el filtro
            if (rgv_reporte_consolidado != null)
            {
                rgv_reporte_consolidado.MasterTemplate.Refresh();
            }
        }
    }
}
